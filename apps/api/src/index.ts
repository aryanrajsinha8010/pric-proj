import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import axios from 'axios';
import { ScraperService } from './services/ScraperService';
import { prisma } from './services/PrismaService';
import { ValueScoreService } from './services/ValueScoreService';

dotenv.config();

const app = express();
const port = process.env.PORT || 3001;
const MATCHER_URL = process.env.MATCHER_URL || 'http://localhost:8000';

app.use(cors());
app.use(express.json());

// Health Check
app.get('/health', (req, res) => {
  res.json({ status: 'ok', database: 'connected', timestamp: new Date().toISOString() });
});

// ──────────────────────────────────────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────────────────────────────────────

/** Strip noise-only words but ONLY if enough words remain to form a valid query */
function normalizeQuery(raw: string): string {
  const noise = new Set([
    "price", "buy", "online", "india", "deals", "offers", "latest",
    "best", "top", "comparison", "compare", "vs", "review", "reviews",
    "selling", "cheap", "cheapest", "discount"
  ]);
  const words = raw.trim().split(/\s+/);
  const cleaned = words.filter(w => !noise.has(w.toLowerCase()));
  // Only use cleaned query if it still has at least 2 meaningful words
  return cleaned.length >= 2 ? cleaned.join(' ') : words.join(' ');
}



/** Returns true if the most recent listing was scraped within the last hour */
function isFresh(scrapedAt: Date | null | undefined): boolean {
  if (!scrapedAt) return false;
  return Date.now() - new Date(scrapedAt).getTime() < 60 * 60 * 1000;
}

// ──────────────────────────────────────────────────────────────────────────────
// Suggestions API
// ──────────────────────────────────────────────────────────────────────────────
app.get('/api/v1/suggestions', async (req, res) => {
  const { q } = req.query;
  if (!q || typeof q !== 'string' || q.trim().length < 2) {
    return res.json({ suggestions: [] });
  }
  
  try {
    const products = await prisma.product.findMany({
      where: { canonicalName: { contains: q } },
      select: { canonicalName: true },
      take: 5,
      orderBy: { updatedAt: 'desc' }
    });
    
    let suggestions = products.map(p => p.canonicalName);
    
    // If no db hits, simulate AI intent predictions based on the user's typing
    if (suggestions.length === 0) {
      suggestions = [
        `${q.trim()} latest model`,
        `Best ${q.trim()} under 20000`,
        `${q.trim()} 128GB`,
        `${q.trim()} comparison`
      ];
    }
    res.json({ suggestions: [...new Set(suggestions)].slice(0, 5) });
  } catch (error) {
    res.json({ suggestions: [] });
  }
});

// ──────────────────────────────────────────────────────────────────────────────
// Main Search API
// ──────────────────────────────────────────────────────────────────────────────
app.get('/api/v1/search', async (req, res) => {
  const { q } = req.query;

  if (!q || typeof q !== 'string' || q.trim().length < 2) {
    return res.status(400).json({ error: 'Search query must be at least 2 characters' });
  }

  try {
    const cleanQ = normalizeQuery(q);
    console.log(`[PriceWise] Query: "${q}" → normalized: "${cleanQ}"`);

    // ── Step 1: Cache lookup (split query words for SQLite 'contains') ──
    const queryWords = cleanQ.split(/\s+/).filter(w => w.length > 2);
    const andConditions = queryWords.map(w => ({ canonicalName: { contains: w } }));

    let product = await prisma.product.findFirst({
      where: queryWords.length > 0 ? { AND: andConditions } : { canonicalName: { contains: cleanQ } },
      include: { listings: { include: { priceHistory: true } } },
      orderBy: { updatedAt: 'desc' }
    });

    const cacheIsFresh = isFresh(product?.listings?.[0]?.scrapedAt);

    // ── Step 2: Scrape if no valid/fresh cache ────────────────────────────────
    if (!product || product.listings.length === 0 || !cacheIsFresh) {
      console.log(`[PriceWise] Scraping live data for "${cleanQ}"...`);

      const rawScrapedListings = await ScraperService.scrapeAll(cleanQ);

      if (rawScrapedListings.length === 0) {
        if (product && product.listings.length > 0) {
          console.log(`[PriceWise] Scrape failed for "${cleanQ}", falling back to stale cache.`);
          const scoredListings = ValueScoreService.calculateScores((product.listings ?? []) as any, 'balanced');
          return res.json({ product: { ...product, listings: scoredListings } });
        }
        return res.json({ product: null });
      }

      // ── Step 3: Filter listings vs the SEARCH QUERY (not Amazon result #1) ──
      const queryWordCount = cleanQ.trim().split(/\s+/).length;
      const matchThreshold = queryWordCount === 1 ? 0.10 : queryWordCount === 2 ? 0.18 : 0.25;

      let matched = rawScrapedListings;
      try {
        const batchRes = await axios.post(`${MATCHER_URL}/match_batch`, {
          query: cleanQ,
          titles: rawScrapedListings.map(r => r.title),
          threshold: matchThreshold
        }, { timeout: 5000 });

        const { matches } = batchRes.data;
        matched = rawScrapedListings.filter((raw, idx) => matches[idx]);

        if (matched.length === 0) {
          console.warn('[PriceWise] Matcher filtered all results — using raw listings as fallback');
          matched = rawScrapedListings;
        }
      } catch (err) {
        console.warn('[PriceWise] Matcher batch API failed, using all raw listings.');
      }

      // ── Step 4: Keep top 3 per platform (prevents one platform dominating) ──
      const platformCounts = new Map<string, number>();
      const filteredListings: any[] = [];
      for (const listing of matched) {
        const count = platformCounts.get(listing.platform) ?? 0;
        if (count < 3) {
          filteredListings.push(listing);
          platformCounts.set(listing.platform, count + 1);
        }
      }

      console.log(`[PriceWise] Platforms in results: ${[...platformCounts.entries()].map(([p, n]) => `${p}(${n})`).join(', ')}`);

      // ── Step 5: Pick canonical product name from cheapest matched listing ───
      const canonicalListing = [...filteredListings].sort(
        (a, b) => a.effectivePrice - b.effectivePrice
      )[0] ?? rawScrapedListings[0];

      // ── Step 6: Upsert product record ────────────────────────────────────────
      if (product) {
        product = await prisma.product.update({
          where: { id: product.id },
          data: { 
            updatedAt: new Date(),
            canonicalName: canonicalListing.title // update to latest matched title
          },
          include: { listings: { include: { priceHistory: true } } }
        });
      } else {
        product = await prisma.product.create({
          data: {
            canonicalName: canonicalListing.title,
            brand: 'Generic',
            category: 'Electronics'
          },
          include: { listings: { include: { priceHistory: true } } }
        });
      }

      // ── Step 7: Persist listings (atomic replace) ─────────────────────────────
      const listingsToCreate = filteredListings.map(raw => ({
        productId: product!.id,
        platform: raw.platform,
        mrp: raw.mrp,
        listedPrice: raw.listedPrice,
        effectivePrice: raw.effectivePrice,
        rating: raw.rating,
        reviewCount: raw.reviewCount,
        deliveryDays: raw.deliveryDays,
        productUrl: raw.productUrl,
        imageUrl: raw.imageUrl ?? '',
        hasOffer: raw.hasOffer,
        offerValuePercent: raw.offerValuePercent
      }));

      if (listingsToCreate.length > 0) {
        await prisma.platformListing.deleteMany({ where: { productId: product.id } });
        await prisma.platformListing.createMany({ data: listingsToCreate });
      }

      product = await prisma.product.findUnique({
        where: { id: product.id },
        include: { listings: { include: { priceHistory: true } } }
      });

      if (!product) {
        return res.status(500).json({ error: 'Product record lost after save' });
      }

    } else {
      console.log(`[PriceWise] Cache hit: "${product.canonicalName}" (${product.listings.length} listings)`);
    }

    // ── Step 8: Score and return ───────────────────────────────────────────────
    const scoredListings = ValueScoreService.calculateScores(
      (product?.listings ?? []) as any,
      'balanced'
    );

    res.json({
      product: {
        ...product,
        listings: scoredListings
      }
    });

  } catch (error) {
    console.error('[PriceWise API Error]:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.listen(port, () => {
  console.log(`[PriceWise API]: Server is running at http://localhost:${port}`);
});

process.on('SIGINT', async () => { await prisma.$disconnect(); process.exit(0); });
process.on('SIGTERM', async () => { await prisma.$disconnect(); process.exit(0); });
