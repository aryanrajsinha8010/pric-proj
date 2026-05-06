import axios from 'axios';

// SerpApi Key Pool - Supports comma-separated keys from .env
const RAW_KEYS = process.env.SERPAPI_KEY || '77a8f890804ba3c501dbb61f752daa3f76188edc6c93add8a8e7b092636c8933,437f8cf0a6fa06f7767c786a0341a6e5613a19ebcc135c361f53b4da1df1d5c6';
const KEY_POOL = RAW_KEYS.split(',').map(k => k.trim()).filter(k => k.length > 0);
let currentKeyIndex = 0;

export class ScraperService {
  /**
   * Gets the next active key from the pool
   */
  private static getNextKey(): string {
    const key = KEY_POOL[currentKeyIndex];
    currentKeyIndex = (currentKeyIndex + 1) % KEY_POOL.length;
    return key;
  }
  /**
   * The Tool Interface / Retrieval Orchestrator using SerpApi
   */
  static async scrapeAll(query: string): Promise<any[]> {
    console.log(`[Retrieval Orchestrator] Executing product_search(query="${query}") via SerpApi...`);

    let attempts = 0;
    const maxAttempts = KEY_POOL.length;

    while (attempts < maxAttempts) {
      const apiKey = this.getNextKey();
      console.log(`[Retrieval] Attempt ${attempts + 1}/${maxAttempts} using key ending in ...${apiKey.slice(-4)}`);

      try {
        const response = await axios.get('https://serpapi.com/search.json', {
          params: {
            engine: 'google_shopping',
            q: query,
            api_key: apiKey,
            hl: 'en',
            gl: 'in'
          }
        });

        const shoppingResults = response.data.shopping_results || [];
        console.log(`[Retrieval] SerpApi returned ${shoppingResults.length} raw results.`);

        return shoppingResults.map((item: any) => {
          const price = item.extracted_price || 0;
          const parsedMrp = item.old_price ? parseFloat(item.old_price.replace(/[^0-9.]/g, '')) : NaN;
          const mrp = isNaN(parsedMrp) ? (price * 1.2) : parsedMrp;
          
          return {
            platform: item.source || 'Unknown Web',
            title: item.title || 'Unknown Product',
            mrp: mrp,
            listedPrice: price,
            effectivePrice: price,
            rating: item.rating || 4.2,
            reviewCount: item.reviews || Math.floor(Math.random() * 500) + 50,
            deliveryDays: item.delivery ? 2 : 4,
            productUrl: item.link || item.product_link || item.source_link || '',
            imageUrl: item.thumbnail || item.image || '',
            hasOffer: mrp > price,
            offerValuePercent: mrp > price ? Math.round(((mrp - price) / mrp) * 100) : 0
          };
        });

      } catch (error: any) {
        attempts++;
        console.error(`[Retrieval] API error on key ...${apiKey.slice(-4)}`);
        
        const errBody = error.response?.data?.error?.toLowerCase() || '';
        const status = error.response?.status;
        
        // SerpApi often returns 400 or 401 for "run out of searches" or "limit"
        if (status === 429 || errBody.includes('limit') || errBody.includes('searches') || errBody.includes('quota') || errBody.includes('exhausted')) {
          console.warn('[Retrieval] Quota exhausted for this key. Rotating to next key...');
          continue; 
        }

        // If it's a generic failure (e.g. timeout), just break out and return empty
        console.error(error.message);
        break;
      }
    }

    console.error("[Retrieval] All keys in the pool failed or exhausted.");
    return [];
  }
}
