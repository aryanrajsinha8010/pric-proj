export interface PlatformListing {
  platform: string;
  effectivePrice: number;
  rating: number;
  reviewCount: number;
  deliveryDays: number;
  hasOffer: boolean;
  offerValuePercent: number;
}

export interface ValueScoreWeights {
  price: number;
  quality: number;
  offer: number;
  delivery: number;
}

export const SCORING_MODES: Record<string, ValueScoreWeights> = {
  balanced: { price: 0.40, quality: 0.25, offer: 0.20, delivery: 0.15 },
  priceFirst: { price: 0.65, quality: 0.15, offer: 0.15, delivery: 0.05 },
  qualityFirst: { price: 0.20, quality: 0.60, offer: 0.10, delivery: 0.10 },
};

export class ValueScoreService {
  /**
   * Calculates Value Score (0-100) for each listing.
   */
  public static calculateScores(
    listings: PlatformListing[],
    mode: keyof typeof SCORING_MODES = 'balanced'
  ) {
    if (listings.length === 0) return [];

    const weights = SCORING_MODES[mode];
    const prices = listings.map(l => l.effectivePrice);
    const minPrice = Math.min(...prices);
    const maxPrice = Math.max(...prices);
    
    const deliveryTimes = listings.map(l => l.deliveryDays);
    const minDelivery = Math.min(...deliveryTimes);
    const maxDelivery = Math.max(...deliveryTimes);

    return listings.map(listing => {
      // 1. Price Score (Normalized)
      // If all prices are the same, score is 100
      let priceScore = 100;
      if (maxPrice !== minPrice) {
        priceScore = 100 * (maxPrice - listing.effectivePrice) / (maxPrice - minPrice);
      }

      // 2. Quality Score (Rating 1-5 maps to 0-100; clamp to 0 when no rating)
      const clampedRating = Math.max(1, listing.rating || 1);
      const qualityScore = ((clampedRating - 1) / 4) * 100;

      // 3. Review Score (Log-scaled, 10k+ reviews = 100)
      const reviewScore = Math.min(100, (Math.log10(listing.reviewCount + 1) / Math.log10(10001)) * 100);

      // 4. Offer Score
      const offerScore = Math.min(100, listing.offerValuePercent * 2);

      // 5. Delivery Score
      let deliveryScore = 100;
      if (maxDelivery !== minDelivery) {
        deliveryScore = 100 * (maxDelivery - listing.deliveryDays) / (maxDelivery - minDelivery);
      }

      // Composite Calculation
      // Quality is split between rating and review volume (80/20 split)
      const finalQualityScore = (qualityScore * 0.8) + (reviewScore * 0.2);

      const totalScore = 
        (priceScore * weights.price) +
        (finalQualityScore * weights.quality) +
        (offerScore * weights.offer) +
        (deliveryScore * weights.delivery);

      return {
        ...listing,
        valueScore: Math.round(totalScore),
        scoreBreakdown: {
          price: Math.round(priceScore),
          quality: Math.round(finalQualityScore),
          offer: Math.round(offerScore),
          delivery: Math.round(deliveryScore)
        }
      };
    }).sort((a, b) => b.valueScore - a.valueScore);
  }
}
