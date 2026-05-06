-- RedefineTables
PRAGMA defer_foreign_keys=ON;
PRAGMA foreign_keys=OFF;
CREATE TABLE "new_PlatformListing" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "productId" TEXT NOT NULL,
    "platform" TEXT NOT NULL,
    "mrp" REAL NOT NULL,
    "listedPrice" REAL NOT NULL,
    "effectivePrice" REAL NOT NULL,
    "rating" REAL NOT NULL,
    "reviewCount" INTEGER NOT NULL,
    "deliveryDays" INTEGER NOT NULL,
    "productUrl" TEXT NOT NULL,
    "imageUrl" TEXT NOT NULL DEFAULT '',
    "hasOffer" BOOLEAN NOT NULL DEFAULT false,
    "offerValuePercent" REAL NOT NULL DEFAULT 0,
    "scrapedAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "PlatformListing_productId_fkey" FOREIGN KEY ("productId") REFERENCES "Product" ("id") ON DELETE RESTRICT ON UPDATE CASCADE
);
INSERT INTO "new_PlatformListing" ("deliveryDays", "effectivePrice", "hasOffer", "id", "listedPrice", "mrp", "offerValuePercent", "platform", "productId", "productUrl", "rating", "reviewCount", "scrapedAt") SELECT "deliveryDays", "effectivePrice", "hasOffer", "id", "listedPrice", "mrp", "offerValuePercent", "platform", "productId", "productUrl", "rating", "reviewCount", "scrapedAt" FROM "PlatformListing";
DROP TABLE "PlatformListing";
ALTER TABLE "new_PlatformListing" RENAME TO "PlatformListing";
PRAGMA foreign_keys=ON;
PRAGMA defer_foreign_keys=OFF;
