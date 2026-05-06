-- CreateTable
CREATE TABLE "Product" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "canonicalName" TEXT NOT NULL,
    "brand" TEXT,
    "category" TEXT,
    "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" DATETIME NOT NULL
);

-- CreateTable
CREATE TABLE "PlatformListing" (
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
    "hasOffer" BOOLEAN NOT NULL DEFAULT false,
    "offerValuePercent" REAL NOT NULL DEFAULT 0,
    "scrapedAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "PlatformListing_productId_fkey" FOREIGN KEY ("productId") REFERENCES "Product" ("id") ON DELETE RESTRICT ON UPDATE CASCADE
);

-- CreateTable
CREATE TABLE "PriceHistory" (
    "id" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
    "listingId" TEXT NOT NULL,
    "effectivePrice" REAL NOT NULL,
    "recordedAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "PriceHistory_listingId_fkey" FOREIGN KEY ("listingId") REFERENCES "PlatformListing" ("id") ON DELETE RESTRICT ON UPDATE CASCADE
);
