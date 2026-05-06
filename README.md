[README.md](https://github.com/user-attachments/files/27445350/README.md)
# PriceWise: Truth & Trust Architecture 🛡️

PriceWise is an AI-driven, multi-platform price comparison engine designed to solve the "trust deficit" in online shopping. Instead of just showing the lowest price, PriceWise uses NLP-powered matching and a visual "Truth Score" to help users find the most reliable deals across the web.

## 🚀 System Architecture

PriceWise is built as a high-performance distributed system:

1.  **Web Frontend (`/web`)**: A premium Next.js 15 application featuring responsive insights, price history visualizations (Recharts), and smooth layout animations (Framer Motion).
2.  **API Gateway (`/apps/api`)**: A robust Node.js middleware that orchestrates search flows, manages the SQLite/Prisma cache, and integrates multiple SerpApi keys for high-capacity retrieval.
3.  **NLP Matcher Service (`/services/processing`)**: A specialized Python microservice (FastAPI) that performs real-time product title matching using TF-IDF and Cosine Similarity to ensure "apples-to-apples" comparisons.

## 🛠️ Tech Stack

- **Frontend**: Next.js, Tailwind CSS, Lucide, Recharts, Framer Motion.
- **Backend**: Node.js, Express, Prisma ORM.
- **Data Layer**: SQLite (Local Cache).
- **Processing**: Python 3.x, FastAPI.
- **APIs**: SerpApi (Google Shopping, Amazon, Flipkart integration).

## 🚦 Getting Started

### Prerequisites
- **Node.js 18+**
- **Python 3.9+**
- **Git**

### Installation

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd pric-proj
   ```

2. **Setup API Environment**:
   Navigate to `apps/api`, copy `.env.example` to `.env`, and add your `SERPAPI_KEY`.

3. **Install Dependencies**:
   ```bash
   # In apps/api
   npm install
   npx prisma generate

   # In web
   npm install

   # In services/processing
   pip install fastapi uvicorn pydantic
   ```

### Manual Launch (3 Terminals)

1. **Terminal 1 (Python Matcher)**:
   ```bash
   cd services/processing
   python main.py
   ```

2. **Terminal 2 (API Gateway)**:
   ```bash
   cd apps/api
   npm run dev
   ```

3. **Terminal 3 (Web UI)**:
   ```bash
   cd web
   npm run dev
   ```

Access the application at `http://localhost:3000`.

## ✨ Core Features

- **Multi-Platform Scraping**: Real-time data from Amazon, Flipkart, and Google Shopping.
- **Smart Matching**: Python-based NLP service filters out noise and irrelevant search results.
- **Value Scoring**: Advanced algorithm weighing Price, Ratings, and Delivery speed.
- **Price History**: Visual trends to help users time their purchases correctly.
- **Truth Header**: A unique UI component that builds trust through transparency.

---
*Built with ❤️ for the future of e-commerce trust.*
