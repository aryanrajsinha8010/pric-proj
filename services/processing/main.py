from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import List, Tuple
import re
import math
from collections import Counter

app = FastAPI(title="PriceWise Processing Service")

class ProductMatcher:
    def __init__(self):
        self.stop_words = {'a', 'an', 'the', 'is', 'at', 'which', 'on', 'in', 'of', 'for', 'with', 'and'}

    def preprocess(self, text):
        text = text.lower()
        text = re.sub(r'[^a-z0-9 ]+', ' ', text)
        words = text.split()
        return [w for w in words if w not in self.stop_words]

    def get_cosine_similarity(self, text1, text2):
        words1 = self.preprocess(text1)
        words2 = self.preprocess(text2)
        vec1 = Counter(words1)
        vec2 = Counter(words2)
        intersection = set(vec1.keys()) & set(vec2.keys())
        numerator = sum([vec1[x] * vec2[x] for x in intersection])
        sum1 = sum([vec1[x]**2 for x in vec1.keys()])
        sum2 = sum([vec2[x]**2 for x in vec2.keys()])
        denominator = math.sqrt(sum1) * math.sqrt(sum2)
        return float(numerator) / denominator if denominator else 0.0

matcher = ProductMatcher()



class BatchMatchRequest(BaseModel):
    query: str
    titles: List[str]
    threshold: float = 0.65

class BatchMatchResponse(BaseModel):
    matches: List[bool]
    scores: List[float]

@app.post("/match_batch", response_model=BatchMatchResponse)
def match_products_batch(req: BatchMatchRequest):
    matches = []
    scores = []
    for title in req.titles:
        score = matcher.get_cosine_similarity(req.query, title)
        scores.append(score)
        matches.append(score >= req.threshold)
    return BatchMatchResponse(matches=matches, scores=scores)


@app.get("/health")
def health():
    return {"status": "ok"}



if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
