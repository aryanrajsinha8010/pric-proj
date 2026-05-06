import re
import math
from collections import Counter

class ProductMatcher:
    """
    A lightweight Product Matching engine.
    Uses TF-IDF similarity to determine if two product listings are the same physical item.
    """
    
    def __init__(self):
        self.stop_words = {'a', 'an', 'the', 'is', 'at', 'which', 'on', 'in', 'of', 'for', 'with', 'and'}

    def preprocess(self, text):
        # Lowercase, remove special chars
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

        if not denominator:
            return 0.0
        else:
            return float(numerator) / denominator

    def match(self, title1, title2, threshold=0.5):
        score = self.get_cosine_similarity(title1, title2)
        return score >= threshold, score

if __name__ == "__main__":
    matcher = ProductMatcher()
    
    # Test Cases
    p1 = "Samsung Galaxy S24 Ultra 5G (Titanium Gray, 12GB, 256GB Storage)"
    p2 = "Samsung Galaxy S24 Ultra 5G (Titanium Gray, 256 GB)" # Different platform title
    p3 = "Apple iPhone 15 Pro (128 GB) - Blue Titanium"
    
    match12, score12 = matcher.match(p1, p2)
    match13, score13 = matcher.match(p1, p3)
    
    print(f"Comparing:")
    print(f"1. {p1}")
    print(f"2. {p2}")
    print(f"3. {p3}")
    print("-" * 20)
    print(f"Match 1 & 2: {match12} (Score: {score12:.2f})")
    print(f"Match 1 & 3: {match13} (Score: {score13:.2f})")
