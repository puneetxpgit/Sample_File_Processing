"""Request-type classifier fit on the Tower and Ground sheets.

Each run reads the labeled BRIEF SOW rows and builds a TF-IDF nearest-neighbor
model. A new SOW is scored against those rows. The type is returned only when
one type is clearly closer than the others.
"""

from __future__ import annotations

import math
from collections import Counter
from pathlib import Path

from processor.sow_catalog import CatalogExample, SowCatalog, normalize_sow

_CACHE: dict[tuple[str, int], "SowClassifier"] = {}


class SowClassifier:
    def __init__(self, indexes: dict[str, "_CategoryIndex"]):
        self.indexes = indexes

    @classmethod
    def load(cls, path: Path) -> "SowClassifier":
        if not path.exists():
            return cls({})
        key = (str(path.resolve()), path.stat().st_mtime_ns)
        cached = _CACHE.get(key)
        if cached is not None:
            return cached
        catalog = SowCatalog.load(path)
        grouped: dict[str, list[CatalogExample]] = {}
        for example in catalog.examples:
            grouped.setdefault(example.category, []).append(example)
        model = cls({category: _CategoryIndex(examples) for category, examples in grouped.items()})
        _CACHE[key] = model
        return model

    def predict(self, category: str, text: str, allowed: list[str]) -> dict | None:
        index = self.indexes.get(category)
        if index is None:
            return None
        return index.predict(text, allowed)

    def rank_sheets(self, text: str) -> list[dict]:
        ranked = []
        for category, index in self.indexes.items():
            prediction = index.predict(text, allowed=None)
            ranked.append(
                {
                    "category": category,
                    "requestType": (prediction or {}).get("requestType", ""),
                    "similarity": (prediction or {}).get("similarity", 0),
                    "confidence": (prediction or {}).get("confidence", 0),
                }
            )
        ranked.sort(key=lambda item: item["similarity"], reverse=True)
        return ranked


class _CategoryIndex:
    def __init__(self, examples: list[CatalogExample]):
        self.examples = examples
        document_count = len(examples)
        document_frequency: Counter[str] = Counter()
        counts: list[Counter[str]] = []
        for example in examples:
            count = Counter(example.tokens)
            counts.append(count)
            document_frequency.update(count.keys())
        self.idf = {
            token: math.log((1 + document_count) / (1 + frequency)) + 1
            for token, frequency in document_frequency.items()
        }
        self.vectors = [_normalize(_tfidf(count, self.idf)) for count in counts]

    def predict(self, text: str, allowed: list[str] | None) -> dict | None:
        query = _normalize(_tfidf(Counter(normalize_sow(text).split()), self.idf))
        if not query:
            return None
        best_by_type: dict[str, tuple[float, CatalogExample]] = {}
        for example, vector in zip(self.examples, self.vectors):
            if allowed is not None and example.request_type not in allowed:
                continue
            similarity = _dot(query, vector)
            current = best_by_type.get(example.request_type)
            if current is None or similarity > current[0]:
                best_by_type[example.request_type] = (similarity, example)
        if not best_by_type:
            return None
        ordered = sorted(best_by_type.items(), key=lambda item: item[1][0], reverse=True)
        label, (best, example) = ordered[0]
        second = ordered[1][1][0] if len(ordered) > 1 else 0.0
        confidence = _confidence(best, second)
        preview = example.sow if len(example.sow) <= 140 else example.sow[:137] + "..."
        return {
            "requestType": label if confidence >= 90 else "",
            "confidence": confidence,
            "similarity": round(best, 4),
            "stage": "excel_model",
            "reason": f"Nearest {label} row in the {example.category} sheet ({best:.0%} similar): {preview}",
        }


def _tfidf(counts: Counter[str], idf: dict[str, float]) -> dict[str, float]:
    vector = {}
    for token, count in counts.items():
        if len(token) < 2 or token not in idf or count <= 0:
            continue
        vector[token] = (1 + math.log(count)) * idf[token]
    return vector


def _normalize(vector: dict[str, float]) -> dict[str, float]:
    norm = math.sqrt(sum(value * value for value in vector.values()))
    if norm == 0:
        return {}
    return {token: value / norm for token, value in vector.items()}


def _dot(left: dict[str, float], right: dict[str, float]) -> float:
    if len(left) > len(right):
        left, right = right, left
    return sum(value * right.get(token, 0.0) for token, value in left.items())


def _confidence(best: float, second: float) -> int:
    gap = best - second
    if best >= 0.72 and gap >= 0.08:
        return min(99, round(88 + best * 12))
    if best >= 0.58 and gap >= 0.15:
        return min(94, round(best * 100))
    return min(89, round(best * 80))
