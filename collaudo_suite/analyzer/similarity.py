from __future__ import annotations

from collections import defaultdict
from difflib import SequenceMatcher
from itertools import combinations
from typing import Iterable, Iterator

try:
    from rapidfuzz import fuzz as rapid_fuzz  # type: ignore
except Exception:  # pragma: no cover - fallback for machines without rapidfuzz
    rapid_fuzz = None


class UnionFind:
    def __init__(self, size: int):
        self.parent = list(range(size))
        self.rank = [0] * size

    def find(self, item: int) -> int:
        parent = self.parent[item]
        if parent != item:
            self.parent[item] = self.find(parent)
        return self.parent[item]

    def union(self, a: int, b: int) -> None:
        root_a = self.find(a)
        root_b = self.find(b)
        if root_a == root_b:
            return
        if self.rank[root_a] < self.rank[root_b]:
            self.parent[root_a] = root_b
        elif self.rank[root_a] > self.rank[root_b]:
            self.parent[root_b] = root_a
        else:
            self.parent[root_b] = root_a
            self.rank[root_a] += 1

    def groups(self) -> dict[int, list[int]]:
        out: dict[int, list[int]] = defaultdict(list)
        for idx in range(len(self.parent)):
            out[self.find(idx)].append(idx)
        return dict(out)


def fuzzy_score(s1: str, s2: str) -> float:
    if not s1 or not s2:
        return 0.0
    if rapid_fuzz is not None:
        return float(rapid_fuzz.token_sort_ratio(s1, s2))
    return 100.0 * SequenceMatcher(None, s1, s2).ratio()


def partial_score(s1: str, s2: str) -> float:
    if not s1 or not s2:
        return 0.0
    if rapid_fuzz is not None:
        return float(rapid_fuzz.partial_ratio(s1, s2))
    a = s1.lower()
    b = s2.lower()
    if len(a) > len(b):
        a, b = b, a
    best = 0.0
    window = max(len(a), 1)
    for start in range(0, max(len(b) - window + 1, 1)):
        best = max(best, 100.0 * SequenceMatcher(None, a, b[start : start + window]).ratio())
    return best


def jaccard_score(tokens1: set[str], tokens2: set[str]) -> float:
    if not tokens1 or not tokens2:
        return 0.0
    union = tokens1 | tokens2
    if not union:
        return 0.0
    return 100.0 * len(tokens1 & tokens2) / len(union)


def combined_score(text1: str, text2: str, tokens1: set[str], tokens2: set[str]) -> float:
    return (fuzzy_score(text1, text2) + jaccard_score(tokens1, tokens2)) / 2.0


def calculate_similarity(
    text1: str,
    text2: str,
    tokens1: set[str],
    tokens2: set[str],
    algorithm: str,
) -> float:
    algorithm = (algorithm or "combinato").lower()
    if algorithm == "fuzzy":
        return fuzzy_score(text1, text2)
    if algorithm == "jaccard":
        return jaccard_score(tokens1, tokens2)
    return combined_score(text1, text2, tokens1, tokens2)


def compatible_for_comparison(tokens1: set[str], tokens2: set[str], text1: str, text2: str) -> bool:
    """Cheap prefilter to avoid obviously useless fuzzy comparisons."""
    if not text1 or not text2:
        return False

    len1 = len(text1)
    len2 = len(text2)
    if min(len1, len2) / max(len1, len2) < 0.35:
        return False

    if tokens1 & tokens2:
        return True

    # For very short descriptions, absence of common tokens is not always enough
    # to reject the pair, because spelling errors may alter the only useful word.
    if min(len(tokens1), len(tokens2)) <= 2 and abs(len1 - len2) <= 12:
        return True

    return False


def _record_keys(record) -> set[str]:
    keys: set[str] = set()
    codes = getattr(record, "technical_codes", set())
    tokens = getattr(record, "tokens", set())
    normalized = getattr(record, "normalized_text", "")

    for code in codes:
        keys.add(f"code:{code}")

    useful_tokens = sorted(t for t in tokens if len(t) >= 4 and not t.isdigit())
    for token in useful_tokens[:8]:
        keys.add(f"tok:{token}")

    words = normalized.split()
    if words:
        keys.add(f"first:{words[0]}")
    if len(words) >= 2:
        keys.add("first2:" + "|".join(words[:2]))

    # Length band reduces comparisons between tiny and very long notes.
    keys.add(f"len:{max(1, len(normalized) // 30)}")
    return keys


def candidate_pairs(records: list, exhaustive_limit: int = 1800, max_bucket_size: int = 900) -> Iterator[tuple[int, int]]:
    """Yield likely candidate pairs.

    Small datasets are compared exhaustively. Large datasets are blocked by
    technical codes, relevant tokens, first words and length band.
    """
    n = len(records)
    if n <= exhaustive_limit:
        yield from combinations(range(n), 2)
        return

    buckets: dict[str, list[int]] = defaultdict(list)
    for idx, record in enumerate(records):
        for key in _record_keys(record):
            buckets[key].append(idx)

    seen: set[int] = set()
    for indexes in buckets.values():
        if len(indexes) < 2 or len(indexes) > max_bucket_size:
            continue
        indexes = sorted(set(indexes))
        for i_pos, i in enumerate(indexes):
            for j in indexes[i_pos + 1 :]:
                packed = i * n + j
                if packed in seen:
                    continue
                seen.add(packed)
                yield i, j


def count_candidate_pairs(records: list, exhaustive_limit: int = 1800, max_bucket_size: int = 900) -> int | None:
    n = len(records)
    if n <= exhaustive_limit:
        return n * (n - 1) // 2
    # Counting exactly would require materializing the same de-dup set. Avoid
    # doing that twice on large files; progress will be approximate.
    return None
