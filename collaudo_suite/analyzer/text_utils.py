from __future__ import annotations

import re
from dataclasses import dataclass


# Conservative Italian stopword fallback. It is intentionally small: in technical
# logs, aggressive stopword removal can remove useful context.
FALLBACK_STOPWORDS = {
    "a", "ad", "al", "allo", "alla", "ai", "agli", "alle", "anche", "che", "chi", "ci",
    "coi", "col", "con", "da", "dal", "dallo", "dalla", "dai", "dagli", "dalle", "de", "dei",
    "del", "dello", "della", "degli", "delle", "di", "e", "ed", "gli", "i", "il", "in", "la",
    "le", "lo", "ma", "nel", "nello", "nella", "nei", "negli", "nelle", "non", "o", "od", "per",
    "piu", "se", "su", "sul", "sullo", "sulla", "sui", "sugli", "sulle", "tra", "un", "una",
    "uno", "come", "dove", "quando", "quindi", "poi", "prima", "dopo", "essere", "avere",
}

# Technical codes that must survive punctuation cleanup, for example: E32.0,
# DM30, S41A, 30F-6, BH4.0, X=3800.
TECH_CODE_PATTERNS = [
    re.compile(r"\b[A-Z]{1,8}\d+(?:[._-]\d+)*(?:[A-Z])?\b", re.IGNORECASE),
    re.compile(r"\b\d+[A-Z]+(?:[-_]\d+)?\b", re.IGNORECASE),
    re.compile(r"\b[A-Z]{1,4}=\d+(?:[.,]\d+)?\b", re.IGNORECASE),
]

_WORD_RE = re.compile(r"[a-zA-Z0-9]+")
_AXIS_RE = re.compile(r"\basse\s+([XYZAC])\b", re.IGNORECASE)


@dataclass(slots=True)
class NormalizedText:
    normalized: str
    tokens: set[str]
    technical_codes: set[str]


class TextNormalizer:
    """Normalize Italian technical text without destroying machine codes."""

    def __init__(self, stemming: bool = True):
        self.stemming_requested = stemming
        self.stop_words = set(FALLBACK_STOPWORDS)
        self.stemmer = None
        self.using_nltk_stopwords = False
        self.warning: str | None = None
        self._init_optional_nltk()

    def _init_optional_nltk(self) -> None:
        """Load local NLTK data only. Do not download at runtime."""
        try:
            from nltk.corpus import stopwords  # type: ignore
            from nltk.stem import SnowballStemmer  # type: ignore

            self.stop_words = set(stopwords.words("italian")) | FALLBACK_STOPWORDS
            self.stemmer = SnowballStemmer("italian") if self.stemming_requested else None
            self.using_nltk_stopwords = True
        except Exception:
            self.stemmer = None
            if self.stemming_requested:
                self.warning = (
                    "Stopword/stemmer NLTK non disponibili localmente: uso lista base interna "
                    "e disattivo lo stemming."
                )

    @staticmethod
    def extract_technical_codes(text: str) -> set[str]:
        source = str(text or "").upper()
        codes: set[str] = set()
        for pattern in TECH_CODE_PATTERNS:
            codes.update(match.group(0).replace(",", ".") for match in pattern.finditer(source))
        return codes

    def normalize(self, text: str) -> NormalizedText:
        original = str(text or "")
        codes = self.extract_technical_codes(original)
        axis_tokens = {f"ASSE_{match.group(1).upper()}" for match in _AXIS_RE.finditer(original)}

        # Remove technical codes before generic word tokenization. Without this,
        # E32.0 would also generate noisy tokens such as e32 and 0.
        masked = original
        for pattern in TECH_CODE_PATTERNS:
            masked = pattern.sub(" ", masked)

        lowered = masked.lower().replace("'", " ").replace("\u2019", " ")
        words = _WORD_RE.findall(lowered)
        filtered: list[str] = []

        for word in words:
            if word in self.stop_words:
                continue
            if len(word) <= 1:
                continue
            if self.stemmer is not None:
                word = self.stemmer.stem(word)
            filtered.append(word)

        # Technical codes and explicit axis references are added as exact tokens.
        # This keeps E32.0 different from E320 and prevents "asse A" from losing A.
        ordered_preserved = sorted(codes) + sorted(axis_tokens)
        tokens = set(filtered) | set(ordered_preserved)
        normalized = " ".join(ordered_preserved + filtered).strip()
        return NormalizedText(normalized=normalized, tokens=tokens, technical_codes=codes)


def short_label(text: str, max_words: int = 5, max_chars: int = 55) -> str:
    words = str(text or "").split()
    label = " ".join(words[:max_words]) if words else "N/D"
    if len(label) > max_chars:
        label = label[: max_chars - 3].rstrip() + "..."
    elif len(words) > max_words:
        label += "..."
    return label
