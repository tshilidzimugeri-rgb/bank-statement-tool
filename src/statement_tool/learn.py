"""Suggests a category for transactions the rules in config/categories.yaml
don't match, learnt from the categories a person typed in the same workbook.

It looks for the categorised transactions whose descriptions share the most
telling words with this one (TF-IDF cosine similarity: a word on many of
the workbook's transactions, like "purchase", counts for little; a shop's
name counts for a lot), comparing only money going the same way (in or
out). A transaction counts as alike only if it has this one's most
distinctive word - usually the shop or person paid - so sharing how it was
paid ("DEBIT CARD PURCHASE", "PAYSHAP PAY BY PROXY") or its month never
makes two transactions alike. A category is suggested only when the alike
transactions agree on it; otherwise the transaction stays uncategorised.

Nothing is stored or sent anywhere: it learns afresh from the workbook each
time, so a category typed into the workbook teaches it at once. A
suggestion only ever sets the Category - never an amount, date or balance -
and is marked as a suggestion in the workbook, for a person to check.
"""
from __future__ import annotations

import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass

# How alike (0-1) a transaction a person categorised must be to count; one
# sharing only common words scores far lower.
MIN_SIMILARITY = 0.4
MIN_AGREEMENT = 0.75  # share of the alike transactions' weight that must be on one category
NEIGHBOURS = 7  # most alike transactions considered

# Words only: numbers (references, card and account numbers, dates) say
# nothing about what a transaction is for - nor do month names ("10 MAY").
_WORD = re.compile(r"[a-z][a-z&']+")
_MONTHS = {"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "sept", "oct", "nov", "dec", "january",
           "february", "march", "april", "june", "july", "august", "september", "october", "november", "december"}


def words(description: str | None) -> list[str]:
    return [w for w in _WORD.findall((description or "").lower()) if w not in _MONTHS]


@dataclass
class Suggestion:
    category: str
    alike: int  # transactions a person categorised that it was learnt from
    example: str  # the most alike of them

    def note(self) -> str:
        """What the workbook's Category Source column says."""
        return (f'Suggested - like {self.alike} transaction{"s" if self.alike > 1 else ""} you categorised, '
                f'e.g. "{self.example[:50]}"')


class CategoryLearner:
    def __init__(self, examples: list[tuple[str, bool, str]], other_descriptions: list[str] = ()):
        """examples: (description, is money in, category) of transactions a
        person categorised. other_descriptions: the rest of the workbook's, used
        only to tell common words from distinctive ones."""
        docs = [(Counter(words(desc)), money_in, category, desc) for desc, money_in, category in examples]
        docs = [d for d in docs if d[0]]
        corpus = [set(counts) for counts, *_ in docs] + [set(words(d)) for d in other_descriptions]
        corpus = [c for c in corpus if c]
        self._seen_in = Counter(w for c in corpus for w in c)
        self._size = len(corpus)
        self._docs: list[tuple[dict[str, float], bool, str, str]] = []
        self._by_word: dict[str, list[int]] = defaultdict(list)
        for counts, money_in, category, desc in docs:
            for w in counts:
                self._by_word[w].append(len(self._docs))
            self._docs.append((self._vector(counts), money_in, category, desc))

    def _idf(self, word: str) -> float:
        return math.log((1 + self._size) / (1 + self._seen_in.get(word, 0))) + 1

    def _distinctive(self, word: str) -> bool:
        return self._seen_in.get(word, 0) <= self._size / 2

    def _vector(self, counts: Counter) -> dict[str, float]:
        vec = {w: n * self._idf(w) for w, n in counts.items()}
        norm = math.sqrt(sum(v * v for v in vec.values()))
        return {w: v / norm for w, v in vec.items()} if norm else {}

    def suggest(self, description: str | None, money_in: bool) -> Suggestion | None:
        vec = self._vector(Counter(words(description)))
        if not vec:
            return None
        # Its most distinctive word - usually the shop or person ("SASOL",
        # "J SMITH") - must be one a person categorised it by; sharing only
        # how it was paid ("PAYSHAP PAY BY PROXY") doesn't make it alike.
        rarest = min(self._seen_in.get(w, 0) for w in vec)
        key_words = {w for w in vec if self._seen_in.get(w, 0) == rarest and self._distinctive(w)}
        similarity: dict[int, float] = defaultdict(float)
        for w, v in vec.items():
            for i in self._by_word.get(w, ()):
                if self._docs[i][1] == money_in:
                    similarity[i] += v * self._docs[i][0][w]
        alike = [(i, s) for i, s in sorted(similarity.items(), key=lambda x: -x[1])
                 if s >= MIN_SIMILARITY and key_words & self._docs[i][0].keys()][:NEIGHBOURS]
        if not alike:
            return None
        weight: dict[str, float] = defaultdict(float)
        for i, s in alike:
            weight[self._docs[i][2]] += s
        best = max(weight, key=weight.get)
        if weight[best] / sum(weight.values()) < MIN_AGREEMENT:
            return None  # alike transactions disagree: leave it for a person
        agreeing = [i for i, _ in alike if self._docs[i][2] == best]
        return Suggestion(best, len(agreeing), self._docs[agreeing[0]][3])
