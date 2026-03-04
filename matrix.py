"""
matrix.py — Fusion, déduplication et enrichissement
"""

import re
import os
import json
import asyncio
import anthropic
from loguru import logger

MODEL = os.getenv("CLAUDE_MODEL", "claude-3-5-haiku-20241022")

_TYPE_PREFIX = {
    "functional":  "F",
    "performance": "NF-PERF",
    "security":    "NF-SEC",
    "operational": "NF-OPS",
    "data":        "NF-DATA",
    "compliance":  "NF-CONF",
    "usability":   "NF-UX",
    "implicit":    "IMP",
    "predictive":  "PRED",
}

_MOSCOW_ORDER = {"must": 0, "should": 1, "could": 2, "wont": 3}

_SOURCE_PRIORITY = {
    "explicit": 0, "signal": 1, "functional": 0, "constraints": 0,
    "implicit": 2, "stakeholders": 2, "predictive": 3, "signals": 1,
}


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.lower().strip())[:120]


def _deduplicate(items: list) -> list:
    items_sorted = sorted(
        items,
        key=lambda x: _SOURCE_PRIORITY.get(x.get("source", "predictive"), 3)
    )
    seen_quotes = {}
    seen_prefix = {}
    unique = []
    for item in items_sorted:
        quote = _normalize(item.get("source_quote", ""))
        if quote and len(quote) > 20:
            if quote in seen_quotes:
                continue
            seen_quotes[quote] = True
        prefix = _normalize(item.get("text", ""))[:80]
        if len(prefix) < 15:
            continue
        if prefix in seen_prefix:
            continue
        seen_prefix[prefix] = True
        unique.append(item)
    logger.info(f"🧹 Déduplication : {len(items)} → {len(unique)}")
    return unique


def _assign_ids(items: list) -> list:
    counters = {}
    for item in items:
        t = item.get("type", "functional")
        prefix = _TYPE_PREFIX.get(t, "F")
        counters[prefix] = counters.get(prefix, 0) + 1
        item["id"] = f"{prefix}-{counters[prefix]:03d}"
    return items


def merge(readings: dict) -> list:
    all_items = []
    for reader_name, items in readings.items():
        for item in items:
            item.setdefault("type", "functional")
            item.setdefault("moscow", "should")
            item.setdefault("source", reader_name)
            item.setdefault("source_quote", "")
            item.setdefault("rationale", "")
            item.setdefault("stakeholder", "")
            item.setdefault("text", "")
            all_items.append(item)
    unique = _deduplicate(all_items)
    unique.sort(key=lambda x: (
        _MOSCOW_ORDER.get(x.get("moscow", "should"), 1),
        x.get("type", "functional"),
    ))
    unique = _assign_ids(unique)
    logger.success(
        f"✅ Matrice finale : {len(unique)} exigences "
        f"({sum(1 for x in unique if x[\'moscow\']==\'must\')} Must | "
        f"{sum(1 for x in unique if x[\'moscow\']==\'should\')} Should)"
    )
    return unique


_RESPONSE_PROMPT = (
    "Tu es un expert en réponse à appel d\'offres.\n\n"
    "Pour chaque exigence ci-dessous, rédige une proposition de réponse concise (1 à 3 phrases)\n"
    "qui montre comment notre solution y répond.\n"
    "Commence par \'Notre solution...\' ou \'Nous proposons...\'.\n\n"
    "Retourne un tableau JSON et rien d\'autre :\n"
    "[{{\"id\": \"...\", \"response\": \"...\"}}]\n\n"
    "Exigences :\n{requirements_json}"
)


async def enrich_with_responses(matrix: list) -> list:
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    BATCH = 30
    batches = [matrix[i:i+BATCH] for i in range(0, len(matrix), BATCH)]
    id_to_response = {}

    def _clean(raw):
        raw = raw.strip()
        if raw.startswith("```"):
            raw = re.sub(r"^```[a-z]*\n?", "", raw)
            raw = re.sub(r"\n?```$", "", raw)
        return raw.strip()

    for i, batch in enumerate(batches):
        logger.info(f"✍️  Réponses — lot {i+1}/{len(batches)}")
        batch_input = [{"id": x["id"], "text": x["text"], "type": x["type"]} for x in batch]
        prompt = _RESPONSE_PROMPT.format(
            requirements_json=json.dumps(batch_input, ensure_ascii=False)
        )
        try:
            loop = asyncio.get_event_loop()
            response = await loop.run_in_executor(
                None,
                lambda p=prompt: client.messages.create(
                    model=MODEL,
                    max_tokens=4096,
                    messages=[{"role": "user", "content": p}],
                )
            )
            for r in json.loads(_clean(response.content[0].text)):
                id_to_response[r["id"]] = r.get("response", "")
        except Exception as e:
            logger.error(f"❌ Erreur lot {i+1} : {e}")

    for item in matrix:
        item["proposed_response"] = id_to_response.get(item["id"], "")

    logger.success(f"✅ {len(id_to_response)}/{len(matrix)} réponses générées")
    return matrix
