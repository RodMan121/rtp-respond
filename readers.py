"""
readers.py — 5 lectures parallèles du RFP
==========================================
"""

import os
import re
import json
import base64
import asyncio
import anthropic
from pathlib import Path
from loguru import logger

MODEL = os.getenv("CLAUDE_MODEL", "claude-opus-4-5")

_SYSTEM = (
    "Tu es un expert en analyse de cahiers des charges (RFP/CCTP). "
    "Tu réponds UNIQUEMENT en JSON valide, sans markdown, sans commentaire. "
    "Sois exhaustif et précis."
)

_PROMPTS = {

    "functional": """
Lis ce document de bout en bout.

Extrais TOUTES les exigences fonctionnelles : ce que le système doit faire,
les workflows, les règles métier, les actions utilisateur, les automatismes.

Retourne une liste JSON et rien d\'autre.

Format de chaque objet :
{
  "text": "Le système doit ... (reformulé clairement en français)",
  "type": "functional",
  "moscow": "must | should | could | wont",
  "source": "explicit",
  "source_quote": "citation verbatim extraite du document (50-200 chars)",
  "rationale": "pourquoi c\'est important",
  "stakeholder": "rôle concerné ou vide"
}

Règles MoSCoW :
- must   : explicitement obligatoire (must, shall, doit, devra)
- should : recommandé ou fortement attendu
- could  : optionnel, souhaitable
- wont   : explicitement exclu ou hors périmètre
""",

    "constraints": """
Lis ce document de bout en bout.

Extrais TOUTES les exigences non fonctionnelles :
performance, sécurité, hébergement, disponibilité, conformité réglementaire,
contraintes techniques, interopérabilité, volumétrie, SLA.

Retourne une liste JSON et rien d\'autre.

Format de chaque objet :
{
  "text": "reformulation claire en français",
  "type": "performance | security | operational | data | compliance | usability",
  "moscow": "must | should | could | wont",
  "source": "explicit",
  "source_quote": "citation verbatim du document",
  "rationale": "impact si non respecté",
  "stakeholder": "rôle concerné ou vide"
}
""",

    "stakeholders": """
Lis ce document de bout en bout.

Identifie chaque rôle ou partie prenante mentionné(e) dans le document.
Pour chaque rôle, déduis les besoins implicites que le document ne formule pas
explicitement mais qu\'un professionnel de ce domaine attendrait.

Exemples :
- "Administrateur"      → besoin de supervision, gestion des comptes, logs
- "Auditeur"            → besoin de traçabilité, exports, rapports d\'activité
- "Utilisateur externe" → besoin d\'onboarding, aide contextuelle, reset MDP

Retourne une liste JSON et rien d\'autre.

Format de chaque objet :
{
  "text": "reformulation de l\'exigence implicite en français",
  "type": "functional | security | usability | operational",
  "moscow": "should | could",
  "source": "implicit",
  "source_quote": "",
  "rationale": "ce rôle attend cela car ...",
  "stakeholder": "nom exact du rôle"
}
""",

    "implicit": """
Lis ce document de bout en bout.

Pour un projet de ce type et de ce domaine, identifie les besoins
que le client attend implicitement mais a oublié d\'écrire.

Catégories universelles à couvrir systématiquement :
1. Gestion du changement : migration des données existantes, reprise historique
2. Continuité : plan de sauvegarde, reprise après incident, PRA/PCA
3. Cycle de vie : versioning, rollback, gestion des mises à jour
4. Recette : procédure de test, critères d\'acceptance, environnement de QA
5. Formation : documentation utilisateur, formation initiale, support post-démarrage
6. Supervision : monitoring, alerting, tableau de bord opérationnel
7. Conformité : RGPD si données personnelles, normes sectorielles détectées
8. Intégration : APIs tierces implicitement nécessaires, SSO, annuaire

Retourne UNIQUEMENT les catégories réellement pertinentes pour CE document.
Retourne une liste JSON et rien d\'autre.

Format de chaque objet :
{
  "text": "reformulation de l\'exigence prédictive en français",
  "type": "operational | security | data | compliance | functional",
  "moscow": "should | could",
  "source": "predictive",
  "source_quote": "",
  "rationale": "pourquoi ce client en aura besoin",
  "stakeholder": ""
}
""",

    "signals": """
Lis ce document de bout en bout.

Identifie les signaux faibles : formulations qui révèlent une douleur passée,
un incident, une frustration avec un système précédent.

Indices :
- "must alert if duplication"  → ils ont eu des doublons
- "history must be displayed"  → ils ont perdu de la traçabilité
- "not editable once closed"   → des modifications non contrôlées ont eu lieu
- "must be stored in UTC"      → des problèmes de fuseau horaire ont existé

Pour chaque signal, formule l\'exigence renforcée qu\'il implique.
Retourne une liste JSON et rien d\'autre.

Format de chaque objet :
{
  "text": "exigence renforcée déduite du signal",
  "type": "functional | security | data | operational",
  "moscow": "must | should",
  "source": "signal",
  "source_quote": "formulation exacte révélant le signal",
  "rationale": "ce signal révèle que le client a vécu : ...",
  "stakeholder": ""
}
"""
}


def _load_pdf_b64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.standard_b64encode(f.read()).decode("utf-8")


def _clean_json(raw: str) -> str:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
    return raw.strip()


async def _call_reader(
    client: anthropic.Anthropic,
    reader_name: str,
    pdf_b64: str,
    semaphore: asyncio.Semaphore,
) -> list:
    async with semaphore:
        logger.info(f"📖 Lecteur \'{reader_name}\' — envoi au LLM...")
        try:
            loop = asyncio.get_event_loop()
            response = await loop.run_in_executor(
                None,
                lambda: client.messages.create(
                    model=MODEL,
                    max_tokens=8192,
                    system=_SYSTEM,
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "document",
                                    "source": {
                                        "type": "base64",
                                        "media_type": "application/pdf",
                                        "data": pdf_b64,
                                    },
                                },
                                {
                                    "type": "text",
                                    "text": _PROMPTS[reader_name],
                                },
                            ],
                        }
                    ],
                )
            )

            raw = _clean_json(response.content[0].text)
            items = json.loads(raw)
            if not isinstance(items, list):
                items = [items]

            for item in items:
                item.setdefault("source", reader_name)
                item.setdefault("stakeholder", "")
                item.setdefault("source_quote", "")
                item.setdefault("rationale", "")

            logger.success(f"✅ Lecteur \'{reader_name}\' — {len(items)} exigences")
            return items

        except json.JSONDecodeError as e:
            logger.error(f"❌ Lecteur \'{reader_name}\' — JSON invalide : {e}")
            return []
        except Exception as e:
            logger.error(f"❌ Lecteur \'{reader_name}\' — erreur : {e}")
            return []


async def run_all_readers(pdf_path: str) -> dict:
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    pdf_b64 = _load_pdf_b64(pdf_path)
    logger.info(f"📄 PDF chargé : {Path(pdf_path).name} ({len(pdf_b64)//1024} KB base64)")

    semaphore = asyncio.Semaphore(3)
    results = {}
    tasks = []
    for name in _PROMPTS:
        tasks.append(_call_reader(client, name, pdf_b64, semaphore))
    
    responses = await asyncio.gather(*tasks)
    for i, name in enumerate(_PROMPTS):
        results[name] = responses[i]

    total = sum(len(v) for v in results.values())
    logger.success(f"🏁 5 lectures terminées — {total} exigences brutes")
    return results
