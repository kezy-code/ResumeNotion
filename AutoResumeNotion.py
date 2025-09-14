import os
import re
import json
import requests
import subprocess
import sys
import traceback
from shutil import which
from typing import List
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn

# -------------------------
def fatal_error(message: str, details: str = ""):
    """Affiche un message d'erreur clair et quitte."""
    sys.exit(f"\n❌ ERREUR CRITIQUE — {message}\n{details}\n")

# -------------------------
# Configuration via variables d’environnement
NOTION_TOKEN = "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"
DATABASE_ID = "cfe629450b9141f0be15d67b9969fa27"
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "gpt-oss:20b")
OLLAMA_PATH = os.environ.get("OLLAMA_PATH")

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}" if NOTION_TOKEN else "",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json"
}

# -------------------------
# Vérification des prérequis
if not NOTION_TOKEN:
    fatal_error("Le token Notion (NOTION_TOKEN) est manquant.")
if not DATABASE_ID:
    fatal_error("L’ID de la base Notion (NOTION_DATABASE_ID) est manquant.")

ollama_exec = which("ollama") or OLLAMA_PATH
if not ollama_exec or not os.path.exists(ollama_exec):
    fatal_error("Ollama n’a pas été trouvé.", f"Chemin recherché : {OLLAMA_PATH}")

# -------------------------
def chunk_text_preserve_words(s: str, max_len: int) -> List[str]:
    if not s:
        return []
    i, L = 0, len(s)
    parts = []
    while i < L:
        if L - i <= max_len:
            parts.append(s[i:].rstrip())
            break
        cut = i + max_len
        chunk = s[i:cut]
        last_nl = chunk.rfind('\n')
        last_sp = chunk.rfind(' ')
        if last_nl > 0:
            split = i + last_nl
        elif last_sp > int(max_len * 0.6):
            split = i + last_sp
        else:
            split = cut
        parts.append(s[i:split].rstrip())
        i = split
        while i < L and s[i] in (' ', '\n'):
            i += 1
    return parts

def extract_json_from_output(s: str) -> dict:
    if not s:
        fatal_error("Sortie du modèle vide.")
    candidates = re.findall(r"\{.*?\}", s, re.DOTALL)
    if not candidates:
        fatal_error("Aucun objet JSON trouvé dans la sortie du modèle.", f"Sortie brute:\n{s}")
    last_error = None
    for cand in reversed(candidates):
        try:
            return json.loads(cand)
        except Exception:
            try:
                cleaned = re.sub(r",\s*}", "}", cand)
                cleaned = re.sub(r",\s*]", "]", cleaned)
                return json.loads(cleaned)
            except Exception as e2:
                last_error = e2
                continue
    fatal_error(f"Échec du parsing JSON malgré {len(candidates)} candidats.",
                f"Dernière erreur: {last_error}\nSortie brute:\n{s}")

def get_latest_page(database_id: str) -> dict:
    url = f"https://api.notion.com/v1/databases/{database_id}/query"
    try:
        r = requests.post(url, headers=NOTION_HEADERS)
        r.raise_for_status()
    except Exception as e:
        fatal_error("Erreur lors de l’accès à la base Notion.", str(e))
    res = r.json()
    if "results" not in res or not res["results"]:
        fatal_error("Aucune page trouvée dans la base Notion.")
    pages = sorted(res["results"], key=lambda x: x["created_time"], reverse=True)
    return pages[0]

def gather_page_text(page_id: str) -> (str, str):
    try:
        page_resp = requests.get(f"https://api.notion.com/v1/pages/{page_id}", headers=NOTION_HEADERS)
        page_resp.raise_for_status()
        page = page_resp.json()
    except Exception as e:
        fatal_error("Erreur lors de la récupération des métadonnées de la page Notion.", str(e))
    title = ""
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title" and prop.get("title"):
            title = "".join([t.get("plain_text", "") for t in prop["title"]]).strip()
            break
    try:
        blocks_resp = requests.get(f"https://api.notion.com/v1/blocks/{page_id}/children", headers=NOTION_HEADERS)
        blocks_resp.raise_for_status()
        blocks = blocks_resp.json().get("results", [])
    except Exception as e:
        fatal_error("Erreur lors de la récupération du contenu de la page Notion.", str(e))
    pieces = [f"Titre de la page: {title}\n"]
    for b in blocks:
        t = b.get("type")
        segs = b.get(t, {}).get("rich_text", [])
        text = " ".join([s.get("plain_text", "") for s in segs]).strip()
        if t.startswith("heading_"):
            pieces.append(text.upper())
        elif t in ("paragraph", "bulleted_list_item", "numbered_list_item"):
            prefix = "- " if "list" in t else ""
            pieces.append(prefix + text)
    full_text = "\n\n".join([p for p in pieces if p])
    return title, full_text

def call_ollama_make_json(text: str, model: str) -> dict:
    prompt = f"""
Tu es un assistant expert en rédaction professionnelle en français.
Reçois le texte suivant entre triple guillemets et génère **strictement** un objet JSON (aucun commentaire, aucun texte additionnel)
avec les clés suivantes :
- "taches": liste de chaînes
- "appris": liste de chaînes
- "difficultes": liste de chaînes
- "objectifs": liste de chaînes
- "short_markdown": chaîne (résumé en Markdown)

Règles:
1. C'est un résumé de la semaine, des actions et observations professionnelles.
2. Si le texte est court, sois synthétique et propose des suggestions concrètes.
3. Limite chaque chaîne à environ 300-400 caractères maximum.
4. Ne retourne que l’objet JSON.

Texte:
\"\"\"{text}\"\"\"
"""
    cmd = [ollama_exec, "run", model, prompt]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    except Exception as e:
        fatal_error("Impossible d’exécuter Ollama.", str(e))
    if result.returncode != 0:
        fatal_error("Échec de l’appel à Ollama.", result.stderr)
    stdout = result.stdout.strip()
    if not stdout:
        fatal_error("Ollama n’a rien renvoyé.", result.stderr)
    parsed = extract_json_from_output(stdout)
    for k in ("taches", "appris", "difficultes", "objectifs", "short_markdown"):
        if k not in parsed:
            parsed[k] = [] if k != "short_markdown" else ""
    return parsed

def build_notion_blocks_from_json(title: str, parsed: dict) -> List[dict]:
    blocks = []
    title_text = f"📝 Résumé généré — {title}" if title else "📝 Résumé généré"
    blocks.append({
        "object": "block", "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": title_text}}]}
    })
    sections = [
        ("✅ Tâches effectuées", parsed.get("taches", [])),
        ("📚 Ce que j’ai appris", parsed.get("appris", [])),
        ("⚠️ Difficultés rencontrées", parsed.get("difficultes", [])),
        ("🎯 Objectifs pour la suite", parsed.get("objectifs", [])),
    ]
    for heading, items in sections:
        blocks.append({
            "object": "block", "type": "heading_3",
            "heading_3": {"rich_text": [{"type": "text", "text": {"content": heading}}]}
        })
        if not items:
            blocks.append({
                "object": "block", "type": "paragraph",
                "paragraph": {"rich_text":[{"type":"text","text":{"content": "Aucune information fournie / aucune donnée détectée."}}]}
            })
        else:
            for it in items:
                for piece in chunk_text_preserve_words(it, 1800):
                    blocks.append({
                        "object":"block", "type":"bulleted_list_item",
                        "bulleted_list_item":{"rich_text":[{"type":"text","text":{"content": piece}}]}
                    })
    return blocks

def append_children_to_page(page_id: str, children: List[dict]):
    url = f"https://api.notion.com/v1/blocks/{page_id}/children"
    resp = requests.patch(url, headers=NOTION_HEADERS, json={"children": children})
    if resp.status_code not in (200, 201):
        fatal_error("Erreur lors de l’ajout dans Notion.",
                    f"Code HTTP: {resp.status_code}\nRéponse: {resp.text}")
    return resp.json()

def main():
    total_steps = 4
    with Progress(
        SpinnerColumn(),
        TextColumn("[bold blue]Étape {task.completed}/{task.total} : {task.description}"),
        BarColumn(),
        TimeElapsedColumn(),
        transient=True
    ) as progress:
        task = progress.add_task("Initialisation…", total=total_steps)
        
        progress.update(task, description="Récupération de la page Notion", advance=1)
        page = get_latest_page(DATABASE_ID)
        page_id = page["id"]
        
        progress.update(task, description="Lecture du contenu", advance=1)
        title, full_text = gather_page_text(page_id)
        if not full_text.strip():
            fatal_error("La page ne contient pas de texte à résumer.")
        
        progress.update(task, description="Génération du résumé JSON", advance=1)
        parsed = call_ollama_make_json(full_text, OLLAMA_MODEL)
        
        progress.update(task, description="Construction des blocs pour Notion", advance=1)
        children = build_notion_blocks_from_json(title, parsed)
        append_children_to_page(page_id, children)

    print("\n✅ Résumé ajouté avec succès à la page Notion.")

if __name__ == "__main__":
    try:
        main()
    except Exception:
        fatal_error("Erreur inattendue.", traceback.format_exc())
