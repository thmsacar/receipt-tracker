#!/usr/bin/env python3
"""CLI tool for tracking, parsing, and analyzing expense receipts."""

import difflib
import hashlib
import json
import mimetypes
import os
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Literal, Optional

from dotenv import load_dotenv
from pydantic import BaseModel, Field
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
import typer

# Disable bytecode generation (__pycache__)
sys.dont_write_bytecode = True

load_dotenv()

console = Console()
app = typer.Typer(
    help="CLI tool for receipt extraction, tracking, and financial analysis.",
    add_completion=False,
)

# ==============================================================================
# Data Models and Taxonomy
# ==============================================================================

CategoryEnum = Literal[
    "Épicerie & Base",
    "Boucherie & Poissonnerie",
    "Fruits & Légumes",
    "Boissons & Snacks",
    "Entretien & Maison",
    "Hygiène & Beauté",
    "Autre",
]

CATEGORIES_ALIMENTAIRES: List[str] = [
    "Épicerie & Base",
    "Boucherie & Poissonnerie",
    "Fruits & Légumes",
    "Boissons & Snacks",
]

CATEGORIES_NON_ALIMENTAIRES: List[str] = [
    "Entretien & Maison",
    "Hygiène & Beauté",
    "Autre",
]

CATEGORIES_VALIDES: List[str] = CATEGORIES_ALIMENTAIRES + CATEGORIES_NON_ALIMENTAIRES


class ReceiptItem(BaseModel):
    """Single line item from a receipt."""

    raw_text: str = Field(
        description="Raw item text as printed on the receipt."
    )
    description: str = Field(
        description="Clean product name in French."
    )
    category: CategoryEnum = Field(
        description="Item category from the allowed taxonomy."
    )
    quantity: float = Field(
        default=1.0,
        description="Purchased quantity or weight in kg.",
    )
    unit: str = Field(
        default="pièce",
        description="Unit of measure (e.g. pièce, kg).",
    )
    unit_price: float = Field(
        description="Unit price in EUR."
    )
    total_price: float = Field(
        description="Total line amount in EUR."
    )


class ReceiptData(BaseModel):
    """Structured receipt data."""

    id: Optional[str] = Field(
        default=None,
        description="Receipt identifier (YYYY-MM-DD_XX).",
    )
    market: str = Field(
        description="Store or market brand name."
    )
    branch: Optional[str] = Field(
        default=None,
        description="Store branch or city.",
    )
    date: Optional[str] = Field(
        default=None,
        description="Receipt date (YYYY-MM-DD).",
    )
    receipt_no: Optional[str] = Field(
        default=None,
        description="Receipt or transaction number.",
    )
    total_amount: float = Field(
        description="Total amount paid in EUR."
    )
    currency: str = Field(
        default="EUR",
        description="Currency code (EUR).",
    )
    items: List[ReceiptItem] = Field(
        description="Purchased items on the receipt."
    )


# ==============================================================================
# AI Extraction Specification (System Prompt)
# ==============================================================================

SYSTEM_PROMPT = """Rôle : Extraction et numérisation de tickets de caisse de commerces et supermarchés français.

1. Résolution des libellés et abréviations :
   - Renseigner chaque article sous un nom clair et explicite en français dans le champ 'description'.
   - Développer les abréviations de caisse courantes.
   - Supprimer les codes-barres partiels, préfixes de pesée ou références internes.
   - Traduire en français clair les mentions manuscrites si présentes.

2. Catégorisation :
   Classer chaque article dans l'une des catégories suivantes :
   - 'Épicerie & Base' : Denrées de base, produits laitiers, conserves, huiles, féculents, pain, épices.
   - 'Boucherie & Poissonnerie' : Viandes, volailles, charcuterie, saucisses, poissons, fruits de mer.
   - 'Fruits & Légumes' : Fruits et légumes frais, salades, herbes fraîches, champignons.
   - 'Boissons & Snacks' : Boissons, eaux, jus, sodas, biscuits, chocolat, bonbons, fruits secs.
   - 'Entretien & Maison' : Produits d'entretien, droguerie, vaisselle, ustensiles.
   - 'Hygiène & Beauté' : Soins, cosmétiques, savon, shampoing, hygiène corporelle.
   - 'Autre' : Articles hors alimentation, maison ou hygiène (sacs payants, piles, presse).

3. Métadonnées :
   - 'market' : Nom de l'enseigne sans l'adresse (ex: 'Lidl', 'Carrefour Express', 'Boucherie Saint Bruno', 'Action', 'Ikea').
   - 'branch' : Ville ou quartier si présent sur le ticket (ex: 'Saint-Égrève', 'Grenoble').
   - 'date' : Date de transaction (format YYYY-MM-DD).
   - 'receipt_no' : Numéro de ticket si présent, sinon null.
   - 'total_amount' : Montant total payé en euros.
   - 'currency' : Par défaut 'EUR'.

4. Cohérence arithmétique :
   - Vérifier que total_price = quantity * unit_price (ajusté des remises de ligne si applicable).
   - La somme des articles doit correspondre au total_amount."""


# ==============================================================================
# Normalization and Matching Utilities
# ==============================================================================

def normalize_category(cat: Optional[str]) -> str:
    """Normalize a category string to the allowed taxonomy."""
    if not cat:
        return "Autre"
    cat_clean = cat.strip()
    if cat_clean in CATEGORIES_VALIDES:
        return cat_clean
    for valid_cat in CATEGORIES_VALIDES:
        if valid_cat.lower() == cat_clean.lower():
            return valid_cat
    return "Autre"


def normalize_market_name(
    market: Optional[str],
    existing_markets: Optional[Iterable[str]] = None,
) -> str:
    """Normalize store name by matching against existing known markets using fuzzy matching."""
    if not market or not str(market).strip():
        return "Commerce local"

    clean = re.sub(r"\(.*?\)", "", str(market)).strip()
    clean = re.sub(r"\s+", " ", clean)

    if existing_markets is None:
        try:
            p = Path("receipts.json")
            if p.exists():
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        existing_markets = [r.get("market") for r in data if r.get("market")]
        except Exception:
            existing_markets = None

    if not existing_markets:
        return clean.title()

    existing_list = [m for m in existing_markets if m and str(m).strip()]
    if not existing_list:
        return clean.title()

    # 1. Exact case-insensitive match
    for em in existing_list:
        if clean.lower() == em.lower():
            return em

    def simplify(s: str) -> str:
        s = s.lower()
        s = re.sub(r"\bst\b", "saint", s)
        s = re.sub(r"\bste\b", "sainte", s)
        s = re.sub(r"[^\w\s]", " ", s)
        return re.sub(r"\s+", " ", s).strip()

    s_clean = simplify(clean)

    # 2. Simplified match (abbreviations and punctuation)
    for em in existing_list:
        if s_clean == simplify(em):
            return em

    tokens_clean = set(s_clean.split())
    best_match = None
    best_score = 0.0

    # 3. Fuzzy similarity & token containment
    for em in existing_list:
        s_em = simplify(em)
        tokens_em = set(s_em.split())

        intersection = tokens_clean & tokens_em
        union = tokens_clean | tokens_em
        jaccard = len(intersection) / len(union) if union else 0.0

        subset_score = 0.0
        if intersection:
            smaller_len = min(len(tokens_clean), len(tokens_em))
            if len(intersection) == smaller_len:
                subset_score = 0.8 + (len(intersection) / max(len(tokens_clean), len(tokens_em))) * 0.15

        ratio = difflib.SequenceMatcher(None, s_clean, s_em).ratio()
        score = max(ratio, jaccard, subset_score)

        if score > best_score:
            best_score = score
            best_match = em

    if best_match and best_score >= 0.75:
        return best_match

    return clean.title()


def find_duplicate_receipt(
    receipt: Dict[str, Any],
    existing_receipts: List[Dict[str, Any]],
    image_hash: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    """Check if receipt already exists by image hash, receipt number, or date/amount match."""
    known_markets = [r.get("market") for r in existing_receipts if r.get("market")]
    m_new = normalize_market_name(receipt.get("market"), existing_markets=known_markets)
    d_new = receipt.get("date")
    tot_new = float(receipt.get("total_amount") or 0.0)
    no_new = str(receipt.get("receipt_no") or "").strip()

    for r in existing_receipts:
        if image_hash and r.get("image_hash") and r.get("image_hash") == image_hash:
            return r

        m_old = normalize_market_name(r.get("market"), existing_markets=known_markets)
        d_old = r.get("date")
        tot_old = float(r.get("total_amount") or 0.0)
        no_old = str(r.get("receipt_no") or "").strip()

        if no_new and no_old and no_new != "-" and no_new == no_old and m_new == m_old:
            return r

        if m_new == m_old and d_new == d_old and abs(tot_new - tot_old) < 0.01:
            return r

    return None


def sort_receipts_chronologically(receipts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Sort receipts by date and assign sequential IDs (YYYY-MM-DD_XX)."""
    def sort_key(r: Dict[str, Any]) -> str:
        d = str(r.get("date") or "9999-99-99")
        if len(d) == 7:
            d = f"{d}-01"
        return d

    sorted_list = sorted(receipts, key=sort_key)

    reindexed = []
    for idx, r in enumerate(sorted_list, start=1):
        d = str(r.get("date") or "2026-01-01")
        if len(d) == 7:
            d = f"{d}-01"
            r["date"] = d
        r["id"] = f"{d}_{idx:02d}"
        reindexed.append(r)

    return reindexed


def normalize_receipt_dict(r: Dict[str, Any], default_index: Optional[int] = None, existing_markets: Optional[Iterable[str]] = None) -> Dict[str, Any]:
    """Normalize raw receipt dictionary to standard schema."""
    raw_market = r.get("market") or "Commerce local"
    market = normalize_market_name(raw_market, existing_markets=existing_markets)
    branch = r.get("branch")
    m_paren = re.search(r"\((.*?)\)", str(raw_market))
    if m_paren:
        addr = m_paren.group(1).strip()
        if branch and addr.lower() not in branch.lower():
            branch = f"{addr}, {branch}"
        elif not branch:
            branch = addr

    date = r.get("date")
    if date and len(date) == 7:
        date = f"{date}-01"

    receipt_no = r.get("receipt_no")
    if receipt_no == "-":
        receipt_no = None

    total_amount = float(r.get("total_amount") or 0.0)
    currency = r.get("currency") or "EUR"

    r_id = r.get("id")
    if not r_id or not (date and date in r_id):
        idx_str = f"_{default_index:02d}" if default_index else "_01"
        r_id = f"{date or '2026-01-01'}{idx_str}"

    items_raw = r.get("items") or []
    norm_items = []

    for it in items_raw:
        raw = it.get("raw_text") or ""
        desc = it.get("description") or raw
        cat = normalize_category(it.get("category"))

        try:
            qty = float(it.get("quantity", 1.0))
        except (ValueError, TypeError):
            qty = 1.0

        unit = it.get("unit") or ("kg" if qty < 5 and "." in str(qty) and qty != int(qty) else "pièce")
        if unit == "adet":
            unit = "pièce"

        u_price = float(it.get("unit_price") or 0.0)
        t_price = float(it.get("total_price") or 0.0)

        norm_items.append({
            "raw_text": raw,
            "description": desc,
            "category": cat,
            "quantity": qty,
            "unit": unit,
            "unit_price": u_price,
            "total_price": t_price,
        })

    return {
        "id": r_id,
        "market": market,
        "branch": branch,
        "date": date,
        "receipt_no": receipt_no,
        "total_amount": round(total_amount, 2),
        "currency": currency,
        "items": norm_items,
    }


def load_receipts(file_path: Path) -> List[Dict[str, Any]]:
    """Load receipts from a JSON file."""
    if not file_path.exists():
        return []
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except json.JSONDecodeError as e:
        console.print(f"[bold red]Erreur de lecture JSON dans {file_path}:[/bold red] {e}")
        return []


def save_receipts(file_path: Path, data: List[Dict[str, Any]]) -> None:
    """Save receipts to a JSON file."""
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def sync_excel_file(
    json_file: Path = Path("receipts.json"),
    excel_file: Path = Path("depenses.xlsx"),
) -> None:
    """Rebuild the Excel workbook from receipts.json."""
    try:
        from generate_excel import generate_excel
        generate_excel(str(json_file), str(excel_file))
        console.print(f"[dim green]Le fichier Excel [bold]{excel_file.name}[/bold] a été synchronisé avec succès.[/dim green]")
    except Exception as e:
        console.print(f"[dim yellow]Avertissement : La synchronisation Excel a échoué : {e}[/dim yellow]")


def compute_financial_statistics(receipts: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Compute financial summary metrics, food vs non-food breakdown, and store totals."""
    total_general = sum(float(r.get("total_amount", 0.0)) for r in receipts)
    cat_spend: Dict[str, float] = {c: 0.0 for c in CATEGORIES_VALIDES}
    cat_items: Dict[str, int] = {c: 0 for c in CATEGORIES_VALIDES}
    market_spend: Dict[str, float] = {}
    market_counts: Dict[str, int] = {}

    known_markets = [r.get("market") for r in receipts if r.get("market")]

    for r in receipts:
        m = normalize_market_name(r.get("market") or "Inconnu", existing_markets=known_markets)
        market_spend[m] = market_spend.get(m, 0.0) + float(r.get("total_amount", 0.0))
        market_counts[m] = market_counts.get(m, 0) + 1

        for it in r.get("items", []):
            c = normalize_category(it.get("category"))
            cat_spend[c] += float(it.get("total_price", 0.0))
            cat_items[c] += 1

    food_total = sum(cat_spend[c] for c in CATEGORIES_ALIMENTAIRES)
    food_items = sum(cat_items[c] for c in CATEGORIES_ALIMENTAIRES)
    nonfood_total = sum(cat_spend[c] for c in CATEGORIES_NON_ALIMENTAIRES)
    nonfood_items = sum(cat_items[c] for c in CATEGORIES_NON_ALIMENTAIRES)

    food_pct = (food_total / total_general * 100) if total_general else 0.0
    nonfood_pct = (nonfood_total / total_general * 100) if total_general else 0.0

    return {
        "total_general": total_general,
        "total_items": sum(cat_items.values()),
        "total_receipts": len(receipts),
        "avg_receipt": (total_general / len(receipts)) if receipts else 0.0,
        "food_total": food_total,
        "food_items": food_items,
        "food_pct": food_pct,
        "nonfood_total": nonfood_total,
        "nonfood_items": nonfood_items,
        "nonfood_pct": nonfood_pct,
        "cat_spend": cat_spend,
        "cat_items": cat_items,
        "market_spend": market_spend,
        "market_counts": market_counts,
    }


# ==============================================================================
# Gemini Vision API Integration
# ==============================================================================

def analyze_receipt_with_gemini(
    image_path: Path,
    api_key: str,
    model_name: str = "gemini-2.5-flash",
) -> ReceiptData:
    """Extract structured receipt data from an image using Gemini vision API."""
    try:
        from google import genai
        from google.genai import types
    except ImportError:
        console.print("[bold red]Le paquet 'google-genai' n'est pas installé.[/bold red]")
        console.print("Veuillez exécuter : [cyan]pip install -r requirements.txt[/cyan]")
        raise typer.Exit(code=1)

    if not image_path.exists():
        console.print(f"[bold red]Erreur : Le fichier image '{image_path}' est introuvable.[/bold red]")
        raise typer.Exit(code=1)

    mime_type, _ = mimetypes.guess_type(str(image_path))
    if not mime_type or not mime_type.startswith("image/"):
        ext = image_path.suffix.lower()
        if ext in [".jpg", ".jpeg"]:
            mime_type = "image/jpeg"
        elif ext == ".png":
            mime_type = "image/png"
        elif ext == ".webp":
            mime_type = "image/webp"
        else:
            mime_type = "image/jpeg"

    with open(image_path, "rb") as f:
        image_bytes = f.read()

    client = genai.Client(api_key=api_key)
    image_part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)

    user_prompt = (
        "Analyse cette image de ticket de caisse. Extrais toutes les informations requises, "
        "normalise les libellés d'articles en français standard et classe-les selon le schéma JSON fourni."
    )

    contents = [image_part, user_prompt]

    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        response_mime_type="application/json",
        response_schema=ReceiptData,
        temperature=0.1,
    )

    try:
        response = client.models.generate_content(
            model=model_name,
            contents=contents,
            config=config,
        )
    except Exception as e:
        if "not found" in str(e).lower() and model_name != "gemini-2.0-flash":
            console.print(f"[yellow]Modèle '{model_name}' indisponible, bascule sur 'gemini-2.0-flash'...[/yellow]")
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=contents,
                config=config,
            )
        else:
            raise e

    if hasattr(response, "parsed") and response.parsed:
        return response.parsed

    return ReceiptData.model_validate_json(response.text)


# ==============================================================================
# CLI Commands
# ==============================================================================

@app.command()
def add(
    image: Path = typer.Argument(
        ...,
        help="Chemin vers l'image du ticket de caisse (JPG, PNG, WEBP)",
        exists=True,
        file_okay=True,
        dir_okay=False,
        readable=True,
    ),
    file: Path = typer.Option(
        Path("receipts.json"),
        "--file", "-f",
        help="Chemin vers le fichier JSON de stockage des tickets",
    ),
    model: str = typer.Option(
        os.environ.get("GEMINI_MODEL", "gemini-2.5-flash"),
        "--model", "-m",
        help="Modèle Gemini Flash à utiliser",
    ),
    force: bool = typer.Option(
        False,
        "--force", "-F",
        help="Forcer l'ajout même si le ticket ou la photo semble être un doublon",
    ),
) -> None:
    """Scan a receipt image, store it, and update depenses.xlsx."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        console.print(Panel.fit(
            "[bold red]Clé d'API GEMINI_API_KEY introuvable.[/bold red]\n\n"
            "Pour utiliser l'extraction automatique :\n"
            "1. Créez un fichier [bold cyan].env[/bold cyan] à la racine du projet.\n"
            "2. Ajoutez votre clé gratuite Google AI Studio :\n"
            "   [bold yellow]GEMINI_API_KEY=votre_cle_ici[/bold yellow]\n\n"
            "Ou exportez-la dans votre terminal :\n"
            "   [bold yellow]export GEMINI_API_KEY=votre_cle_ici[/bold yellow]",
            title="Configuration requise",
            border_style="red",
        ))
        raise typer.Exit(code=1)

    current_receipts = load_receipts(file)

    with open(image, "rb") as f:
        img_bytes = f.read()
    image_hash = hashlib.sha256(img_bytes).hexdigest()

    if not force:
        dup_by_hash = next((r for r in current_receipts if r.get("image_hash") == image_hash), None)
        if dup_by_hash:
            console.print(Panel(
                f"[bold yellow]Cette photo exacte a déjà été ajoutée.[/bold yellow]\n\n"
                f"• ID du ticket  : [bold cyan]{dup_by_hash.get('id')}[/bold cyan]\n"
                f"• Magasin       : [bold white]{dup_by_hash.get('market')}[/bold white]\n"
                f"• Date          : [bold white]{dup_by_hash.get('date')}[/bold white]\n"
                f"• Montant total : [bold green]{dup_by_hash.get('total_amount'):.2f} €[/bold green]\n\n"
                f"[dim]L'ajout a été annulé pour éviter de doubler vos dépenses.\n"
                f"Pour forcer l'ajout, relancez avec l'option : [bold cyan]--force[/bold cyan][/dim]",
                title="Doublon détecté",
                border_style="yellow",
            ))
            return

    with console.status(f"[bold green]Analyse du ticket avec Gemini Flash ({model})...[/bold green]"):
        try:
            parsed_data = analyze_receipt_with_gemini(image, api_key=api_key, model_name=model)
        except Exception as e:
            console.print(f"[bold red]Erreur lors de l'appel à l'API Gemini :[/bold red] {e}")
            raise typer.Exit(code=1)

    receipt_dict = parsed_data.model_dump()
    receipt_dict["image_hash"] = image_hash

    existing_markets = [r.get("market") for r in current_receipts if r.get("market")]
    receipt_dict["market"] = normalize_market_name(receipt_dict.get("market"), existing_markets=existing_markets)

    if not force:
        dup = find_duplicate_receipt(receipt_dict, current_receipts, image_hash=image_hash)
        if dup:
            console.print(Panel(
                f"[bold yellow]Ce ticket semble déjà exister dans la base.[/bold yellow]\n\n"
                f"Un ticket identique a été détecté :\n"
                f"• ID existant   : [bold cyan]{dup.get('id')}[/bold cyan]\n"
                f"• Magasin       : [bold white]{dup.get('market')}[/bold white]\n"
                f"• Date          : [bold white]{dup.get('date')}[/bold white]\n"
                f"• Montant       : [bold green]{dup.get('total_amount'):.2f} €[/bold green]\n"
                f"• N° Reçu       : [bold white]{dup.get('receipt_no') or '-'}[/bold white]\n\n"
                f"[dim]Pour forcer l'ajout, relancez avec : [bold cyan]python cli.py add \"{image}\" --force[/bold cyan][/dim]",
                title="Ticket doublon détecté",
                border_style="yellow",
            ))
            return

    current_receipts.append(receipt_dict)
    current_receipts = sort_receipts_chronologically(current_receipts)
    save_receipts(file, current_receipts)

    assigned_id = receipt_dict.get("id", "Nouveau")
    for r in current_receipts:
        if r.get("receipt_no") == receipt_dict.get("receipt_no") and r.get("total_amount") == receipt_dict.get("total_amount"):
            assigned_id = r.get("id")
            break

    meta_table = Table.grid(padding=1)
    meta_table.add_column(style="bold cyan", justify="right")
    meta_table.add_column(style="white")
    meta_table.add_row("ID Ticket :", f"[bold yellow]{assigned_id}[/bold yellow]")
    meta_table.add_row("Magasin :", receipt_dict.get("market") or "-")
    meta_table.add_row("Localisation :", receipt_dict.get("branch") or "-")
    meta_table.add_row("Date :", receipt_dict.get("date") or "-")
    meta_table.add_row("N° Reçu :", receipt_dict.get("receipt_no") or "-")
    meta_table.add_row("Montant Total :", f"[bold green]{receipt_dict.get('total_amount'):.2f} {receipt_dict.get('currency', 'EUR')}[/bold green]")

    console.print(Panel(meta_table, title=f"Ticket extrait : {image.name}", border_style="cyan"))

    items_table = Table(title="Articles détectés et normalisés", show_header=True, header_style="bold magenta")
    items_table.add_column("#", justify="right", style="dim", width=4)
    items_table.add_column("Description normalisée", style="bold white")
    items_table.add_column("Texte brut", style="dim")
    items_table.add_column("Catégorie", style="cyan")
    items_table.add_column("Quantité", justify="right")
    items_table.add_column("P.U. €", justify="right")
    items_table.add_column("Total €", justify="right", style="green")

    for i, it in enumerate(receipt_dict.get("items", []), start=1):
        items_table.add_row(
            str(i),
            it.get("description", ""),
            it.get("raw_text", ""),
            it.get("category", ""),
            f"{it.get('quantity', 1):.3g} {it.get('unit', '')}".strip(),
            f"{it.get('unit_price', 0):.2f} €",
            f"{it.get('total_price', 0):.2f} €",
        )

    console.print(items_table)
    console.print(
        f"\n[bold green]Succès :[/bold green] Ticket enregistré sous l'identifiant [bold yellow]{assigned_id}[/bold yellow] dans [bold]{file.name}[/bold]. "
        f"Total : [bold cyan]{len(current_receipts)}[/bold cyan] tickets.\n"
    )

    sync_excel_file(file, Path("depenses.xlsx"))


@app.command(name="list")
def list_receipts(
    file: Path = typer.Option(
        Path("receipts.json"),
        "--file", "-f",
        help="Chemin vers le fichier JSON",
    ),
    limit: Optional[int] = typer.Option(
        None,
        "--limit", "-n",
        help="Nombre maximum de tickets à afficher",
    ),
) -> None:
    """List receipts in chronological order."""
    receipts_raw = load_receipts(file)
    if not receipts_raw:
        console.print(f"[yellow]Aucun ticket trouvé dans '{file}'.[/yellow]")
        console.print("Utilisez [bold cyan]python cli.py add image.jpg[/bold cyan] pour ajouter un premier ticket.")
        return

    sorted_receipts = sort_receipts_chronologically(receipts_raw)

    if limit and limit > 0:
        display_receipts = sorted_receipts[-limit:]
    else:
        display_receipts = sorted_receipts

    table = Table(
        title=f"📊 Tickets de caisse enregistrés - Total : {len(sorted_receipts)}",
        show_header=True,
        header_style="bold navy_blue on white",
        border_style="bright_blue",
    )
    table.add_column("#", justify="right", style="dim", width=4)
    table.add_column("ID Ticket", style="bold yellow", width=16)
    table.add_column("Date", style="bold cyan", width=12)
    table.add_column("Magasin", style="bold white", width=28)
    table.add_column("Localisation", style="white", width=22)
    table.add_column("Articles", justify="right", width=9)
    table.add_column("Montant €", justify="right", style="bold green", width=13)
    table.add_column("N° Reçu", justify="center", style="dim", width=10)

    total_depenses = 0.0
    total_articles = 0

    for i, r in enumerate(display_receipts, start=1):
        amt = r.get("total_amount", 0.0)
        items_cnt = len(r.get("items", []))
        total_depenses += amt
        total_articles += items_cnt

        table.add_row(
            str(i),
            str(r.get("id") or "-"),
            str(r.get("date") or "-"),
            str(r.get("market") or "-"),
            str(r.get("branch") or "-"),
            str(items_cnt),
            f"{amt:,.2f} €",
            str(r.get("receipt_no") or "-"),
        )

    console.print(table)

    avg_ticket = total_depenses / len(display_receipts) if display_receipts else 0.0
    summary_panel = Panel.fit(
        f"[bold]Dépenses totales :[/bold] [bold green]{total_depenses:,.2f} €[/bold green]   |   "
        f"[bold]Tickets :[/bold] [bold cyan]{len(display_receipts)}[/bold cyan]   |   "
        f"[bold]Articles achetés :[/bold] [bold yellow]{total_articles}[/bold yellow]   |   "
        f"[bold]Moyenne / ticket :[/bold] [bold magenta]{avg_ticket:,.2f} €[/bold magenta]",
        border_style="green",
        title="📈 Synthèse financière",
    )
    console.print(summary_panel)


@app.command()
def show(
    identifier: str = typer.Argument(
        ...,
        help="Numéro d'ordre ou identifiant ID du ticket",
    ),
    file: Path = typer.Option(
        Path("receipts.json"),
        "--file", "-f",
        help="Chemin vers le fichier JSON",
    ),
) -> None:
    """Show itemized details for a receipt."""
    receipts_raw = load_receipts(file)
    if not receipts_raw:
        console.print(f"[red]Aucun ticket trouvé dans '{file}'.[/red]")
        raise typer.Exit(code=1)

    sorted_receipts = sort_receipts_chronologically(receipts_raw)

    target_receipt = None
    target_idx = None

    if identifier.isdigit():
        idx = int(identifier)
        if 1 <= idx <= len(sorted_receipts):
            target_receipt = sorted_receipts[idx - 1]
            target_idx = idx

    if not target_receipt:
        for idx, r in enumerate(sorted_receipts, start=1):
            if r.get("id") == identifier:
                target_receipt = r
                target_idx = idx
                break

    if not target_receipt:
        console.print(
            f"[bold red]Erreur : Aucun ticket trouvé pour l'identifiant '{identifier}'.[/bold red]\n"
            f"Vérifiez les identifiants avec [bold cyan]python cli.py list[/bold cyan]."
        )
        raise typer.Exit(code=1)

    r = target_receipt
    branch_str = f" - {r.get('branch')}" if r.get("branch") else ""
    header_text = (
        f"[bold cyan]ID :[/bold cyan] [bold yellow]{r.get('id')}[/bold yellow] - Ticket #{target_idx}\n"
        f"[bold cyan]Magasin :[/bold cyan] {r.get('market')}{branch_str}\n"
        f"[bold cyan]Date :[/bold cyan] {r.get('date') or '-'}\n"
        f"[bold cyan]Numéro de reçu :[/bold cyan] {r.get('receipt_no') or '-'}\n"
        f"[bold cyan]Montant total TTC :[/bold cyan] [bold green]{r.get('total_amount', 0):.2f} {r.get('currency', 'EUR')}[/bold green]"
    )
    console.print(Panel(header_text, title=f"🧾 Détail du Ticket {r.get('id')}", border_style="cyan"))

    table = Table(show_header=True, header_style="bold magenta")
    table.add_column("#", justify="right", width=4)
    table.add_column("Description normalisée", style="bold white")
    table.add_column("Texte brut", style="dim")
    table.add_column("Catégorie", style="cyan")
    table.add_column("Quantité", justify="right")
    table.add_column("P.U. €", justify="right")
    table.add_column("Total €", justify="right", style="bold green")

    for idx, it in enumerate(r.get("items", []), start=1):
        table.add_row(
            str(idx),
            it.get("description", ""),
            it.get("raw_text", ""),
            it.get("category", ""),
            f"{it.get('quantity', 1):.3g} {it.get('unit', '')}".strip(),
            f"{it.get('unit_price', 0):.2f} €",
            f"{it.get('total_price', 0):.2f} €",
        )

    console.print(table)


@app.command()
def stats(
    file: Path = typer.Option(
        Path("receipts.json"),
        "--file", "-f",
        help="Chemin vers le fichier JSON",
    ),
) -> None:
    """Show spending statistics by category and store."""
    receipts_raw = load_receipts(file)
    if not receipts_raw:
        console.print(f"[yellow]Aucun ticket trouvé dans '{file}'.[/yellow]")
        return

    normalized_receipts = sort_receipts_chronologically(receipts_raw)
    st = compute_financial_statistics(normalized_receipts)

    total_general = st["total_general"]
    cat_spend = st["cat_spend"]
    cat_items = st["cat_items"]
    food_total = st["food_total"]
    food_items = st["food_items"]
    food_pct = st["food_pct"]
    nonfood_total = st["nonfood_total"]
    nonfood_items = st["nonfood_items"]
    nonfood_pct = st["nonfood_pct"]
    market_spend = st["market_spend"]
    market_counts = st["market_counts"]

    # 1. Macro Summary: Food vs Non-Food
    macro_table = Table(
        title="⚖️ Synthèse Macro : Alimentaire vs Non-Alimentaire",
        show_header=True,
        header_style="bold white on navy_blue",
    )
    macro_table.add_column("Secteur de Dépense", style="bold white", width=36)
    macro_table.add_column("Articles", justify="right", width=10)
    macro_table.add_column("Total Dépensé €", justify="right", style="bold green", width=16)
    macro_table.add_column("Part %", justify="right", style="cyan", width=10)

    macro_table.add_row(
        "🥗 Alimentaire & Comestibles",
        str(food_items),
        f"{food_total:,.2f} €",
        f"{food_pct:.1f} %",
    )
    macro_table.add_row(
        "🛍️ Non-Alimentaire : Maison, Hygiène, Divers",
        str(nonfood_items),
        f"{nonfood_total:,.2f} €",
        f"{nonfood_pct:.1f} %",
    )
    macro_table.add_row(
        "[bold]TOTAL GÉNÉRAL[/bold]",
        f"[bold]{st['total_items']}[/bold]",
        f"[bold green]{total_general:,.2f} €[/bold green]",
        "[bold cyan]100.0 %[/bold cyan]",
    )
    console.print(macro_table)

    # 2. Food categories breakdown
    food_table = Table(
        title="🥗 Dépenses Alimentaires",
        show_header=True,
        header_style="bold white on dark_green",
    )
    food_table.add_column("Catégorie Alimentaire", style="bold white", width=28)
    food_table.add_column("Articles", justify="right", width=10)
    food_table.add_column("Dépenses €", justify="right", style="bold green", width=14)
    food_table.add_column("Part Alim. %", justify="right", style="cyan", width=14)
    food_table.add_column("Part Budget %", justify="right", style="dim cyan", width=14)

    for cat_name in sorted(CATEGORIES_ALIMENTAIRES, key=lambda c: cat_spend[c], reverse=True):
        amt = cat_spend[cat_name]
        p_food = (amt / food_total * 100) if food_total else 0.0
        p_glob = (amt / total_general * 100) if total_general else 0.0
        food_table.add_row(
            cat_name,
            str(cat_items.get(cat_name, 0)),
            f"{amt:,.2f} €",
            f"{p_food:.1f} %",
            f"{p_glob:.1f} %",
        )
    food_table.add_row(
        "[bold green]Sous-Total Alimentaire[/bold green]",
        f"[bold]{food_items}[/bold]",
        f"[bold green]{food_total:,.2f} €[/bold green]",
        "[bold cyan]100.0 %[/bold cyan]",
        f"[bold cyan]{food_pct:.1f} %[/bold cyan]",
    )
    console.print(food_table)

    # 3. Non-Food categories breakdown
    nonfood_table = Table(
        title="🛍️ Dépenses Non-Alimentaires",
        show_header=True,
        header_style="bold white on grey23",
    )
    nonfood_table.add_column("Catégorie Non-Alimentaire", style="bold white", width=28)
    nonfood_table.add_column("Articles", justify="right", width=10)
    nonfood_table.add_column("Dépenses €", justify="right", style="bold green", width=14)
    nonfood_table.add_column("Part Non-Alim. %", justify="right", style="cyan", width=16)
    nonfood_table.add_column("Part Budget %", justify="right", style="dim cyan", width=14)

    for cat_name in sorted(CATEGORIES_NON_ALIMENTAIRES, key=lambda c: cat_spend[c], reverse=True):
        amt = cat_spend[cat_name]
        p_nf = (amt / nonfood_total * 100) if nonfood_total else 0.0
        p_glob = (amt / total_general * 100) if total_general else 0.0
        nonfood_table.add_row(
            cat_name,
            str(cat_items.get(cat_name, 0)),
            f"{amt:,.2f} €",
            f"{p_nf:.1f} %",
            f"{p_glob:.1f} %",
        )
    nonfood_table.add_row(
        "[bold]Sous-Total Non-Alimentaire[/bold]",
        f"[bold]{nonfood_items}[/bold]",
        f"[bold green]{nonfood_total:,.2f} €[/bold green]",
        "[bold cyan]100.0 %[/bold cyan]",
        f"[bold cyan]{nonfood_pct:.1f} %[/bold cyan]",
    )
    console.print(nonfood_table)

    # 4. Store breakdown
    mkt_table = Table(
        title="🏪 Répartition par Enseigne / Magasin",
        show_header=True,
        header_style="bold white on dark_green",
    )
    mkt_table.add_column("Magasin", style="bold white", width=32)
    mkt_table.add_column("Tickets", justify="right", width=10)
    mkt_table.add_column("Total Dépensé €", justify="right", style="bold green", width=16)
    mkt_table.add_column("Part %", justify="right", style="cyan", width=10)

    for m_name, amt in sorted(market_spend.items(), key=lambda x: x[1], reverse=True):
        pct = (amt / total_general * 100) if total_general else 0.0
        mkt_table.add_row(
            m_name,
            str(market_counts.get(m_name, 0)),
            f"{amt:,.2f} €",
            f"{pct:.1f} %",
        )

    console.print(mkt_table)


@app.command(name="export")
def export_excel(
    file: Path = typer.Option(
        Path("receipts.json"),
        "--file", "-f",
        help="Chemin vers le fichier JSON source",
    ),
    excel: Path = typer.Option(
        Path("depenses.xlsx"),
        "--excel", "-e",
        help="Chemin vers le fichier Excel de sortie",
    ),
) -> None:
    """Export receipts from JSON to depenses.xlsx."""
    if not file.exists():
        console.print(f"[red]Le fichier '{file}' est introuvable.[/red]")
        raise typer.Exit(code=1)

    sync_excel_file(file, excel)


@app.command()
def migrate(
    file: Path = typer.Option(
        Path("receipts.json"),
        "--file", "-f",
        help="Chemin vers le fichier JSON à normaliser et trier",
    ),
) -> None:
    """Normalize and re-sort all receipts in the JSON file."""
    if not file.exists():
        console.print(f"[red]Le fichier '{file}' est introuvable.[/red]")
        raise typer.Exit(code=1)

    backup_path = file.with_suffix(".bak.json")
    shutil.copyfile(file, backup_path)
    console.print(f"[dim]Sauvegarde de secours créée : {backup_path.name}[/dim]")

    receipts_raw = load_receipts(file)
    known_markets = [r.get("market") for r in receipts_raw if r.get("market")]
    normalized = [normalize_receipt_dict(r, existing_markets=known_markets) for r in receipts_raw]
    sorted_receipts = sort_receipts_chronologically(normalized)
    save_receipts(file, sorted_receipts)
    sync_excel_file(file, Path("depenses.xlsx"))

    console.print(
        f"[bold green]Succès :[/bold green] Migration terminée. "
        f"Total : {len(sorted_receipts)} tickets enregistrés dans [bold]{file.name}[/bold]."
    )


if __name__ == "__main__":
    app()
