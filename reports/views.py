import asyncio
import re
from typing import Dict, Any, List, Optional, Tuple
from urllib.parse import urlparse

import pandas as pd
from django.conf import settings
from django.http import HttpResponse
from django.shortcuts import render, redirect
from django.urls import reverse

from playwright.async_api import async_playwright

from .forms import ExcelUploadForm


# ─────────────────────────────────────────
# B3-underbeteenden ↔ TQ-kompetenser
# ─────────────────────────────────────────


UNDERBEHAVIOR_HEAVY_WEIGHT = 5.0
UNDERBEHAVIOR_NORMAL_WEIGHT = 1.0
COMPETENCY_HEAVY_WEIGHT = 5.0

B3_UNDERBEHAVIORS = [
    # ─────────────────────────────────────────
    # Affärs- och värderingsdrivet ledarskap
    # ─────────────────────────────────────────
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag driver försäljning och bygger långsiktiga kundrelationer",
        "competencies": ["Developing relationships", "Results orientation"],
        "weighted_competencies": ["Results orientation"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag följer upp mål och agerar snabbt när något behöver justeras",
        "competencies": ["Adaptability", "Reliability"],
        "weighted_competencies": ["Adaptability"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag kommunicerar öppet och tydligt så att alla vet vad som gäller",
        "competencies": ["Written communication"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag lyfter och bekräftar medarbetare för att skapa engagemang och tillit",
        "competencies": ["Engaging others"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag attraherar rätt kompetens och formar team som matchar kundernas behov",
        "competencies": ["Delegating", "Customer focus"],
        "weighted_competencies": ["Customer focus"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag stöttar teamet och visar riktning – både i medvind och motvind",
        "competencies": ["Resilience", "Supporting others"],
        "weighted_competencies": ["Supporting others"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },

    # ─────────────────────────────────────────
    # Kommunicera precist och tydligt
    # ─────────────────────────────────────────

    {
        "cluster": "Kommunicera precist och tydligt",
        "name": "Jag använder ett enkelt och tydligt språk för att undvika missförstånd",
        "competencies": ["Written communication"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Kommunicera precist och tydligt",
        "name": "Jag tar initiativ till samtal även när det är svårt, och förklarar syftet",
        "competencies": ["Managing conflicts"],
        "weighted_competencies": ["Managing conflicts"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },
    {
        "cluster": "Kommunicera precist och tydligt",
        "name": "Jag lyfter fram det som fungerar och sprider goda exempel",
        "competencies": ["Engaging others"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Kommunicera precist och tydligt",
        "name": "Jag leder genom dialog och bjuder in till reflektion och gemensam förståelse",
        "competencies": ["Directing others", "Organisational awareness"],
        "weighted_competencies": ["Organisational awareness"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },
    {
        "cluster": "Kommunicera precist och tydligt",
        "name": "Jag kommunicerar med respekt, mod och tydlighet för att skapa trygghet",
        "competencies": ["Interpersonal communication", "Dealing with ambiguity"],
        "weighted_competencies": ["Interpersonal communication"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },

    # ─────────────────────────────────────────
    # Bygg och främja en prestationsdriven kultur
    # ─────────────────────────────────────────
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag bygger team med kompletterande styrkor och kundfokus",
        "competencies": ["Delegating", "Customer focus"],
        "weighted_competencies": ["Delegating"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag skapar utrymme för idéer och initiativ",
        "competencies": ["Embracing diversity", "Optimizing processes"],
        "weighted_competencies": ["Embracing diversity"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag kommunicerar öppet och tydligt",
        "competencies": ["Written communication"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag skapar trygghet där olika perspektiv ryms",
        "competencies": ["Embracing diversity"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag bjuder in till engagemang genom dialog och samarbete",
        "competencies": ["Networking", "Driving vision and purpose"],
        "weighted_competencies": ["Networking"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag stärker kulturen genom att visa att vi står tillsammans – både i med- och motgång",
        "competencies": ["Driving vision and purpose", "Results orientation"],
        "weighted_competencies": ["Results orientation"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },

    # ─────────────────────────────────────────
    # Driva mot måldrivna och ambitiösa mål
    # ─────────────────────────────────────────
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag förankrar mål så att alla förstår och känner motivation",
        "competencies": ["Engaging others", "Driving vision and purpose"],
        "weighted_competencies": ["Engaging others"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag följer upp och stöttar för att nå förväntat resultat",
        "competencies": ["Directing others", "Supporting others"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag samarbetar över gränser för att nå gemensamma mål",
        "competencies": ["Networking"],
        "weighted_competencies": ["Networking"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag skapar tydliga arbetssätt som ger fokus och framdrift",
        "competencies": ["Drive", "Optimizing processes"],
        "weighted_competencies": ["Optimizing processes"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag gör mål hanterbara och hjälper teamet att prioritera rätt",
        "competencies": ["Resilience", "Organizing and prioritizing"],
        "weighted_competencies": ["Organizing and prioritizing"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag ser till helheten och agerar långsiktigt, även när det är kortsiktigt utmanande.",
        "competencies": ["Strategic focus", "Drive"],
        "weighted_competencies": ["Strategic focus"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },

    # ─────────────────────────────────────────
    # Rekrytera, utveckla och behåll rätt förmågor och personer
    # ─────────────────────────────────────────
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag hittar personer som stärker teamet affärsmässigt, kulturellt och kompetensmässigt",
        "competencies": ["Delegating", "Customer focus"],
        "weighted_competencies": ["Delegating"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag får medarbetare att växa genom att se potential och främja lärande",
        "competencies": ["Supporting others"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag skapar tydlighet i rollen som konsult och kollega",
        "competencies": ["Written communication"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag förtydligar vad som förväntas i uppdrag och kultur",
        "competencies": ["Written communication"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag bygger delaktighet genom gemenskap, respekt och goda förebilder",
        "competencies": ["Embracing diversity", "Developing relationships"],
        "weighted_competencies": ["Embracing diversity"],
        "weight": UNDERBEHAVIOR_NORMAL_WEIGHT,
    },
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag ser till att vi har rätt personer på bussen och är modig att fatta beslut när en roll inte är rätt för individen eller teamet",
        "competencies": ["Decisiveness", "Organisational awareness"],
        "weighted_competencies": ["Decisiveness"],
        "weight": UNDERBEHAVIOR_HEAVY_WEIGHT,
    },
]


COMP_ALIASES = {
    "customer focus": ["customer focus", "customerfocus"],
    "managing conflicts": ["managing conflict", "managing conflicts"],
    "optimizing processes": ["optimising processes", "optimizing processes"],
    "organizing and prioritizing": ["organising and prioritising", "organizing and prioritizing"],
    "driving vision and purpose": ["driving vision & purpose", "driving vision and purpose", "driving vision purpose"],
    "developing relationships": ["developing relationships", "developing relationship"],
    "results orientation": ["results orientation", "result orientation"],
}



# ─────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────

def _extract_competency_values(df: pd.DataFrame) -> Dict[str, float]:
    """Plockar ut kompetenser från första raden i Excel och returnerar {label: score}."""
    row = df.iloc[0]
    competency_values: Dict[str, float] = {}

    for col in df.columns:
        if isinstance(col, str) and col.startswith("Competency Score:"):
            label = col.replace("Competency Score:", "").strip()
            label = label.replace("(STIVE)", "").strip()

            try:
                competency_values[label] = float(row[col])
            except (TypeError, ValueError):
                continue

    return competency_values

def _build_cluster_calc_line(
    items_used: List[Dict[str, Any]],
    result: Optional[float],
    denominator: Optional[float] = None
) -> str:
    """
    Ex:
    - gamla: (UB1 3.03×2 + UB2 1.42×1 + ...) / (2+1+...) = 2.78
    - nya:    (UB1 3.03×2 + UB2 1.42×1 + ...) / 6 = 3.76
    """
    if not items_used or result is None:
        return "Ingen uträkning (saknar underbeteenden)."

    left = " + ".join([
        f'{x["underbehavior"]} {_fmt(x["score"])}×{_fmt(x["weight"],0)}'
        for x in items_used
    ])

    if denominator is None:
        # fallback till gamla beteendet om du någonsin vill återanvända funktionen
        wsum = " + ".join([f'{_fmt(x["weight"],0)}' for x in items_used])
        return f"({left}) / ({wsum}) = {_fmt(result)}"

    return f"({left}) / {_fmt(float(denominator), 0)} = {_fmt(result)}"

def round_to_half(value: Optional[float]) -> Optional[float]:
    """
    1–5-skala → avrundar till 0.5-steg
    Ex: 2.74 → 2.5, 2.76 → 3.0
    """
    if value is None:
        return None
    return round(value * 2) / 2

def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("&", "and")
    s = re.sub(r"\s+", " ", s)
    return s

def _build_lookup(competency_values: Dict[str, float]) -> Dict[str, float]:
    """
    Bygger en normaliserad lookup av Excel-kompetenser, så vi kan matcha stabilt.
    """
    lookup: Dict[str, float] = {}
    for k, v in competency_values.items():
        lookup[_norm(str(k))] = v
    return lookup

def _find_score(lookup: Dict[str, float], target: str) -> Optional[float]:
    t = _norm(target)

    # 1) Direkt match
    if t in lookup:
        return lookup[t]

    # 2) Alias-lista
    for canon, variants in COMP_ALIASES.items():
        if t == canon:
            for var in variants:
                nv = _norm(var)
                if nv in lookup:
                    return lookup[nv]

    # 3) “contains” fallback (snäll men kan rädda små skillnader)
    for k, v in lookup.items():
        if t in k or k in t:
            return v

    return None

def _simple_average(values: List[float]) -> Optional[float]:
    if not values:
        return None
    return sum(values) / len(values)

def _weighted_average(pairs: List[Tuple[float, float]]) -> Optional[float]:
    """
    pairs = [(score, weight), ...]
    returnerar viktat medel (normaliserat) => håller sig i 1–5
    """
    if not pairs:
        return None
    w_sum = sum(score * w for score, w in pairs)
    w_tot = sum(w for _, w in pairs)
    if w_tot == 0:
        return None
    return w_sum / w_tot

def _fmt(n: Optional[float], decimals: int = 2) -> str:
    if n is None:
        return "—"
    return f"{n:.{decimals}f}"

def _build_under_calc_line(comp_debug: List[Dict[str, Any]], result: Optional[float]) -> str:
    """
    Ex: (Delegating 3.58×2 + Customer focus 1.93×1) / (2+1) = 3.03
    """
    if not comp_debug or result is None:
        return "Ingen uträkning (saknar värden)."

    left = " + ".join([f'{c["competency"]} {_fmt(c["score"])}×{_fmt(c["weight"],0)}' for c in comp_debug])
    wsum = " + ".join([f'{_fmt(c["weight"],0)}' for c in comp_debug])

    return f"({left}) / ({wsum}) = {_fmt(result)}"

def calculate_b3_underbehaviors_and_clusters(
    competency_values: Dict[str, float],
    b3_underbehaviors_def: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], str, List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Returnerar:
    - underbehaviors (med score_5 = VIKTAT)
    - clusters (med score_5 = VIKTAT)
    - calc_explain_text
    - under_compare_rows (tabell: viktat vs oviktat per underbeteende)
    - cluster_compare_rows (tabell: viktat vs oviktat per huvudbeteende)
    """

    lookup = _build_lookup(competency_values)

    underbehaviors: List[Dict[str, Any]] = []
    cluster_items: Dict[str, List[Dict[str, Any]]] = {}

    # -------- 0) Klusterordning (för att slippa NameError) --------
    cluster_order: List[str] = []
    for beh in b3_underbehaviors_def:
        if beh["cluster"] not in cluster_order:
            cluster_order.append(beh["cluster"])

    # -------- 1) Beräkna varje underbeteende --------
    for beh in b3_underbehaviors_def:
        comp_debug: List[Dict[str, Any]] = []
        comp_pairs: List[Tuple[float, float]] = []
        comp_values_unweighted: List[float] = []
        missing: List[str] = []

        weighted_comps = {_norm(c) for c in beh.get("weighted_competencies", [])}
        unknown = weighted_comps - {_norm(c) for c in beh["competencies"]}
        if unknown:
            raise ValueError(
                f"Fel i mapping: weighted_competencies innehåller {unknown} "
                f"som inte finns i competencies för underbeteende: {beh['name']}"
            )

        for comp in beh["competencies"]:
            v = _find_score(lookup, comp)
            if v is None:
                missing.append(comp)
                continue

            v = float(v)
            comp_values_unweighted.append(v)

            comp_weight = COMPETENCY_HEAVY_WEIGHT if _norm(comp) in weighted_comps else 1.0
            comp_pairs.append((v, float(comp_weight)))

            comp_debug.append({
                "competency": comp,
                "score": v,
                "weight": float(comp_weight),
                "weighted": v * float(comp_weight),
            })

        # VIKTAT: dividera med ANTAL KOMPETENSER
        under_score_weighted = _weighted_average(comp_pairs)

        # OVIKTAT: vanligt snitt
        under_score_unweighted = _simple_average(comp_values_unweighted)

        # Pluppar (viktat spår = det som visas i UI)
        under_half = round_to_half(under_score_weighted)
        under_half_steps = int(under_half * 2) if under_half is not None else None
        under_rounded = int(round(under_score_weighted)) if under_score_weighted is not None else None

        under_weight = float(beh.get("weight", 1.0))

        # “Uträkning”-rad som visar / N (antal kompetenser)
        if comp_debug and under_score_weighted is not None:
            left = " + ".join([f'{c["competency"]} {_fmt(c["score"])}×{_fmt(c["weight"],0)}' for c in comp_debug])
            wsum = " + ".join([f'{_fmt(c["weight"],0)}' for c in comp_debug])
            human_under_line = f"({left}) / ({wsum}) = {_fmt(under_score_weighted)}"
        else:
            human_under_line = "Ingen uträkning (saknar värden)."

        item = {
            "cluster": beh["cluster"],
            "name": beh["name"],
            "competencies": beh["competencies"],

            # Båda spår
            "score_5_weighted": under_score_weighted,
            "score_5_unweighted": under_score_unweighted,

            # UI ska använda viktad (som ni vill)
            "score_5": under_score_weighted,
            "score_5_half": under_half,
            "score_5_half_steps": under_half_steps,
            "score_5_rounded": under_rounded,

            "weight": under_weight,
            "missing": missing,

            "calc_debug": {
                "formula_weighted": "(Σ(score×vikt)) / (Σ(vikter)",
                "formula_unweighted": "(Σ(score)) / N (antal kompetenser)",
                "components": comp_debug,
                "result_weighted": under_score_weighted,
                "result_unweighted": under_score_unweighted,
            },
            "calc_human": {
                "title": "Uträkning (underbeteende)",
                "line": human_under_line,
                "note": "Viktning påverkar täljaren, men vi dividerar med antal kompetenser (inte summa vikter).",
            },
        }

        underbehaviors.append(item)
        cluster_items.setdefault(beh["cluster"], []).append(item)

    # -------- 2) Beräkna varje huvudbeteende (kluster) --------
    clusters: List[Dict[str, Any]] = []

    for cluster_name in cluster_order:
        items = [x for x in cluster_items.get(cluster_name, []) if x.get("score_5") is not None]

        # VIKTAT KLUSTER: Σ(under_score_weighted × under_weight) / N (antal underbeteenden)
        if items:
            pairs = [
                (x["score_5_weighted"], x["weight"])
                for x in items
                if x.get("score_5_weighted") is not None
            ]
            cluster_score_weighted = _weighted_average(pairs)

            weighted_sum = sum(score * w for score, w in pairs)
            weight_total = sum(w for _, w in pairs)

            # OVIKTAT KLUSTER: snitt av under_score_unweighted (ingen UB-vikt, inga comp-vikter)
            unweighted_vals = [x["score_5_unweighted"] for x in items if x.get("score_5_unweighted") is not None]
            cluster_score_unweighted = (sum(unweighted_vals) / len(unweighted_vals)) if unweighted_vals else None
        else:
            weighted_sum = 0
            denom = 0
            cluster_score_weighted = None
            cluster_score_unweighted = None

        cluster_half = round_to_half(cluster_score_weighted)
        cluster_half_steps = int(cluster_half * 2) if cluster_half is not None else None
        cluster_rounded = int(round(cluster_score_weighted)) if cluster_score_weighted is not None else None

        items_used = [
            {
                "underbehavior": x["name"],
                "score": x.get("score_5_weighted"),
                "weight": x["weight"],
                "weighted": (x.get("score_5_weighted") * x["weight"]) if x.get("score_5_weighted") is not None else None,
            }
            for x in items
            if x.get("score_5_weighted") is not None
        ]

        human_cluster_line = _build_cluster_calc_line(items_used, cluster_score_weighted, None)

        clusters.append({
            "name": cluster_name,

            # UI (viktat)
            "score_5": cluster_score_weighted,
            "score_5_half": cluster_half,
            "score_5_half_steps": cluster_half_steps,
            "score_5_rounded": cluster_rounded,

            # jämförelse
            "score_5_weighted": cluster_score_weighted,
            "score_5_unweighted": cluster_score_unweighted,

            "calc_debug": {
                "formula": "(Σ(under_score×under_vikt)) / (Σ(under_vikter))",
                "weighted_sum": weighted_sum,
                "denominator": weight_total,
                "result_weighted": cluster_score_weighted,
                "result_unweighted": cluster_score_unweighted,
                "items_used": items_used,
            },
            "calc_human": {
                "title": "Uträkning (huvudbeteende)",
                "line": human_cluster_line,
                "note": "Huvudbeteenden beräknas som Σ(under_score×under_vikt) / N (antal underbeteenden).",
            },
        })

    # -------- 3) Text i rapporten --------
    calc_explain_text = (
        "Beräkningsprincip: Underbeteenden beräknas som ett viktat medelvärde "
        "(Σ(score×vikt)) / (Σ(vikter)). "
        "Huvudbeteenden (kluster) beräknas på samma sätt utifrån underbeteendenas poäng "
        "och deras vikt: (Σ(under_score×under_vikt)) / (Σ(under_vikter))."
    )

    # -------- 4) Jämförelse-tabeller (viktat vs oviktat) --------
    under_compare_rows = [
        {
            "cluster": u["cluster"],
            "name": u["name"],
            "weighted": u.get("score_5_weighted"),
            "unweighted": u.get("score_5_unweighted"),
            "diff": (
                u["score_5_weighted"] - u["score_5_unweighted"]
                if u.get("score_5_weighted") is not None and u.get("score_5_unweighted") is not None
                else None
            ),
        }
        for u in underbehaviors
    ]

    cluster_compare_rows = [
        {
            "name": c["name"],
            "weighted": c.get("score_5_weighted"),
            "unweighted": c.get("score_5_unweighted"),
            "diff": (
                c["score_5_weighted"] - c["score_5_unweighted"]
                if c.get("score_5_weighted") is not None and c.get("score_5_unweighted") is not None
                else None
            ),
        }
        for c in clusters
    ]

    return underbehaviors, clusters, calc_explain_text, under_compare_rows, cluster_compare_rows



async def _render_pdf_async(url: str, cookie_name: str, cookie_value: Optional[str]) -> bytes:
    """Renderar en URL till PDF via Playwright, med session-cookie så vi inte blir redirectade."""
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox"] if not settings.DEBUG else []
        )
        context = await browser.new_context()

        if cookie_value:
            parsed = urlparse(url)
            await context.add_cookies([{
                "name": cookie_name,
                "value": cookie_value,
                "domain": parsed.hostname,  # ex: 127.0.0.1
                "path": "/",
            }])

        page = await context.new_page()
        await page.goto(url, wait_until="domcontentloaded")

        # Liten buffert så att ev. chart hinner ritas
        await page.wait_for_timeout(300)

        pdf_bytes = await page.pdf(
            format="A4",
            print_background=True,
            margin={"top": "14mm", "right": "12mm", "bottom": "14mm", "left": "12mm"},
        )

        await browser.close()
        return pdf_bytes


# ─────────────────────────────────────────
# Views
# ─────────────────────────────────────────

def upload_view(request):
    """
    En sida: upload + rapport under.
    Sparar report_data i session.
    """
    context: Dict[str, Any] = {"form": ExcelUploadForm()}

    if request.method == "POST":
        form = ExcelUploadForm(request.POST, request.FILES)
        context["form"] = form

        if not form.is_valid():
            context["error"] = "Något blev fel med filuppladdningen."
            return render(request, "reports/upload.html", context)

        excel_file = form.cleaned_data["file"]
        df = pd.read_excel(excel_file)

        if df.empty:
            context["error"] = "Excel-filen verkar vara tom."
            return render(request, "reports/upload.html", context)

        row = df.iloc[0]
        first_name = row.get("First Name", "")
        last_name = row.get("Last Name", "")
        full_name = f"{first_name} {last_name}".strip() or "Kandidaten"

        competency_values = _extract_competency_values(df)
        labels = list(competency_values.keys())
        values = list(competency_values.values())
        avg_score = (sum(values) / len(values)) if values else None

        competencies_list = []
        for k, v in competency_values.items():
            competencies_list.append({
                "name": k,
                "score_5": float(v),
                "score_5_rounded": int(round(float(v))),  # ✅ för pluppar
            })

        if avg_score is not None:
            if avg_score >= 3.5:
                summary_text = "Ditt genomsnittliga resultat ligger på en hög nivå."
            elif avg_score >= 2.5:
                summary_text = "Ditt genomsnittliga resultat ligger på en medelnivå."
            else:
                summary_text = "Ditt genomsnittliga resultat ligger på en lägre nivå."
        else:
            summary_text = "Inga kompetensvärden hittades i filen."

        (
            b3_underbehaviors,
            b3_clusters,
            calc_explain_text,
            under_compare_rows,
            cluster_compare_rows,
        ) = calculate_b3_underbehaviors_and_clusters(
            competency_values,
            B3_UNDERBEHAVIORS
        )

        report_data = {
            "full_name": full_name,
            "avg_score": avg_score,
            "summary_text": summary_text,
            "competencies": competencies_list, 
            "chart_labels": labels,
            "chart_values": values,
            "b3_underbehaviors": b3_underbehaviors,
            "b3_clusters": b3_clusters,
            "calc_explain_text": calc_explain_text,
            "under_compare_rows": under_compare_rows,
            "cluster_compare_rows": cluster_compare_rows,
        }

        print("B3 DEBUG missing count:", sum(len(x["missing"]) for x in b3_underbehaviors))
        print("Example missing:", [x for x in b3_underbehaviors if x["missing"]][:2])

        request.session["report_data"] = report_data
        context.update(report_data)

    return render(request, "reports/upload.html", context)


def report_pdf_page(request):
    """
    Ren HTML-sida för PDF (utan upload-form).
    Denna renderas av Playwright.
    """
    report_data = request.session.get("report_data")
    if not report_data:
        return redirect("report_upload")
    return render(request, "reports/report_pdf.html", report_data)


def report_pdf_download(request):
    """
    Laddar ner PDF för ENDAST rapporten (printar report_pdf_page).
    Viktigt: skickar med session-cookie till Playwright.
    """
    report_data = request.session.get("report_data")
    if not report_data:
        return redirect("report_upload")

    url = request.build_absolute_uri(reverse("report_pdf_page"))

    cookie_name = settings.SESSION_COOKIE_NAME
    cookie_value = request.COOKIES.get(cookie_name)

    pdf_bytes = asyncio.run(_render_pdf_async(url, cookie_name, cookie_value))

    response = HttpResponse(pdf_bytes, content_type="application/pdf")
    response["Content-Disposition"] = 'attachment; filename="rapport.pdf"'
    return response
