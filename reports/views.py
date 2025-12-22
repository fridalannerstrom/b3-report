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


UNDERBEHAVIOR_HEAVY_WEIGHT = 2.0
UNDERBEHAVIOR_NORMAL_WEIGHT = 1.0
COMPETENCY_HEAVY_WEIGHT = 1.0

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
) -> Tuple[
    List[Dict[str, Any]],  # underbehaviors
    List[Dict[str, Any]],  # clusters
    str,                   # calc_explain_text
    List[Dict[str, Any]],  # under_compare_rows
    List[Dict[str, Any]],  # cluster_compare_rows
    Dict[str, Any],        # insights
]:
    """
    LOGIK:
    - Underbeteende (1–5): vanligt medel av kopplade kompetenser (ingen kompetensviktning).
    - Huvudbeteende/kluster: total_score = Σ(under_score × under_weight), weight 1 eller 2.
    - UI:
        pct_total        = ratio 0..1  (perfekt för donut)
        pct_total_percent= 0..100      (perfekt för text/radar om du vill)
    - mapped_competencies: [{name, score}] för UI-kompetensraderna (bar-grafen).
    """

    lookup = _build_lookup(competency_values)

    underbehaviors: List[Dict[str, Any]] = []
    cluster_items: Dict[str, List[Dict[str, Any]]] = {}

    # Klusterordning (behåller den ordning du definierat i B3_UNDERBEHAVIORS)
    cluster_order: List[str] = []
    for beh in b3_underbehaviors_def:
        c = beh.get("cluster")
        if c and c not in cluster_order:
            cluster_order.append(c)

    # 1) Underbeteenden
    for beh in b3_underbehaviors_def:
        comp_debug: List[Dict[str, Any]] = []
        comp_values: List[float] = []
        missing: List[str] = []

        comps = beh.get("competencies", [])

        # Hämta scores för beräkningen (med din alias-matchning)
        for comp in comps:
            v = _find_score(lookup, comp)
            if v is None:
                missing.append(comp)
                continue
            v = float(v)
            comp_values.append(v)

            comp_debug.append({
                "competency": comp,
                "score": v,
                "weight": 1.0,
                "weighted": v,
            })

        # ✅ Detta är NYCKELN: bygg en lista som UI kan loopa över
        # Den innehåller ALLA kompetenser (även de som saknar score => None)
        mapped_competencies: List[Dict[str, Any]] = []
        for comp in comps:
            v = _find_score(lookup, comp)
            mapped_competencies.append({
                "name": comp,
                "score": float(v) if v is not None else None,
            })

        under_score = _simple_average(comp_values)

        under_half = round_to_half(under_score)
        under_half_steps = int(under_half * 2) if under_half is not None else None
        under_rounded = int(round(under_score)) if under_score is not None else None

        raw_weight = float(beh.get("weight", 1.0))
        under_weight = 2.0 if raw_weight >= 2.0 else 1.0

        if comp_debug and under_score is not None:
            left = " + ".join([f'{c["competency"]} {_fmt(c["score"])}' for c in comp_debug])
            n = len(comp_debug)
            human_under_line = f"({left}) / {n} = {_fmt(under_score)}"
        else:
            human_under_line = "Ingen uträkning (saknar värden)."

        item = {
            "cluster": beh.get("cluster"),
            "name": beh.get("name"),
            "competencies": comps,

            # ✅ här finns det du vill använda i template för bar-grafen
            "mapped_competencies": mapped_competencies,

            "score_5": under_score,
            "score_5_half": under_half,
            "score_5_half_steps": under_half_steps,
            "score_5_rounded": under_rounded,

            "weight": under_weight,
            "missing": missing,

            # debug kan du ta bort sen, men behåller så länge
            "calc_debug": {
                "formula": "(Σ(score)) / N (antal kompetenser)",
                "components": comp_debug,
                "result": under_score,
            },
            "calc_human": {
                "title": "Uträkning (underbeteende)",
                "line": human_under_line,
                "note": "Underbeteenden beräknas som ett vanligt medelvärde av kompetenser (ingen kompetensviktning).",
            },
        }

        underbehaviors.append(item)
        cluster_items.setdefault(item["cluster"], []).append(item)

    # 2) Kluster (huvudbeteenden)
    clusters: List[Dict[str, Any]] = []

    for cluster_name in cluster_order:
        items = [x for x in cluster_items.get(cluster_name, []) if x.get("score_5") is not None]

        total_score: Optional[float] = None
        max_total: Optional[float] = None
        pct_ratio: Optional[float] = None       # 0..1
        pct_percent: Optional[float] = None     # 0..100

        if items:
            total_score = sum((x["score_5"] * x["weight"]) for x in items if x.get("score_5") is not None)
            max_total = sum((5.0 * x["weight"]) for x in items)

            if max_total and max_total > 0:
                pct_ratio = total_score / max_total
                pct_percent = pct_ratio * 100.0

        items_used = [
            {
                "underbehavior": x["name"],
                "score": x.get("score_5"),
                "weight": x["weight"],
                "weighted": (x.get("score_5") * x["weight"]) if x.get("score_5") is not None else None,
            }
            for x in items
            if x.get("score_5") is not None
        ]

        if total_score is not None and items_used:
            left = " + ".join([f'{x["underbehavior"]} {_fmt(x["score"])}×{_fmt(x["weight"],0)}' for x in items_used])
            human_cluster_line = f"{left} = {_fmt(total_score)}"
        else:
            human_cluster_line = "Ingen uträkning (saknar underbeteenden)."

        clusters.append({
            "name": cluster_name,

            "total_score": total_score,
            "max_total": max_total,

            # ✅ för UI
            "pct_total": pct_percent,                 # 0..100
            "pct_total_ratio": pct_ratio,             # 0..1 (om du vill ha kvar)

            # bakåtkomp
            "score_5": total_score,

            "calc_debug": {
                "formula": "Σ(under_score × under_vikt)",
                "result": total_score,
                "max_total": max_total,
                "pct_total": pct_percent,
                "pct_total_ratio": pct_ratio,
                "items_used": items_used,
            },
            "calc_human": {
                "title": "Uträkning (huvudbeteende)",
                "line": human_cluster_line,
                "note": "Huvudbeteenden visas som totalpoäng: summan av underbeteenden där vissa viktas ×2.",
            },
        })

    calc_explain_text = (
        "Beräkningsprincip:\n"
        "- Underbeteenden (1–5) beräknas som ett vanligt medelvärde av kopplade kompetenser (ingen kompetensviktning).\n"
        "- Huvudbeteenden beräknas som totalpoäng: Σ(underbeteende-poäng × underbeteende-vikt), där vikt är 1 eller 2."
    )

    under_compare_rows = [
        {"cluster": u["cluster"], "name": u["name"], "weighted": u.get("score_5"), "unweighted": u.get("score_5"), "diff": 0.0}
        for u in underbehaviors if u.get("score_5") is not None
    ]
    cluster_compare_rows = [
        {"name": c["name"], "weighted": c.get("total_score"), "unweighted": None, "diff": None}
        for c in clusters
    ]

    clusters_with_score = [c for c in clusters if c.get("total_score") is not None]
    most_natural = max(clusters_with_score, key=lambda c: c["total_score"]) if clusters_with_score else None
    needs_development = min(clusters_with_score, key=lambda c: c["total_score"]) if clusters_with_score else None

    valid_under = [u for u in underbehaviors if u.get("score_5") is not None]
    top_3_energy = sorted(valid_under, key=lambda u: u["score_5"], reverse=True)[:3]
    bottom_3_energy = sorted(valid_under, key=lambda u: u["score_5"])[:3]

    insights = {
        "most_natural": most_natural,
        "needs_development": needs_development,
        "top_energy": top_3_energy,
        "low_energy": bottom_3_energy,
    }

    return underbehaviors, clusters, calc_explain_text, under_compare_rows, cluster_compare_rows, insights




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
        await page.set_viewport_size({"width": 1440, "height": 900})
        await page.emulate_media(media="screen")  # superviktigt: använd screen istället för print

        await page.goto(url, wait_until="networkidle")

        pdf_bytes = await page.pdf(
            format="A4",
            print_background=True,
            margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
            prefer_css_page_size=True,
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

        # (valfritt att ha kvar) enkel lookup om du vill, men du behöver inte för underbeteenden nu
        def _norm_key(s: str) -> str:
            return (s or "").strip().lower().replace("&", "and")

        competency_lookup = {_norm_key(k): float(v) for k, v in competency_values.items()}

        labels = list(competency_values.keys())
        values = list(competency_values.values())
        avg_score = (sum(values) / len(values)) if values else None

        competencies_list = [{
            "name": k,
            "score_5": float(v),
            "score_5_rounded": int(round(float(v))),
        } for k, v in competency_values.items()]

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
            insights,
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

            "competency_lookup": competency_lookup,  # om du vill använda senare

            "b3_underbehaviors": b3_underbehaviors,
            "b3_clusters": b3_clusters,
            "insights": insights,

            "calc_explain_text": calc_explain_text,
            "under_compare_rows": under_compare_rows,
            "cluster_compare_rows": cluster_compare_rows,
        }

        # Radar chart:
        # Om du vill ha 0..100 i radar:
        radar_labels = [c["name"] for c in b3_clusters]
        radar_values = [
            float(c["pct_total"]) if c.get("pct_total") is not None else 0.0
            for c in b3_clusters
        ]

        report_data.update({
            "radar_labels": radar_labels,
            "radar_values": radar_values,
        })

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

