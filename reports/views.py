import asyncio
from typing import Dict, Any, List, Optional
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

B3_UNDERBEHAVIORS = [
    # Affärs- och värderingsdrivet ledarskap
    {"cluster": "Affärs- och värderingsdrivet ledarskap",
     "name": "Jag driver försäljning och bygger långsiktiga kundrelationer",
     "competencies": ["Developing relationships", "Results orientation"]},

    {"cluster": "Affärs- och värderingsdrivet ledarskap",
     "name": "Jag följer upp mål och agerar snabbt när något behöver justeras",
     "competencies": ["Adaptability", "Reliability"]},

    {"cluster": "Affärs- och värderingsdrivet ledarskap",
     "name": "Jag kommunicerar öppet och tydligt så att alla vet vad som gäller",
     "competencies": ["Written communication"]},

    {"cluster": "Affärs- och värderingsdrivet ledarskap",
     "name": "Jag lyfter och bekräftar medarbetare för att skapa engagemang och tillit",
     "competencies": ["Engaging others"]},

    {"cluster": "Affärs- och värderingsdrivet ledarskap",
     "name": "Jag attraherar rätt kompetens och formar team som matchar kundernas behov",
     "competencies": ["Delegating", "Customer Focus"]},

    # Kommunicera precist och tydligt
    {"cluster": "Kommunicera precist och tydligt",
     "name": "Jag använder ett enkelt och tydligt språk för att undvika missförstånd",
     "competencies": ["Written communication"]},

    {"cluster": "Kommunicera precist och tydligt",
     "name": "Jag tar initiativ till samtal även när det är svårt, och förklarar syftet",
     "competencies": ["Managing conflict"]},

    {"cluster": "Kommunicera precist och tydligt",
     "name": "Jag lyfter fram det som fungerar och sprider goda exempel",
     "competencies": ["Engaging others"]},

    {"cluster": "Kommunicera precist och tydligt",
     "name": "Jag leder genom dialog och bjuder in till reflektion och gemensam förståelse",
     "competencies": ["Directing others", "Organisational awareness"]},

    # Bygg och främja en prestationsdriven kultur
    {"cluster": "Bygg och främja en prestationsdriven kultur",
     "name": "Jag bygger team med kompletterande styrkor och kundfokus",
     "competencies": ["Delegating", "Customer Focus"]},

    {"cluster": "Bygg och främja en prestationsdriven kultur",
     "name": "Jag skapar utrymme för idéer och initiativ",
     "competencies": ["Embracing diversity", "Optimising processes"]},

    {"cluster": "Bygg och främja en prestationsdriven kultur",
     "name": "Jag kommunicerar öppet och tydligt",
     "competencies": ["Written communication"]},

    {"cluster": "Bygg och främja en prestationsdriven kultur",
     "name": "Jag skapar trygghet där olika perspektiv ryms",
     "competencies": ["Embracing diversity"]},

    {"cluster": "Bygg och främja en prestationsdriven kultur",
     "name": "Jag bjuder in till engagemang genom dialog och samarbete",
     "competencies": ["Networking", "Driving vision and purpose"]},

    # Driva mot måldrivna och ambitiösa mål
    {"cluster": "Driva mot måldrivna och ambitiösa mål",
     "name": "Jag förankrar mål så att alla förstår och känner motivation",
     "competencies": ["Engaging others", "Driving vision and purpose"]},

    {"cluster": "Driva mot måldrivna och ambitiösa mål",
     "name": "Jag följer upp och stöttar för att nå förväntat resultat",
     "competencies": ["Directing others", "Supporting others"]},

    {"cluster": "Driva mot måldrivna och ambitiösa mål",
     "name": "Jag samarbetar över gränser för att nå gemensamma mål",
     "competencies": ["Networking"]},

    {"cluster": "Driva mot måldrivna och ambitiösa mål",
     "name": "Jag skapar tydliga arbetssätt som ger fokus och framdrift",
     "competencies": ["Drive", "Optimising processes"]},

    {"cluster": "Driva mot måldrivna och ambitiösa mål",
     "name": "Jag gör mål hanterbara och hjälper teamet att prioritera rätt",
     "competencies": ["Resilience", "Organising and prioritising"]},

    # Rekrytera, utveckla och behåll rätt förmågor och personer
    {"cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
     "name": "Jag hittar personer som stärker teamet affärsmässigt, kulturellt och kompetensmässigt",
     "competencies": ["Delegating", "Customer Focus"]},

    {"cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
     "name": "Jag får medarbetare att växa genom att se potential och främja lärande",
     "competencies": ["Supporting others"]},

    {"cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
     "name": "Jag skapar tydlighet i rollen som konsult och kollega",
     "competencies": ["Written communication"]},

    {"cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
     "name": "Jag förtydligar vad som förväntas i uppdrag och kultur",
     "competencies": ["Written communication"]},

    {"cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
     "name": "Jag bygger delaktighet genom gemenskap, respekt och goda förebilder",
     "competencies": ["Embracing diversity", "Developing relationships"]},
]


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


def calculate_b3_underbehaviors(competency_values: Dict[str, float]) -> List[Dict[str, Any]]:
    """Matchar robust (lowercase + contains) för att tåla små variationer i namn."""
    def find_score(target_name: str) -> Optional[float]:
        target = target_name.lower().strip()
        for key, value in competency_values.items():
            key_norm = str(key).lower().strip()
            if key_norm == target or target in key_norm:
                return value
        return None

    results: List[Dict[str, Any]] = []
    for beh in B3_UNDERBEHAVIORS:
        scores: List[float] = []

        for comp in beh["competencies"]:
            v = find_score(comp)
            if v is not None:
                scores.append(v)

        score = (sum(scores) / len(scores)) if scores else None

        results.append({
            "cluster": beh["cluster"],
            "name": beh["name"],
            "score": score,
            # Debugfält (visas inte i UI om du inte vill):
            "missing_competencies": [c for c in beh["competencies"] if find_score(c) is None],
        })

    return results


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

        if avg_score is not None:
            if avg_score >= 3.5:
                summary_text = "Ditt genomsnittliga resultat ligger på en hög nivå."
            elif avg_score >= 2.5:
                summary_text = "Ditt genomsnittliga resultat ligger på en medelnivå."
            else:
                summary_text = "Ditt genomsnittliga resultat ligger på en lägre nivå."
        else:
            summary_text = "Inga kompetensvärden hittades i filen."

        b3_underbehaviors = calculate_b3_underbehaviors(competency_values)

        report_data = {
            "full_name": full_name,
            "avg_score": avg_score,
            "summary_text": summary_text,
            "competencies": [{"name": k, "score": v} for k, v in competency_values.items()],
            "chart_labels": labels,
            "chart_values": values,
            "b3_underbehaviors": b3_underbehaviors,
        }

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
