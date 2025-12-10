import json
from io import BytesIO

import pandas as pd
from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.template.loader import get_template
from xhtml2pdf import pisa

from .forms import ExcelUploadForm

# ─────────────────────────────────────────
# B3-underbeteenden ↔ TQ-kompetenser
# ─────────────────────────────────────────

B3_UNDERBEHAVIORS = [
    # Affärs- och värderingsdrivet ledarskap
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag driver försäljning och bygger långsiktiga kundrelationer",
        "competencies": ["Developing relationships", "Results orientation"],
    },
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag följer upp mål och agerar snabbt när något behöver justeras",
        "competencies": ["Adaptability", "Reliability"],
    },
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag kommunicerar öppet och tydligt så att alla vet vad som gäller",
        "competencies": ["Written communication"],
    },
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag lyfter och bekräftar medarbetare för att skapa engagemang och tillit",
        "competencies": ["Engaging others"],
    },
    {
        "cluster": "Affärs- och värderingsdrivet ledarskap",
        "name": "Jag attraherar rätt kompetens och formar team som matchar kundernas behov",
        "competencies": ["Delegating", "Customer Focus"],
    },

    # Kommunicera precist och tydligt
    {
        "cluster": "Kommunicera precist och tydligt",
        "name": "Jag använder ett enkelt och tydligt språk för att undvika missförstånd",
        "competencies": ["Written communication"],
    },
    {
        "cluster": "Kommunicera precist och tydligt",
        "name": "Jag tar initiativ till samtal även när det är svårt, och förklarar syftet",
        # Excel: "Managing conflict"
        "competencies": ["Managing conflict"],
    },
    {
        "cluster": "Kommunicera precist och tydligt",
        "name": "Jag lyfter fram det som fungerar och sprider goda exempel",
        "competencies": ["Engaging others"],
    },
    {
        "cluster": "Kommunicera precist och tydligt",
        "name": "Jag leder genom dialog och bjuder in till reflektion och gemensam förståelse",
        "competencies": ["Directing others", "Organisational awareness"],
    },

    # Bygg och främja en prestationsdriven kultur
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag bygger team med kompletterande styrkor och kundfokus",
        "competencies": ["Delegating", "Customer Focus"],
    },
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag skapar utrymme för idéer och initiativ",
        # Excel: "Embracing diversity", "Optimising processes"
        "competencies": ["Embracing diversity", "Optimising processes"],
    },
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag kommunicerar öppet och tydligt",
        "competencies": ["Written communication"],
    },
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag skapar trygghet där olika perspektiv ryms",
        "competencies": ["Embracing diversity"],
    },
    {
        "cluster": "Bygg och främja en prestationsdriven kultur",
        "name": "Jag bjuder in till engagemang genom dialog och samarbete",
        # Excel: "Networking", "Driving vision and purpose"
        "competencies": ["Networking", "Driving vision and purpose"],
    },

    # Driva mot måldrivna och ambitiösa mål
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag förankrar mål så att alla förstår och känner motivation",
        "competencies": ["Engaging others", "Driving vision and purpose"],
    },
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag följer upp och stöttar för att nå förväntat resultat",
        "competencies": ["Directing others", "Supporting others"],
    },
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag samarbetar över gränser för att nå gemensamma mål",
        "competencies": ["Networking"],
    },
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag skapar tydliga arbetssätt som ger fokus och framdrift",
        # Excel: "Drive", "Optimising processes"
        "competencies": ["Drive", "Optimising processes"],
    },
    {
        "cluster": "Driva mot måldrivna och ambitiösa mål",
        "name": "Jag gör mål hanterbara och hjälper teamet att prioritera rätt",
        # Excel: "Resilience", "Organising and prioritising"
        "competencies": ["Resilience", "Organising and prioritising"],
    },

    # Rekrytera, utveckla och behåll rätt förmågor och personer
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag hittar personer som stärker teamet affärsmässigt, kulturellt och kompetensmässigt",
        "competencies": ["Delegating", "Customer Focus"],
    },
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag får medarbetare att växa genom att se potential och främja lärande",
        "competencies": ["Supporting others"],
    },
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag skapar tydlighet i rollen som konsult och kollega",
        "competencies": ["Written communication"],
    },
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag förtydligar vad som förväntas i uppdrag och kultur",
        "competencies": ["Written communication"],
    },
    {
        "cluster": "Rekrytera, utveckla och behåll rätt förmågor och personer",
        "name": "Jag bygger delaktighet genom gemenskap, respekt och goda förebilder",
        "competencies": ["Embracing diversity", "Developing relationships"],
    },
]



def upload_view(request):
    context = {}
    if request.method == "POST":
        form = ExcelUploadForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = form.cleaned_data["file"]

            # Läs Excel direkt från uppladdad fil
            df = pd.read_excel(excel_file)

            if df.empty:
                context["error"] = "Excel-filen verkar vara tom."
            else:
                row = df.iloc[0]

                first_name = row.get("First Name", "")
                last_name = row.get("Last Name", "")
                full_name = f"{first_name} {last_name}".strip()

                # Plocka ut alla kompetenskolumner
                competency_values = {}
                for col in df.columns:
                    if isinstance(col, str) and col.startswith("Competency Score:"):
                        label = col.replace("Competency Score:", "").strip()
                        label = label.replace("(STIVE)", "").strip()
                        competency_values[label] = float(row[col])

                labels = list(competency_values.keys())
                values = list(competency_values.values())

                if values:
                    avg_score = sum(values) / len(values)
                else:
                    avg_score = None

                # Plocka ut alla kompetenskolumner
                competency_values = {}
                for col in df.columns:
                    if isinstance(col, str) and col.startswith("Competency Score:"):
                        # T.ex. "Competency Score: Teamwork (STIVE)" → "Teamwork"
                        label = col.replace("Competency Score:", "").strip()
                        label = label.replace("(STIVE)", "").strip()
                        competency_values[label] = float(row[col])

                labels = list(competency_values.keys())
                values = list(competency_values.values())

                # 🔹 NYTT: beräkna B3-underbeteenden
                b3_underbehaviors = calculate_b3_underbehaviors(competency_values)

                # Enkel tolkning baserat på snitt (justera efter din logik)
                if avg_score is not None:
                    if avg_score >= 3.5:
                        summary_text = "Ditt genomsnittliga resultat ligger på en hög nivå."
                    elif avg_score >= 2.5:
                        summary_text = "Ditt genomsnittliga resultat ligger på en medelnivå."
                    else:
                        summary_text = "Ditt genomsnittliga resultat ligger på en lägre nivå."
                else:
                    summary_text = "Inga kompetensvärden hittades i filen."

                report_data = {
                                    "full_name": full_name or "Kandidaten",
                                    "avg_score": avg_score,
                                    "summary_text": summary_text,
                                    "competencies": [
                                        {"name": name, "score": val}
                                        for name, val in competency_values.items()
                                    ],
                                    "chart_labels": labels,
                                    "chart_values": values,
                                    "b3_underbehaviors": b3_underbehaviors,  # 👈 nytt
                                }

                # Spara i sessionen för PDF-vyn
                request.session["report_data"] = report_data

                context.update(report_data)

        else:
            context["error"] = "Något blev fel med filuppladdningen."
    else:
        form = ExcelUploadForm()

    # Se till att form alltid finns i context
    context.setdefault("form", form if "form" in locals() else ExcelUploadForm())

    return render(request, "reports/upload.html", context)


def report_pdf(request):
    report_data = request.session.get("report_data")
    if not report_data:
        return redirect("report_upload")

    template = get_template("reports/report_pdf.html")
    html = template.render(report_data)

    result = BytesIO()
    pdf_status = pisa.CreatePDF(html, dest=result)

    if pdf_status.err:
        return HttpResponse("Kunde inte skapa PDF just nu.", status=500)

    response = HttpResponse(result.getvalue(), content_type="application/pdf")
    response["Content-Disposition"] = 'attachment; filename="rapport.pdf"'
    return response


def calculate_b3_underbehaviors(competency_values: dict):
    """
    competency_values: dict från Excel, t.ex.
       {
         "Results Orientation": 2.8,
         "Developing Relationships": 3.1,
         "Customer Focus": 2.7,
         ...
       }

    Vi försöker matcha varje TQ-namn i B3_UNDERBEHAVIORS mot en nyckel i competency_values
    genom att jämföra lowercase + 'contains'.
    """
    results = []

    # Gör en hjälpfunktion som letar upp rätt värde
    def find_score(target_name: str):
        target = target_name.lower().strip()
        for key, value in competency_values.items():
            key_norm = str(key).lower().strip()
            # exakt match eller "contain"-match
            if key_norm == target or target in key_norm:
                return value
        return None

    for beh in B3_UNDERBEHAVIORS:
        scores = []
        missing = []

        for comp in beh["competencies"]:
            value = find_score(comp)
            if value is not None:
                scores.append(value)
            else:
                missing.append(comp)

        score = sum(scores) / len(scores) if scores else None

        results.append({
            "cluster": beh["cluster"],
            "name": beh["name"],
            "score": score,
            "missing_competencies": missing,
        })

    return results