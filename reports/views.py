import json
from io import BytesIO

import pandas as pd
from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.template.loader import get_template
from xhtml2pdf import pisa

from .forms import ExcelUploadForm


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
                        # T.ex. "Competency Score: Teamwork (STIVE)" → "Teamwork"
                        label = col.replace("Competency Score:", "").strip()
                        label = label.replace("(STIVE)", "").strip()
                        competency_values[label] = float(row[col])

                labels = list(competency_values.keys())
                values = list(competency_values.values())

                if values:
                    avg_score = sum(values) / len(values)
                else:
                    avg_score = None

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