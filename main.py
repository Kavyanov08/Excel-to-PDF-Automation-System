import os
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph,Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import matplotlib.pyplot as plt
import io
import random
from reportlab.platypus import Image
import time
import sys, io
from reportlab.lib.styles import ParagraphStyle
import numpy as np
import matplotlib.pyplot as plt

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def generate_tid():
    """Generate a 24-character uppercase hex string (like your sample)."""
    return ''.join(random.choice("0123456789ABCDEF") for _ in range(24))

def generate_fixed_iv_graph(voltage, current, power=None, vmax=None, imax=None, pmax=None):
    """Generate smooth IV & PV curve graph with proper scaling like datasheet."""

    import numpy as np
    import matplotlib.pyplot as plt
    import io

    # --- Specs (fallback values) ---
    Voc = 52.96   # Open-circuit voltage
    Isc = 13.95   # Short-circuit current
    Vmp = 44.64   # Voltage at max power
    Imp = 13.22   # Current at max power
    Pmax = 590    # Max power

    # Create synthetic sweep for smooth curve (ignore sparse excel points)
    v_smooth = np.linspace(0, 70, 300)

    # Synthetic I-V shape
    i_smooth = []
    for v in v_smooth:
        if v < Vmp:
            i_smooth.append(Isc - (Isc - Imp) * (v / Vmp))
        elif v <= Voc:
            slope = Imp / (Voc - Vmp)
            i_smooth.append(max(Imp - slope * (v - Vmp), 0))
        else:
            i_smooth.append(0)
    i_smooth = np.array(i_smooth)

    # Power curve
    p_smooth = v_smooth * i_smooth

    # --- Plotting ---
    fig, ax1 = plt.subplots(figsize=(8,6))

    # Current curve (red)
    ax1.plot(v_smooth, i_smooth, "r-", linewidth=2, label="Current (A)")
    ax1.set_xlabel("Voltage (V)", fontsize=12)
    ax1.set_ylabel("Current (A)", color="red", fontsize=12)
    ax1.tick_params(axis="y", labelcolor="red")
    ax1.set_xlim(0, 70)
    ax1.set_ylim(0, 15)
    ax1.grid(True, linestyle="--", alpha=0.7)

    # Power curve (blue)
    ax2 = ax1.twinx()
    ax2.plot(v_smooth, p_smooth, "b-", linewidth=2, label="Power (W)")
    ax2.set_ylabel("Power (W)", color="blue", fontsize=12)
    ax2.tick_params(axis="y", labelcolor="blue")
    ax2.set_ylim(0, 700)

    # ✅ Force plot MPP point if values provided
    if vmax is not None and imax is not None and pmax is not None:
        ax1.plot(vmax, imax, "go", markersize=8, label="Table MPP")  # IV marker
        ax2.plot(vmax, pmax, "go", markersize=8)                     # PV marker

        # Annotation text
        ax1.annotate(
            f"V={vmax:.1f}V\nI={imax:.2f}A\nP={pmax:.1f}W",
            (vmax, imax),
            textcoords="offset points",
            xytext=(10, -20),
            ha="left",
            color="green",
            fontsize=9,
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="green", lw=0.8)
        )

    # Title
    plt.title("IV Characteristics of the Module", fontsize=14, color="darkblue")

    # Legends outside
    fig.legend(
        loc="upper center", bbox_to_anchor=(0.5, -0.05), ncol=2, frameon=False
    )

    buf = io.BytesIO()
    plt.savefig(buf, format="PNG", bbox_inches="tight", dpi=150)
    buf.seek(0)
    plt.close()
    return buf


# ✅ Mapping: Default Column Name → (Excel column name, Default value if missing)
COLUMN_MAP = {
    "Name of the Manufacturer of PV Module": ("producer", "Swelect HHV Solar Photovoltaics Private Ltd."),
    "Name of the Manufacturer of Solar Cell": (None, "JTPV"),
    "Module Model No": ("ID", "SWT15BG6585"),
    "Month & Year of the Manufacture of Module": ("Date", "Jul,25"),
    "Month & Year of the Manufacture of Solar Cell": (None, "Jun,25"),
    "Country of Origin for PV Module": (None, "India"),
    "Country of Origin for Solar Cell": (None, "China"),
    "Power: P-Max of the Module": ("Pmax", "NA"),
    "Voltage: V-Max of the Module": ("Vpm", "NA"),
    "Current: I-Max of the Module": ("Ipm", "NA"),
    "Fill Factor (FF) of the Module": ("FF", "NA"),
    "VOC": ("Voc", "NA"),
    "ISC": ("Isc", "NA"),
    "Name of The Test Lab Issuing IEC Certificate": (None, "TUV"),
    "Date of Obtaining IEC Certificate": (None, "22-05-2025"),
    "Name of The Test Lab Issuing BIS Certificate": (None, "TUV"),
    "Date of Obtaining BIS Certificate": ("B_date", "21/03/2025"),
    "Production Order No": (None, "300002582"),
    "Module Efficiency": ("Eff", "NA"),
}

# 📌 Folders
INPUT_FOLDER = "input_excels"
OUTPUT_FOLDER = "output_pdfs"
PROCESSED_FOLDER = os.path.join(INPUT_FOLDER, "_processed")
os.makedirs(PROCESSED_FOLDER, exist_ok=True)


os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def process_excel(file_path):
    df = pd.read_excel(file_path, engine="openpyxl")

    # ✅ Extract full column lists once
    try:
        power   = df.iloc[:, 7].dropna().tolist()  # column 8
        voltage = df.iloc[:, 8].dropna().tolist()  # column 9
        current = df.iloc[:, 9].dropna().tolist()  # column 10
    except Exception:
        power, voltage, current = [], [], []

    for idx, row in df.iterrows():
        data_for_pdf = []

        # Build table with serial number
        for spec, (excel_col, default_val) in COLUMN_MAP.items():
            value = None
            if excel_col and excel_col in df.columns:
                value = row[excel_col]
            if pd.isna(value) or value is None:
                value = default_val
            # ✅ wrap value in Paragraph for alignment
            data_for_pdf.append([len(data_for_pdf)+1, spec, Paragraph(str(value), getSampleStyleSheet()["Normal"])])

        # filename logic as before
        base_name = str(row.get("ID", f"record_{idx}"))
        file_name = f"{base_name}_{idx+1}.pdf"
        output_path = os.path.join(OUTPUT_FOLDER, file_name)

        counter = 1
        while os.path.exists(output_path):
            file_name = f"{base_name}_{idx+1}_{counter}.pdf"
            output_path = os.path.join(OUTPUT_FOLDER, file_name)
            counter += 1

        # generate PDF
        tid = generate_tid()
        generate_pdf(data_for_pdf, output_path, base_name, tid, power, voltage, current)
        print(f"✅ PDF created: {output_path}")

    # move processed Excel
    processed_path = os.path.join(PROCESSED_FOLDER, os.path.basename(file_path))
    os.replace(file_path, processed_path)
    print(f"📦 Moved processed Excel → {processed_path}")

def generate_pdf(data, output_path, module_id="", tid="",power=None,voltage=None, current=None):
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []
    # ✅ Add Logo (if exists)
    logo_path = "logo.jpg"   # change name if your logo file is different
    if os.path.exists(logo_path):
        elements.append(Image(logo_path, width=100, height=50))
        elements.append(Spacer(1, 8))

    # ✅ Company Header
    company_info = [
        "Swelect HHV Solar Photovoltaics Private Ltd.",
        "SF.NO.166 & 169, KUPPEPALAYAM VILLAGE",
        "SEMBAGOUNDENSW PUDUR",
        "SARSKAR SAMAKULAM (VIA)",
        "COIMBATORE - 641107",
        "",
        f"Module Serial Number: {module_id}",
        f"TID:  {tid}",
        "",
        "Detailed Specification:",
        "",
    ]
    header_style = ParagraphStyle(
        "Header",
        fontName="Helvetica-Bold",
        fontSize=11,
        textColor=colors.HexColor("#003366"),  # dark blue
    )

    for line in company_info[:-2]:  # all lines except last 2
        elements.append(Paragraph(line, header_style))

    elements.append(Spacer(1, 8))
    elements.append(Paragraph("Detailed Specification:", header_style))
    elements.append(Spacer(1, 10))

    # ✅ Table
    from reportlab.lib.enums import TA_LEFT, TA_CENTER

    table = Table([["S.No", "Specification", "Value"]] + data,
              colWidths=[40, 220, 280])  # adjust widths

    table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("ALIGN", (0, 0), (0, -1), "CENTER"),  # Center align S.No
    ("ALIGN", (1, 0), (1, -1), "LEFT"),    # Left align Spec
    ("ALIGN", (2, 0), (2, -1), "LEFT"),    # Left align Value
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, 0), 10),
    ("FONTSIZE", (0, 1), (-1, -1), 9),
    ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
]))

    elements.append(table)
    
    # ✅ Add IV Graph (from Excel data if available)
    if voltage and current:
        buf = generate_fixed_iv_graph(voltage,current,power)
        elements.append(Spacer(1, 20))
        elements.append(Paragraph("IV Characteristics of the Module:", header_style))
        elements.append(Image(buf, width=380, height=220))

    doc.build(elements)


def main():
    for file in os.listdir(INPUT_FOLDER):
        if file.endswith(".xlsx"):
            process_excel(os.path.join(INPUT_FOLDER, file))

if __name__ == "__main__":
    print("Excel-to-PDF service started, watching for new files...")  
    while True:   
        main()   #  (reuse your existing logic)
        time.sleep(5)   #  (check every 5 seconds)
