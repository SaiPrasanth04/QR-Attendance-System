from fpdf import FPDF
def pdf_qr_for_batch(excel_path: str, batch: str, base_qr_dir: str, output_dir: str) -> str:
    """
    Generate a PDF (A4) with 12 QR codes per page, each with the student's name below.
    Returns: path to the PDF file.
    """
    import pandas as pd
    df = pd.read_excel(excel_path, sheet_name=batch, engine="openpyxl")
    if "ID" not in df.columns or "Name" not in df.columns:
        raise ValueError(f"Sheet '{batch}' must have 'ID' and 'Name' columns.")
    folder = os.path.join(base_qr_dir, batch)
    pdf_path = os.path.join(output_dir, f"{batch}_qrs.pdf")
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    qr_size = 40  # mm
    margin_x = 15
    margin_y = 15
    space_x = 10
    space_y = 10
    per_row = 3
    per_col = 4
    per_page = per_row * per_col
    x0 = margin_x
    y0 = margin_y
    import glob
    qrs = []
    for _, row in df.iterrows():
        sid = str(row["ID"]).strip()
        sname = str(row["Name"]).strip()
        # Match QR image ignoring case and spaces
        pattern = os.path.join(folder, f"{sid}_*.png")
        matches = [f for f in glob.glob(pattern) if os.path.splitext(os.path.basename(f))[0].lower().replace(' ', '') == f"{sid}_{sname}".lower().replace(' ', '')]
        if matches:
            qrs.append((matches[0], sname))
    if not qrs:
        pdf.add_page()
        pdf.set_font("Arial", size=16)
        if len(df) == 0:
            pdf.cell(0, 20, f"No students found for batch {batch}.", ln=1, align="C")
        else:
            pdf.cell(0, 20, f"No QR images found for students in batch {batch}.\nPlease generate all QR codes first.", ln=1, align="C")
    else:
        for i, (img_path, sname) in enumerate(qrs):
            if i % per_page == 0:
                pdf.add_page()
            pos = i % per_page
            row = pos // per_row
            col = pos % per_row
            x = x0 + col * (qr_size + space_x)
            y = y0 + row * (qr_size + space_y + 8)
            pdf.image(img_path, x=x, y=y, w=qr_size, h=qr_size)
            pdf.set_xy(x, y + qr_size)
            pdf.set_font("Arial", size=10)
            pdf.cell(qr_size, 8, sname, align="C")
    pdf.output(pdf_path)
    return pdf_path
import os
import json
import qrcode
from zipfile import ZipFile

def _ensure_batch_dir(base_qr_dir: str, batch: str):
    """Ensure the folder for a batch exists and return its path."""
    out_dir = os.path.join(base_qr_dir, batch)
    os.makedirs(out_dir, exist_ok=True)
    return out_dir

def _qr_payload(student_id: str, student_name: str, batch: str) -> str:
    """
    JSON payload for QR code.
    Example: {"id":"A001","name":"John Smith","batch":"BatchA"}
    """
    return json.dumps({"id": student_id, "name": student_name, "batch": batch})

def generate_qr_for_student(student_id: str, student_name: str, batch: str, base_qr_dir: str) -> str:
    """Generate a QR code for a single student and save as PNG."""
    out_dir = _ensure_batch_dir(base_qr_dir, batch)
    payload = _qr_payload(student_id, student_name, batch)
    img = qrcode.make(payload)
    safe_name = f"{student_id}_{student_name}".replace("/", "_").replace("\\", "_")
    out_path = os.path.join(out_dir, f"{safe_name}.png")
    img.save(out_path)
    return out_path

def generate_qr_for_batch(excel_path: str, batch: str, base_qr_dir: str):
    """
    Generate QR codes for all students in a batch (read from Excel).
    Removes old QR codes for students not in Excel.
    Returns: output directory and count of QR codes generated.
    """
    import pandas as pd
    df = pd.read_excel(excel_path, sheet_name=batch, engine="openpyxl")
    # Require ID & Name columns
    if "ID" not in df.columns or "Name" not in df.columns:
        raise ValueError(f"Sheet '{batch}' must have 'ID' and 'Name' columns.")
    out_dir = _ensure_batch_dir(base_qr_dir, batch)
    # Remove all old QR codes for this batch
    for f in os.listdir(out_dir):
        if f.lower().endswith('.png'):
            os.remove(os.path.join(out_dir, f))
    count = 0
    for _, row in df.iterrows():
        sid = str(row["ID"])
        sname = str(row["Name"])
        generate_qr_for_student(sid, sname, batch, base_qr_dir)
        count += 1
    return out_dir, count

def zip_qr_for_batch(batch: str, base_qr_dir: str, output_dir: str) -> str:
    """
    Create a ZIP file of all QR PNGs for a batch.
    Returns: path to the ZIP file.
    """
    folder = os.path.join(base_qr_dir, batch)
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"No QR folder for batch {batch}")
    
    zip_path = os.path.join(output_dir, f"{batch}_qrs.zip")
    with ZipFile(zip_path, "w") as zf:
        for fname in os.listdir(folder):
            if fname.lower().endswith(".png"):
                zf.write(os.path.join(folder, fname), arcname=os.path.join(batch, fname))
    return zip_path
