import os
import math
from PIL import Image, ImageOps
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import ezdxf
from pypdf import PdfReader, PdfWriter, Transformation

# ==========================
# CONFIG
# ==========================
INPUT_FOLDER = "Frentes"
VERSOS_FOLDER = "Versos"
OUTPUT_PDF = "output.pdf"
MARKS_PDF = "marks-landscape.pdf"
FINAL_PDF = "final.pdf"
OUTPUT_DXF = "cut_lines.dxf"

CARD_WIDTH_MM = 45
CARD_HEIGHT_MM = 68
BLEED_MM = 3
CORNER_RADIUS_MM = 2

BLEED_MODE = "mirror"  # "mirror" ou "color"
BLEED_COLOR = (255, 0, 0)

GRID_COLS = 5
GRID_ROWS = 2

ORIENTATION = "landscape"  # "portrait" ou "landscape"

DPI = 300

# ==========================
# UTIL
# ==========================
def mm_to_px(mm_value):
    return int((mm_value / 25.4) * DPI)

# ==========================
# BLEED
# ==========================
def apply_bleed(image, bleed_px, mode="mirror", color=(255, 0, 0)):
    w, h = image.size
    new_img = Image.new("RGB", (w + 2 * bleed_px, h + 2 * bleed_px), color)
    new_img.paste(image, (bleed_px, bleed_px))

    if mode == "color":
        return new_img

    # bordas
    top = ImageOps.flip(image.crop((0, 0, w, bleed_px)))
    bottom = ImageOps.flip(image.crop((0, h - bleed_px, w, h)))
    left = ImageOps.mirror(image.crop((0, 0, bleed_px, h)))
    right = ImageOps.mirror(image.crop((w - bleed_px, 0, w, h)))

    new_img.paste(top, (bleed_px, 0))
    new_img.paste(bottom, (bleed_px, h + bleed_px))
    new_img.paste(left, (0, bleed_px))
    new_img.paste(right, (w + bleed_px, bleed_px))

    # cantos
    def corner(x1, y1, x2, y2):
        return ImageOps.mirror(ImageOps.flip(image.crop((x1, y1, x2, y2))))

    new_img.paste(corner(0, 0, bleed_px, bleed_px), (0, 0))
    new_img.paste(corner(w - bleed_px, 0, w, bleed_px), (w + bleed_px, 0))
    new_img.paste(corner(0, h - bleed_px, bleed_px, h), (0, h + bleed_px))
    new_img.paste(corner(w - bleed_px, h - bleed_px, w, h), (w + bleed_px, h + bleed_px))

    return new_img

# ==========================
# HELPERS
# ==========================
def process_card_image(path, card_w_px, card_h_px, bleed_px):
    img = Image.open(path).convert("RGB")
    filename = os.path.basename(path)
    
    if filename.startswith('N'):
        # Desabilita bleed mirror e expande a carta para o tamanho total (card + bleeds)
        return img.resize((card_w_px + 2 * bleed_px, card_h_px + 2 * bleed_px), Image.LANCZOS)
    else:
        img = img.resize((card_w_px, card_h_px), Image.LANCZOS)
        return apply_bleed(img, bleed_px, BLEED_MODE, BLEED_COLOR)

# ==========================
# PDF
# ==========================
def generate_pdf(front_images, back_images):
    if ORIENTATION == "portrait":
        a4_w, a4_h = 210 * mm, 297 * mm
    else:
        a4_w, a4_h = 297 * mm, 210 * mm

    c = canvas.Canvas(OUTPUT_PDF, pagesize=(a4_w, a4_h))

    card_w_px = mm_to_px(CARD_WIDTH_MM)
    card_h_px = mm_to_px(CARD_HEIGHT_MM)
    bleed_px = mm_to_px(BLEED_MM)

    card_w_pt = (CARD_WIDTH_MM + 2 * BLEED_MM) * mm
    card_h_pt = (CARD_HEIGHT_MM + 2 * BLEED_MM) * mm

    grid_w = GRID_COLS * card_w_pt
    grid_h = GRID_ROWS * card_h_pt

    offset_x = (a4_w - grid_w) / 2
    offset_y = (a4_h - grid_h) / 2

    cards_per_page = GRID_COLS * GRID_ROWS
    num_pages = math.ceil(len(front_images) / cards_per_page)

    for p in range(num_pages):
        # --- PÁGINA DE FRENTES ---
        idx_start = p * cards_per_page
        for row in range(GRID_ROWS):
            for col in range(GRID_COLS):
                idx = idx_start + row * GRID_COLS + col
                if idx < len(front_images):
                    img = process_card_image(front_images[idx], card_w_px, card_h_px, bleed_px)
                    
                    temp = f"_tmp_f_{idx}.jpg"
                    img.save(temp, quality=95)
                    x = offset_x + col * card_w_pt
                    y = offset_y + (GRID_ROWS - 1 - row) * card_h_pt
                    c.drawImage(temp, x, y, width=card_w_pt, height=card_h_pt)
                    os.remove(temp)
        c.showPage()

        # --- PÁGINA DE VERSOS ---
        if back_images:
            if ORIENTATION == "landscape":
                c.saveState()
                # Move a origem para o canto oposto e rotaciona 180 graus
                c.translate(a4_w, a4_h)
                c.rotate(180)

            for row in range(GRID_ROWS):
                for col in range(GRID_COLS):
                    if ORIENTATION == "landscape":
                        # A rotação de 180 graus do canvas já inverte a ordem de colunas e linhas
                        # O verso da carta (row, col) da frente estará na mesma (row, col) do canvas rotacionado
                        draw_col = (GRID_COLS - 1 - col)
                    else:
                        # No modo portrait (duplex em eixo longo), invertemos apenas a ordem das colunas
                        draw_col = (GRID_COLS - 1 - col)

                    idx = idx_start + row * GRID_COLS + col

                    if idx < len(back_images):
                        img = process_card_image(back_images[idx], card_w_px, card_h_px, bleed_px)

                        temp = f"_tmp_b_{idx}.jpg"
                        img.save(temp, quality=95)

                        x = offset_x + draw_col * card_w_pt
                        y = offset_y + (GRID_ROWS - 1 - row) * card_h_pt

                        c.drawImage(temp, x, y, width=card_w_pt, height=card_h_pt)
                        os.remove(temp)

            if ORIENTATION == "landscape":
                c.restoreState()
            c.showPage()

    c.save()

# ==========================
# ROUNDED RECT (DXF)
# ==========================
def add_rounded_rect(msp, x, y, w, h, r, segments=20):
    r = min(r, w/2, h/2)

    def arc(cx, cy, start, end):
        return [
            (
                cx + r * math.cos(math.radians(a)),
                cy + r * math.sin(math.radians(a))
            )
            for a in [start + (end - start) * i / segments for i in range(segments + 1)]
        ]

    pts = []
    pts.append((x + r, y))
    pts.append((x + w - r, y))
    pts.extend(arc(x + w - r, y + r, 270, 360)[1:])
    pts.append((x + w, y + h - r))
    pts.extend(arc(x + w - r, y + h - r, 0, 90)[1:])
    pts.append((x + r, y + h))
    pts.extend(arc(x + r, y + h - r, 90, 180)[1:])
    pts.append((x, y + r))
    pts.extend(arc(x + r, y + r, 180, 270)[1:])

    for i in range(len(pts)):
        msp.add_line(pts[i], pts[(i + 1) % len(pts)])

# ==========================
# DXF
# ==========================
def generate_dxf(images):
    doc = ezdxf.new("R12")
    msp = doc.modelspace()

    if ORIENTATION == "portrait":
        a4_w, a4_h = 210, 297
    else:
        a4_w, a4_h = 297, 210

    full_w = CARD_WIDTH_MM + 2 * BLEED_MM
    full_h = CARD_HEIGHT_MM + 2 * BLEED_MM

    grid_w = GRID_COLS * full_w
    grid_h = GRID_ROWS * full_h

    offset_x = (a4_w - grid_w) / 2
    offset_y = (a4_h - grid_h) / 2

    idx = 0
    total = len(images)

    while idx < total:
        for row in range(GRID_ROWS):
            for col in range(GRID_COLS):
                if idx >= total:
                    break

                x_full = offset_x + col * full_w
                y_full = offset_y + (GRID_ROWS - 1 - row) * full_h

                x_card = x_full + BLEED_MM
                y_card = y_full + BLEED_MM

                add_rounded_rect(
                    msp,
                    x_card,
                    y_card,
                    CARD_WIDTH_MM,
                    CARD_HEIGHT_MM,
                    CORNER_RADIUS_MM
                )

                idx += 1

    # força visibilidade no Silhouette
    msp.add_point((0, 0))

    doc.saveas(OUTPUT_DXF)

import copy

def merge_marks_corners(base_pdf, marks_pdf, output_pdf):
    base_reader = PdfReader(base_pdf)
    marks_reader = PdfReader(marks_pdf)
    writer = PdfWriter()

    for base_page in base_reader.pages:
        width = float(base_page.mediabox.width)
        height = float(base_page.mediabox.height)

        margin = 64  # ajuste fino aqui

        def merge_crop(box):
            mark = copy.copy(marks_reader.pages[0])

            # define cropbox corretamente
            mark.cropbox.lower_left = (box[0], box[1])
            mark.cropbox.upper_right = (box[2], box[3])

            # merge respeitando transformação
            base_page.merge_transformed_page(
                mark,
                Transformation()  # identidade (sem mover)
            )

        # 4 cantos
        merge_crop((0, height - margin, margin, height))                # top-left
        merge_crop((width - margin, height - margin, width, height))    # top-right
        merge_crop((0, 0, margin, margin))                              # bottom-left
        merge_crop((width - margin, 0, width, margin))                  # bottom-right

        writer.add_page(base_page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

# ==========================
# RUN
# ==========================
if __name__ == "__main__":
    front_images = [
        os.path.join(INPUT_FOLDER, f)
        for f in sorted(os.listdir(INPUT_FOLDER))
        if f.lower().endswith((".png", ".jpg", ".jpeg"))
    ]
    
    back_images = []
    if os.path.exists(VERSOS_FOLDER):
        back_images = [
            os.path.join(VERSOS_FOLDER, f)
            for f in sorted(os.listdir(VERSOS_FOLDER))
            if f.lower().endswith((".png", ".jpg", ".jpeg"))
        ]

    generate_pdf(front_images, back_images)
    merge_marks_corners(OUTPUT_PDF, MARKS_PDF, FINAL_PDF)
    generate_dxf(front_images)

    print("✅ PDF e DXF gerados com sucesso!")