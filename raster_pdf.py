from pathlib import Path
import sys
import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont

pdf_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r"D:\project\innovation\proposal_render\proposal_hamrah_entekhab.pdf")
root = pdf_path.parent
for pattern in ("page-*.png", "contact-*.jpg"):
    for old in root.glob(pattern):
        old.unlink()
pdf = pdfium.PdfDocument(str(pdf_path))
pages = []
for i in range(len(pdf)):
    bitmap = pdf[i].render(scale=2.0)
    image = bitmap.to_pil().convert("RGB")
    out = root / f"page-{i+1}.png"
    image.save(out)
    pages.append(image)

font = ImageFont.truetype(r"C:\Windows\Fonts\tahoma.ttf", 24)
pad = 40
label_h = 46
for start in range(0, len(pages), 4):
    batch = pages[start:start+4]
    max_w = max(im.width for im in batch)
    max_h = max(im.height for im in batch)
    sheet = Image.new("RGB", (max_w*2 + pad*3, (max_h+label_h)*2 + pad*3), "#777777")
    draw = ImageDraw.Draw(sheet)
    for j, im in enumerate(batch):
        row, col = divmod(j, 2)
        x = pad + col*(max_w+pad)
        y = pad + row*(max_h+label_h+pad)
        draw.text((x, y), f"Page {start+j+1}", fill="white", font=font)
        sheet.paste(im, (x, y+label_h))
    sheet.save(root / f"contact-{start+1}-{start+len(batch)}.jpg", quality=90)

print(f"pages={len(pages)}")
