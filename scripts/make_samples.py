"""Genera PDFs de ejemplo para probar (requiere reportlab y pdftoppm)."""
import subprocess, os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

out = os.path.join(os.path.dirname(__file__), '..', 'samples')
header = ['Codigo', 'Descripcion', 'Cantidad', 'Precio', 'Total']
rows = [[f'P{i:03d}', d, str(q), f'{p:,.2f}', f'{q*p:,.2f}'] for i, (d, q, p) in enumerate([
    ('Cuaderno cuadriculado A4', 10, 12.5), ('Lapicero azul', 200, 1.2), ('Regla de metal 30 cm', 15, 8),
    ('Archivador de palanca', 8, 22.9), ('Folder manila', 100, 0.8), ('Resaltador amarillo', 40, 2.5),
    ('Tijera escolar', 12, 6.75), ('Goma en barra', 60, 3.1), ('Cinta adhesiva', 25, 4.2),
    ('Papel bond A4 (millar)', 5, 35), ('Corrector liquido', 30, 3.9), ('Grapadora', 6, 18.5),
], 1)]

def draw(path, with_title=True):
    c = canvas.Canvas(path, pagesize=A4)
    W, H = A4
    xs = [50, 120, 330, 400, 480]
    def page(chunk, first):
        y = H - 60
        if first:
            c.setFont('Helvetica-Bold', 14); c.drawString(50, y, 'Orden de compra N° 0045 - Útiles de oficina'); y -= 40
        c.setFont('Helvetica-Bold', 10)
        for x, t in zip(xs, header): c.drawString(x, y, t)
        y -= 18; c.setFont('Helvetica', 10)
        for r in chunk:
            for k, (x, t) in enumerate(zip(xs, r)):
                (c.drawRightString if k >= 3 else c.drawString)(x + (50 if k >= 3 else 0), y, t)
            y -= 18
        c.showPage()
    page(rows[:7], True)
    page(rows[7:], False)
    c.save()

digital = os.path.join(out, 'orden-compra-digital.pdf')
draw(digital)
# Versión "escaneada": se rasteriza y se vuelve a empaquetar como imágenes sin texto.
subprocess.run(['pdftoppm', '-r', '150', '-png', digital, os.path.join(out, 'escaneo')], check=True)
print('ok')
