"""Genera EEFF sintéticos de 3 empresas ficticias para probar el modo financiero (cruce y eliminaciones)."""
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

OUT = os.path.join(os.path.dirname(__file__), '..', 'samples', 'financiero')
W, H = A4
fmt = lambda v: f'({abs(v):,})' if v < 0 else f'{v:,}'

def header(c, nombre, y, titulo=None, cierre='AL 31 DE JULIO DE 2026'):
    c.setFont('Helvetica-Bold', 11); c.drawCentredString(W / 2, y, nombre); y -= 14
    if titulo: c.drawCentredString(W / 2, y, titulo); y -= 14
    c.setFont('Helvetica', 9); c.drawCentredString(W / 2, y, cierre); y -= 12
    c.drawCentredString(W / 2, y, '(Expresado en Nuevos Soles)'); return y - 24

def row(c, y, label, nota=None, amount=None, total=None, bold=False):
    c.setFont('Helvetica-Bold' if bold else 'Helvetica', 9)
    c.drawString(50, y, label)
    if nota: c.drawString(330, y, nota)
    if amount is not None:
        c.drawRightString(440, y, fmt(amount))
        if total: c.drawRightString(510, y, f'{amount / total * 100:.2f}%')
    return y - 14

def empresa(nombre, activos, pasivos, patrimonio, notas, archivo):
    c = canvas.Canvas(os.path.join(OUT, archivo), pagesize=A4)
    y = header(c, nombre, H - 50, 'ESTADO DE SITUACION FINANCIERA')
    ta = sum(a for _, _, a in activos); tp = sum(a for _, _, a in pasivos); tpat = sum(a for _, _, a in patrimonio)
    y = row(c, y, 'ACTIVO CORRIENTE', bold=True)
    for l, nt, a in activos: y = row(c, y, l, nt, a, ta)
    y = row(c, y, 'TOTAL ACTIVO CORRIENTE', None, ta, ta, True); y = row(c, y, 'TOTAL ACTIVO', None, ta, ta, True); y -= 8
    y = row(c, y, 'PASIVO CORRIENTE', bold=True)
    for l, nt, a in pasivos: y = row(c, y, l, nt, a, ta)
    y = row(c, y, 'TOTAL PASIVO CORRIENTE', None, tp, ta, True); y = row(c, y, 'TOTAL PASIVO', None, tp, ta, True); y -= 8
    y = row(c, y, 'PATRIMONIO', bold=True)
    for l, nt, a in patrimonio: y = row(c, y, l, nt, a, ta)
    y = row(c, y, 'TOTAL PATRIMONIO', None, tpat, ta, True); y = row(c, y, 'TOTAL PASIVO Y PATRIMONIO', None, tp + tpat, ta, True)
    assert ta == tp + tpat, (nombre, ta, tp + tpat)
    c.showPage()
    y = header(c, nombre, H - 50)
    for titulo, detalle in notas:
        c.setFont('Helvetica-Bold', 9); c.drawString(50, y, titulo); y -= 14
        for l, a in detalle: c.setFont('Helvetica', 9); c.drawString(50, y, l); c.drawRightString(440, y, fmt(a)); y -= 14
        c.setFont('Helvetica-Bold', 9); c.drawString(50, y, 'SALDOS GENERALES:'); c.drawRightString(440, y, fmt(sum(a for _, a in detalle))); y -= 22
    c.showPage(); c.save()

os.makedirs(OUT, exist_ok=True)
empresa('ALFA PERU SAC',
    [('EFECTIVO Y EQUIVALENTE DE EFECTIVO', 'N.10', 50000), ('CUENTAS POR COBRAR COMERCIALES-TERCEROS', 'N.12', 30000), ('CUENTAS POR COBRAR RELACIONADAS', 'N.14', 100000)],
    [('CUENTAS POR PAGAR COM. - PROVEEDORES', 'N.42', 20000)], [('CAPITAL', 'N.50', 160000)],
    [('EFECTIVO Y EQUIVALENTE DE EFECTIVO', [('CAJA', 20000), ('BANCOS', 30000)]), ('CUENTAS POR COBRAR COMERCIALES', [('FACTURAS EMITIDAS', 30000)]),
     ('CUENTAS POR COBRAR RELACIONADAS', [('PRESTAMO A BETA SERVICIOS SAC', 100000)]), ('CUENTAS POR PAGAR COMERCIALES', [('FACTURAS POR PAGAR', 20000)]), ('CAPITAL SOCIAL', [('ACCIONES', 160000)])],
    'EEFF_ALFA_07-2026.pdf')
empresa('BETA SERVICIOS SAC',
    [('EFECTIVO Y EQUIVALENTE DE EFECTIVO', 'N.10', 220000)],
    [('CUENTAS POR PAGAR COM. - PROVEEDORES', 'N.42', 10000), ('CUENTAS POR PAGAR RELACIONADAS', 'N.46', 148000)], [('CAPITAL', 'N.50', 100000), ('RESULTADOS ACUMULADOS', 'N.59', -38000)],
    [('EFECTIVO Y EQUIVALENTE DE EFECTIVO', [('BANCOS', 220000)]), ('CUENTAS POR PAGAR COMERCIALES', [('FACTURAS POR PAGAR', 10000)]),
     ('CUENTAS POR PAGAR RELACIONADAS', [('PRESTAMO DE ALFA PERU SAC', 100000), ('PRESTAMO DE GAMMA SAC', 48000)]), ('CAPITAL SOCIAL', [('ACCIONES', 100000)]), ('RESULTADOS ACUMULADOS', [('PERDIDAS ACUMULADAS', -38000)])],
    'EEFF_BETA_07-2026.pdf')
empresa('GAMMA SAC',
    [('EFECTIVO Y EQUIVALENTE DE EFECTIVO', 'N.10', 30000), ('CUENTAS POR COBRAR RELACIONADAS', 'N.14', 50000)],
    [('CUENTAS POR PAGAR COM. - PROVEEDORES', 'N.42', 5000)], [('CAPITAL', 'N.50', 75000)],
    [('EFECTIVO Y EQUIVALENTE DE EFECTIVO', [('BANCOS', 30000)]), ('CUENTAS POR COBRAR RELACIONADAS', [('PRESTAMO A BETA SERVICIOS SAC', 50000)]),
     ('CUENTAS POR PAGAR COMERCIALES', [('FACTURAS POR PAGAR', 5000)]), ('CAPITAL SOCIAL', [('ACCIONES', 75000)])],
    'EEFF_GAMMA_07-2026.pdf')
print('ok')
