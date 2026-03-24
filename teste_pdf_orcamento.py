# teste_pdf_orcamento.py — gera PDF idêntico ao GestãoClick com header Locvix
import requests, time
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph,
                                 Spacer, Image)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from io import BytesIO
import os

ACCESS  = "a5054bee0e7905eb8488d4e8e8e671866624a525"
SECRET  = "63e08b56cfde7fdac3afe4c52ef455c8c480b621"
HEADERS = {"access-token": ACCESS, "secret-access-token": SECRET}
API     = "https://api.gestaoclick.com"

# ── 1. Busca o orçamento mais recente ──────────────────────────────────────
r = requests.get(f"{API}/orcamentos?limite=1", headers=HEADERS, timeout=15)
orc = r.json()["data"][0]
orc_id = orc["id"]

# Detalhes do orçamento
r2 = requests.get(f"{API}/orcamentos/{orc_id}", headers=HEADERS, timeout=15)
d = r2.json()["data"]

# Dados do cliente
cli_id = d.get("cliente_id", "")
cli_data = {}
if cli_id:
    time.sleep(0.4)
    r3 = requests.get(f"{API}/clientes/{cli_id}", headers=HEADERS, timeout=15)
    cli_data = r3.json().get("data", {})

# Busca código do catálogo de cada serviço (campo "codigo" do /servicos/{id})
_cache_serv = {}
def get_servico_codigo(servico_id):
    if not servico_id:
        return ""
    if servico_id in _cache_serv:
        return _cache_serv[servico_id]
    time.sleep(0.4)
    try:
        r = requests.get(f"{API}/servicos/{servico_id}", headers=HEADERS, timeout=10)
        cod = r.json().get("data", {}).get("codigo", "")
    except:
        cod = str(servico_id)
    _cache_serv[servico_id] = cod
    return cod

# Busca código do catálogo de cada produto (campo "codigo" do /produtos/{id})
_cache_prod = {}
def get_produto_codigo(produto_id):
    if not produto_id:
        return ""
    if produto_id in _cache_prod:
        return _cache_prod[produto_id]
    time.sleep(0.4)
    try:
        r = requests.get(f"{API}/produtos/{produto_id}", headers=HEADERS, timeout=10)
        cod = r.json().get("data", {}).get("codigo", "")
    except:
        cod = str(produto_id)
    _cache_prod[produto_id] = cod
    return cod

# ── 2. Helpers ─────────────────────────────────────────────────────────────
def brl(v):
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "0,00"

def fmt_date(dt):
    if not dt:
        return ""
    parts = dt.split("-")
    if len(parts) == 3:
        return f"{parts[2]}/{parts[1]}/{parts[0]}"
    return dt

def cli_endereco():
    enderecos = cli_data.get("enderecos", [])
    if not enderecos:
        return {"logradouro":"","numero":"","bairro":"","cep":"","nome_cidade":"","estado":""}
    e = enderecos[0].get("endereco", enderecos[0])
    return e

# ── 3. Estilos (preto/branco — idêntico GestãoClick) ──────────────────────
w, h = A4
PRETO    = colors.black
CINZA_BG = colors.HexColor("#d9d9d9")
BRANCO   = colors.white

st_secao   = ParagraphStyle("sec", fontSize=9, fontName="Helvetica-Bold", textColor=PRETO)
st_label   = ParagraphStyle("lbl", fontSize=8, fontName="Helvetica-Bold", textColor=PRETO)
st_valor   = ParagraphStyle("val", fontSize=8, fontName="Helvetica", textColor=PRETO)
st_th      = ParagraphStyle("th",  fontSize=8, fontName="Helvetica-Bold", textColor=PRETO)
st_th_c    = ParagraphStyle("thc", fontSize=8, fontName="Helvetica-Bold", textColor=PRETO, alignment=TA_CENTER)
st_th_r    = ParagraphStyle("thr", fontSize=8, fontName="Helvetica-Bold", textColor=PRETO, alignment=TA_RIGHT)
st_td      = ParagraphStyle("td",  fontSize=8, fontName="Helvetica", textColor=PRETO)
st_td_c    = ParagraphStyle("tdc", fontSize=8, fontName="Helvetica", textColor=PRETO, alignment=TA_CENTER)
st_td_r    = ParagraphStyle("tdr", fontSize=8, fontName="Helvetica", textColor=PRETO, alignment=TA_RIGHT)
st_tot_lbl = ParagraphStyle("tl",  fontSize=8, fontName="Helvetica-Bold", textColor=PRETO, alignment=TA_RIGHT)
st_tot_val = ParagraphStyle("tv",  fontSize=8, fontName="Helvetica", textColor=PRETO, alignment=TA_RIGHT)
st_tot_fim = ParagraphStyle("tf",  fontSize=9, fontName="Helvetica-Bold", textColor=PRETO, alignment=TA_RIGHT)
st_hdr_emp = ParagraphStyle("he",  fontSize=14, fontName="Helvetica-Bold", alignment=TA_CENTER, textColor=PRETO)
st_hdr_inf = ParagraphStyle("hi",  fontSize=8,  fontName="Helvetica", alignment=TA_CENTER, textColor=PRETO, leading=11)

CW = w - 30*mm

buf = BytesIO()
doc = SimpleDocTemplate(buf, pagesize=A4,
                        leftMargin=15*mm, rightMargin=15*mm,
                        topMargin=12*mm, bottomMargin=12*mm)
elems = []

# ── Helpers de tabela ──────────────────────────────────────────────────────
def borda():
    return [
        ("BOX",       (0,0), (-1,-1), 0.5, PRETO),
        ("INNERGRID", (0,0), (-1,-1), 0.5, PRETO),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 4),
        ("RIGHTPADDING",  (0,0), (-1,-1), 4),
    ]

def secao(texto):
    t = Table([[Paragraph(texto, st_secao)]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), CINZA_BG),
        ("BOX", (0,0), (-1,-1), 0.5, PRETO),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
    ]))
    return t

# ══════════════════════════════════════════════════════════════════════════
# CABEÇALHO LOCVIX (parte que NÃO existe no GestãoClick)
# ══════════════════════════════════════════════════════════════════════════
logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
cab_rows = []
if os.path.exists(logo_path):
    cab_rows.append([Image(logo_path, width=50*mm, height=15*mm)])
cab_rows.append([Paragraph("LOCVIX LOCAÇÕES LTDA", st_hdr_emp)])
cab_rows.append([Paragraph(
    "CNPJ: 29.007.819/0001-96  •  Serra/ES  •  (27) 3065-2627  •  contato@locvix.com.br",
    st_hdr_inf)])
cab_rows.append([Paragraph(
    f"PROPOSTA COMERCIAL Nº {d.get('codigo','')}", st_hdr_emp)])
cab = Table(cab_rows, colWidths=[CW])
cab.setStyle(TableStyle([
    ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ("TOPPADDING", (0,0), (-1,-1), 2),
    ("BOTTOMPADDING", (0,0), (-1,-1), 2),
]))
elems.append(cab)
elems.append(Spacer(1, 4*mm))

# ══════════════════════════════════════════════════════════════════════════
# DAQUI PRA BAIXO — LAYOUT IDÊNTICO AO GESTÃOCLICK
# ══════════════════════════════════════════════════════════════════════════

# ── VALIDADE / PREVISÃO ────────────────────────────────────────────────────
validade = d.get("validade") or "10 DIAS"
previsao = fmt_date(d.get("previsao_entrega", ""))
tbl_val = Table([
    [Paragraph(f"VALIDADE DA PROPOSTA: {validade}", st_label),
     Paragraph(f"PREVISÃO DE ENTREGA: {previsao}", st_label)],
], colWidths=[CW/2, CW/2])
tbl_val.setStyle(TableStyle(borda() + [
    ("BACKGROUND", (0,0), (-1,-1), CINZA_BG),
]))
elems.append(tbl_val)
elems.append(Spacer(1, 2*mm))

# ── DADOS DO CLIENTE ──────────────────────────────────────────────────────
elems.append(secao("DADOS DO CLIENTE"))
end = cli_endereco()
logradouro = end.get("logradouro","")
numero     = end.get("numero","")
bairro     = end.get("bairro","")
endereco_completo = f"{logradouro}, {numero}" + (f" - {bairro}" if bairro else "")
cidade_uf  = f"{end.get('nome_cidade','')}/{end.get('estado','')}"
cep        = end.get("cep","")
telefone   = cli_data.get("telefone","")
email      = cli_data.get("email","")
cnpj_cpf   = cli_data.get("cnpj","") or cli_data.get("cpf","")
nome_fant  = cli_data.get("nome","")
razao      = cli_data.get("razao_social","") or d.get("nome_cliente","")

col_lbl = 24*mm
col_val = CW/2 - col_lbl
rows_cli = [
    [Paragraph("Razão social:", st_label), Paragraph(razao, st_valor),
     Paragraph("Nome fantasia:", st_label), Paragraph(nome_fant, st_valor)],
    [Paragraph("CNPJ/CPF:", st_label), Paragraph(cnpj_cpf, st_valor),
     Paragraph("Endereço:", st_label), Paragraph(endereco_completo, st_valor)],
    [Paragraph("CEP:", st_label), Paragraph(cep, st_valor),
     Paragraph("Cidade/UF:", st_label), Paragraph(cidade_uf, st_valor)],
    [Paragraph("Telefone:", st_label), Paragraph(telefone, st_valor),
     Paragraph("E-mail:", st_label), Paragraph(email, st_valor)],
]
tbl_cli = Table(rows_cli, colWidths=[col_lbl, col_val, col_lbl, col_val])
tbl_cli.setStyle(TableStyle(borda()))
elems.append(tbl_cli)
elems.append(Spacer(1, 3*mm))

# ── SERVIÇOS ──────────────────────────────────────────────────────────────
servicos = d.get("servicos", [])
if servicos:
    elems.append(secao("SERVIÇOS"))
    c_item = 12*mm; c_cod = 32*mm; c_qtd = 18*mm; c_vr = 22*mm; c_sub = 24*mm
    c_nome = CW - c_item - c_cod - c_qtd - c_vr - c_sub
    cws = [c_item, c_cod, c_nome, c_qtd, c_vr, c_sub]
    th = [Paragraph("ITEM", st_th_c), Paragraph("CÓDIGO", st_th_c),
          Paragraph("NOME", st_th), Paragraph("QTD.", st_th_c),
          Paragraph("VR.<br/>UNIT.", st_th_c), Paragraph("SUBTOTAL", st_th_c)]
    rows_s = [th]
    total_qtd_s = 0
    total_val_s = 0
    for i, sv in enumerate(servicos, 1):
        s = sv.get("servico", sv)
        sid = str(s.get("servico_id",""))
        codigo = get_servico_codigo(sid)
        q = float(s.get("quantidade",0) or 0)
        vt = float(s.get("valor_total",0) or 0)
        total_qtd_s += q
        total_val_s += vt
        rows_s.append([
            Paragraph(str(i), st_td_c),
            Paragraph(codigo, st_td_c),
            Paragraph(s.get("nome_servico",""), st_td),
            Paragraph(brl(q), st_td_r),
            Paragraph(brl(s.get("valor_venda",0)), st_td_r),
            Paragraph(brl(vt), st_td_r),
        ])
    rows_s.append([
        Paragraph("TOTAL", st_th), "", "", "",
        Paragraph(brl(total_qtd_s), st_th_r),
        Paragraph(brl(total_val_s), st_th_r),
    ])
    tbl_s = Table(rows_s, colWidths=cws)
    tbl_s.setStyle(TableStyle(borda() + [
        ("BACKGROUND", (0,0), (-1,0), CINZA_BG),
    ]))
    elems.append(tbl_s)
    elems.append(Spacer(1, 3*mm))

# ── PRODUTOS ──────────────────────────────────────────────────────────────
produtos = d.get("produtos", [])
if produtos:
    elems.append(secao("PRODUTOS"))
    c_item = 12*mm; c_cod = 20*mm; c_und = 14*mm; c_qtd = 18*mm; c_vr = 22*mm; c_sub = 24*mm
    c_nome_p = CW - c_item - c_cod - c_und - c_qtd - c_vr - c_sub
    cwp = [c_item, c_cod, c_nome_p, c_und, c_qtd, c_vr, c_sub]
    th = [Paragraph("ITEM", st_th_c), Paragraph("CÓDIGO", st_th_c),
          Paragraph("NOME", st_th), Paragraph("UND.", st_th_c),
          Paragraph("QTD.", st_th_c), Paragraph("VR.<br/>UNIT.", st_th_c),
          Paragraph("SUBTOTAL", st_th_c)]
    rows_p = [th]
    total_qtd_p = 0
    total_val_p = 0
    for i, pv in enumerate(produtos, 1):
        p = pv.get("produto", pv)
        pid = str(p.get("produto_id",""))
        codigo = get_produto_codigo(pid)
        q = float(p.get("quantidade",0) or 0)
        vt = float(p.get("valor_total",0) or 0)
        total_qtd_p += q
        total_val_p += vt
        rows_p.append([
            Paragraph(str(i), st_td_c),
            Paragraph(codigo, st_td_c),
            Paragraph(p.get("nome_produto",""), st_td),
            Paragraph(str(p.get("sigla_unidade","") or ""), st_td_c),
            Paragraph(brl(q), st_td_r),
            Paragraph(brl(p.get("valor_venda",0)), st_td_r),
            Paragraph(brl(vt), st_td_r),
        ])
    rows_p.append([
        Paragraph("TOTAL", st_th), "", "", "", "",
        Paragraph(brl(total_qtd_p), st_th_r),
        Paragraph(brl(total_val_p), st_th_r),
    ])
    tbl_p = Table(rows_p, colWidths=cwp)
    tbl_p.setStyle(TableStyle(borda() + [
        ("BACKGROUND", (0,0), (-1,0), CINZA_BG),
    ]))
    elems.append(tbl_p)
    elems.append(Spacer(1, 3*mm))

# ── TOTAIS (alinhados à direita, como GestãoClick) ────────────────────────
total_geral = float(d.get("valor_total", 0) or 0)
val_serv    = float(d.get("valor_servicos", 0) or 0)
val_prod    = float(d.get("valor_produtos", 0) or 0)
val_frete   = float(d.get("valor_frete", 0) or 0)

totais_rows = []
if val_prod  > 0: totais_rows.append([Paragraph("PRODUTOS:", st_tot_lbl), Paragraph(brl(val_prod), st_tot_val)])
if val_serv  > 0: totais_rows.append([Paragraph("SERVIÇOS:", st_tot_lbl), Paragraph(brl(val_serv), st_tot_val)])
if val_frete > 0: totais_rows.append([Paragraph("FRETE:", st_tot_lbl), Paragraph(brl(val_frete), st_tot_val)])
totais_rows.append([Paragraph("TOTAL:", st_tot_lbl), Paragraph(f"R$ {brl(total_geral)}", st_tot_fim)])
tbl_tot = Table(totais_rows, colWidths=[35*mm, 30*mm], hAlign="RIGHT")
tbl_tot.setStyle(TableStyle(borda()))
elems.append(tbl_tot)
elems.append(Spacer(1, 3*mm))

# ── DADOS DO PAGAMENTO ────────────────────────────────────────────────────
pagamentos = d.get("pagamentos", [])
if pagamentos:
    elems.append(secao("DADOS DO PAGAMENTO"))
    c_venc = 28*mm; c_vlr = 28*mm; c_obs = 40*mm
    c_forma = CW - c_venc - c_vlr - c_obs
    cwpag = [c_venc, c_vlr, c_forma, c_obs]
    th = [Paragraph("VENCIMENTO", st_th_c), Paragraph("VALOR", st_th_c),
          Paragraph("FORMA DE PAGAMENTO", st_th), Paragraph("OBSERVAÇÃO", st_th)]
    rows_pag = [th]
    for pv in pagamentos:
        pg = pv.get("pagamento", pv)
        rows_pag.append([
            Paragraph(fmt_date(pg.get("data_vencimento","")), st_td),
            Paragraph(brl(pg.get("valor",0)), st_td_r),
            Paragraph(str(pg.get("nome_forma_pagamento","")), st_td),
            Paragraph(str(pg.get("observacao","") or ""), st_td),
        ])
    tbl_pag = Table(rows_pag, colWidths=cwpag)
    tbl_pag.setStyle(TableStyle(borda() + [
        ("BACKGROUND", (0,0), (-1,0), CINZA_BG),
    ]))
    elems.append(tbl_pag)
    elems.append(Spacer(1, 4*mm))

# ── INTRODUÇÃO (Termos e Condições) ──────────────────────────────────────
intro = d.get("introducao", "")
if intro:
    elems.append(secao("TERMOS E CONDIÇÕES"))
    st_intro = ParagraphStyle("intro", fontSize=7, fontName="Helvetica",
                              textColor=PRETO, leading=9, alignment=TA_LEFT)
    intro_html = intro.replace("\n", "<br/>").replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;")
    elems.append(Spacer(1, 1*mm))
    elems.append(Paragraph(intro_html, st_intro))
    elems.append(Spacer(1, 4*mm))

# ── GERA PDF ──────────────────────────────────────────────────────────────
doc.build(elems)
out = os.path.join(os.path.dirname(__file__), f"orcamento_{d.get('codigo','teste')}.pdf")
with open(out, "wb") as f:
    f.write(buf.getvalue())

print(f"PDF gerado: {out}")
