# -*- coding: utf-8 -*-
"""
build_boletim.py — Gera o Boletim de Acompanhamento do Plano Rio Doce
Lê a Base Consolidada (xlsx) e produz um boletim.html autocontido, pronto
para impressão em A4 (Ctrl+P) ou publicação em página estática.
Uso: python build_boletim.py [caminho_do_xlsx] [saida.html]
"""
import sys, datetime
from collections import defaultdict, Counter
import openpyxl

import glob, os

def achar_planilha(alvo):
    """Aceita um arquivo .xlsx OU uma pasta: nesse caso pega o .xlsx mais recente,
    ignorando temporários do Excel (~$...). Assim o nome do arquivo pode mudar
    livremente — inclusive com prefixos como [CONSOLIDADO]."""
    if os.path.isfile(alvo):
        return alvo
    if os.path.isdir(alvo):
        cands = [f for f in glob.glob(os.path.join(alvo, "*.xlsx"))
                 if not os.path.basename(f).startswith("~$")]
        if not cands:
            sys.exit(f"ERRO: nenhum .xlsx encontrado em '{alvo}'. "
                     "Suba a planilha consolidada nessa pasta e rode de novo.")
        mais_novo = max(cands, key=os.path.getmtime)
        if len(cands) > 1:
            print(f"[aviso] {len(cands)} planilhas na pasta — usando a mais recente: {os.path.basename(mais_novo)}")
        return mais_novo
    sys.exit(f"ERRO: caminho não encontrado: '{alvo}'")

XLSX = achar_planilha(sys.argv[1] if len(sys.argv) > 1 else "dados")
OUT  = sys.argv[2] if len(sys.argv) > 2 else "boletim.html"
print(f"Planilha de origem: {os.path.basename(XLSX)}")

def brl(v, dec=2):
    s = f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"

def pct(x, dec=1):
    return f"{x:.{dec}f}".replace(".", ",") + "%"

# ── LEITURA ──────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(XLSX, data_only=True)
for aba in ("Base Consolidada", "Atualizações do Plano"):
    if aba not in wb.sheetnames:
        sys.exit(f"ERRO: a planilha não tem a aba '{aba}'. Abas encontradas: {', '.join(wb.sheetnames)}")
ws = wb["Base Consolidada"]
if "Valor Pago" not in [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]:
    sys.exit("ERRO: a aba 'Base Consolidada' não tem a coluna 'Valor Pago' — "
             "use a versão do consolidado que inclui essa coluna.")
H = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
IX = {n: H.index(n) + 1 for n in H}
def g(r, n):  return str(ws.cell(r, IX[n]).value or "").strip()
def gv(r, n): return ws.cell(r, IX[n]).value

rows = range(2, ws.max_row + 1)
OBRAS_ACOES = (1, 2, 6)

tot = pago = emp = 0.0
by_instr = defaultdict(lambda: [0.0, 0.0])
by_urs   = defaultdict(lambda: [0.0, 0.0])
benef    = defaultdict(lambda: [0.0, 0.0])
pg_int = pg_par = pg_emp = pg_sem = 0
obras_status = defaultdict(Counter); obras_val_travado = 0.0
equip_ok = equip_at = equip_na = 0; equip_at_quem = []
conv_grupos = defaultdict(lambda: [0, 0.0, []])
parciais = []
n_ubs = n_caps = 0

for r in rows:
    v  = gv(r, "Valor") or 0
    vp = gv(r, "Valor Pago")
    sp = g(r, "Status do pagamento")
    pg = vp if vp is not None else (v if sp == "Liquidado/Pago" else 0)
    rep, instr = g(r, "Status do Repasse"), g(r, "Instrumento de execução do recurso")
    ben, urs = g(r, "Beneficiário"), g(r, "URS")
    na = gv(r, "N° ação")

    tot += v; pago += pg
    if rep == "Empenhado": emp += max(0, v - pg)
    by_instr[instr][0] += v; by_instr[instr][1] += pg
    by_urs[urs][0] += v;     by_urs[urs][1] += pg
    benef[ben][0] += v;      benef[ben][1] += pg

    if sp == "Liquidado/Pago": pg_int += 1
    elif sp == "Liquidado/Pago PARCIAL":
        pg_par += 1
        parciais.append((ben, g(r, "Tipo de Aplicação"), v, pg))
    elif sp == "Empenhado": pg_emp += 1
    else: pg_sem += 1

    if instr == "Resolução - Investimento" and na in OBRAS_ACOES:
        st = g(r, "Status pré-repasse") or "Não iniciado"
        obras_status[st][ben] += 1
        tipo = g(r, "Tipo de Aplicação")
        if "UBS" in tipo: n_ubs += 1
        if "CAPS" in tipo: n_caps += 1
        if st in ("Aguardando envio dos documentos", "[CAPS] Aguardando PAR da RAPS", "Pleito em reavaliação"):
            obras_val_travado += v
    if instr == "Resolução - Investimento" and na not in OBRAS_ACOES:
        se = g(r, "Status de equipamentos (pré-repasse)")
        if se == "Lista validada": equip_ok += 1
        elif se == "Não se aplica" or not se or se == "-": equip_na += 1
        else:
            equip_at += 1; equip_at_quem.append(f"{ben} ({se})")
    if instr == "Convênio":
        sc = g(r, "Status de celebração de convênio (pré-repasse)") or "Pendente"
        conv_grupos[sc][0] += 1; conv_grupos[sc][1] += v; conv_grupos[sc][2].append(ben)

demais = max(0, tot - pago - emp)
comprometido = pago + emp
p_pago, p_emp, p_dem = 100*pago/tot, 100*emp/tot, 100*demais/tot
p_compr = 100*comprometido/tot
n_benef = len(benef)
n_linhas = ws.max_row - 1
n_obras = n_ubs + n_caps

# obras: agrupamento crítico / análise
ob_crit = {k: v for k, v in obras_status.items() if k in ("Aguardando envio dos documentos", "[CAPS] Aguardando PAR da RAPS", "Pleito em reavaliação", "Não iniciado")}
ob_anda = {k: v for k, v in obras_status.items() if k not in ob_crit}
ob_crit_n = sum(sum(c.values()) for c in ob_crit.values())
ob_anda_n = sum(sum(c.values()) for c in ob_anda.values())

def lista_munic(counter):
    return ", ".join(f"{m} ({q})" if q > 1 else m for m, q in sorted(counter.items()))

def lista_conv(nomes):
    c = Counter(nomes)
    return ", ".join(f"{m} ({q})" if q > 1 else m for m, q in sorted(c.items()))

# convênios: críticos vs tramitação
CONV_CRIT = ("Pleito em reavaliação", "Comunicado DCR", "Pendente")
conv_crit_n  = sum(v[0] for k, v in conv_grupos.items() if k in CONV_CRIT)
conv_crit_val = sum(v[1] for k, v in conv_grupos.items() if k in CONV_CRIT)
conv_tram_n  = sum(v[0] for k, v in conv_grupos.items() if k not in CONV_CRIT)
conv_tram_val = sum(v[1] for k, v in conv_grupos.items() if k not in CONV_CRIT)
conv_total_n = conv_crit_n + conv_tram_n
conv_total_val = by_instr["Convênio"][0]

# alterações
wa = wb["Atualizações do Plano"]
HA = [wa.cell(1, c).value for c in range(1, wa.max_column + 1)]
def ga(r, n): return str(wa.cell(r, HA.index(n) + 1).value or "").strip()
alt_fin, alt_and, alt_and_lista = 0, 0, []
for r in range(2, wa.max_row + 1):
    if not wa.cell(r, 1).value: continue
    st = ga(r, "Status da alteração")
    if st == "Finalizado": alt_fin += 1
    elif st == "Em andamento":
        alt_and += 1
        alt_and_lista.append(f"{ga(r,'Beneficiário')} ({ga(r,'Categoria de alteração de pleito').lower() or 'objeto'})")

# instrumentos em ordem de valor
INSTR_LABEL = {"Resolução - Investimento": "Resolução — Investimento",
               "Resolução - Custeio": "Resolução — Custeio",
               "Convênio": "Convênio", "Contrato": "Contrato"}
instr_rows = sorted(by_instr.items(), key=lambda x: -x[1][0])

INSTR_NOTA = {
    "Resolução - Investimento": f"{n_ubs} UBS e {n_caps} CAPS; equipamentos hospitalares, de APS e de consórcios.",
    "Resolução - Custeio": "VIGIAGUA, VIGIAR e VIGIDESASTRES; custeio hospitalar e de UBS.",
    "Convênio": "Obras de maior porte — hospitais, CEAE, polos de fisioterapia e SAE.",
    "Contrato": "Vigilância de agravos do desastre e monitoramento técnico de obras.",
}

urs_rows = sorted(((k, v) for k, v in by_urs.items()), key=lambda x: -x[1][0])
benef_100 = sorted([b for b, (v, p) in benef.items() if v > 0 and p >= v - 0.01])
n_benef_pagos = sum(1 for b, (v, p) in benef.items() if p > 0)

hoje = datetime.date.today()
MESES = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"]
data_ext = f"{hoje.day} de {MESES[hoje.month-1]} de {hoje.year}"
edicao = f"{hoje.month:02d}/{hoje.year}"

# ── FRAGMENTOS HTML ──────────────────────────────────────────────────────────
def minibar(p):
    return (f"<div class='mini'><div class='mini-fill' style='width:{min(100,p):.1f}%'></div></div>")

linhas_instr = "\n".join(
    f"""<tr>
      <td class='t-nome'>{INSTR_LABEL.get(k,k)}</td>
      <td class='t-num'>{brl(v)}</td>
      <td class='t-num'>{brl(p)}</td>
      <td class='t-exec'><span class='t-pct'>{pct(100*p/v if v else 0)}</span>{minibar(100*p/v if v else 0)}</td>
      <td class='t-nota'>{INSTR_NOTA.get(k,'')}</td>
    </tr>""" for k, (v, p) in instr_rows)

linhas_urs = "\n".join(
    f"""<tr><td class='t-nome'>{k}</td><td class='t-num'>{brl(v)}</td>
      <td class='t-num'>{brl(p)}</td>
      <td class='t-exec'><span class='t-pct'>{pct(100*p/v if v else 0)}</span>{minibar(100*p/v if v else 0)}</td></tr>"""
    for k, (v, p) in urs_rows)

linhas_parciais = "\n".join(
    f"""<tr><td class='t-nome'>{b}</td><td>{t}</td>
      <td class='t-num'>{brl(v)}</td><td class='t-num'>{brl(p)}</td>
      <td class='t-num'>{brl(v-p)}</td></tr>""" for b, t, v, p in sorted(parciais, key=lambda x: -(x[2]-x[3])))

def bloco_pend(titulo, qtd, sub, munics, tom):
    return f"""<div class='pend {tom}'>
      <div class='pend-head'><span class='pend-num'>{qtd}</span><div><div class='pend-t'>{titulo}</div><div class='pend-s'>{sub}</div></div></div>
      <div class='pend-m'>{munics}</div></div>"""

pendencias = ""
if "Aguardando envio dos documentos" in ob_crit:
    c = ob_crit["Aguardando envio dos documentos"]
    pendencias += bloco_pend("Unidades aguardando envio de documentos de obra",
        sum(c.values()), f"em {len(c)} municípios — a etapa está com as prefeituras",
        lista_munic(c), "rio")
if "[CAPS] Aguardando PAR da RAPS" in ob_crit:
    c = ob_crit["[CAPS] Aguardando PAR da RAPS"]
    pendencias += bloco_pend("CAPS aguardando aprovação do PAR junto à RAPS",
        sum(c.values()), "condição prévia à análise de engenharia", lista_munic(c), "rio")
pleitos_ob = obras_status.get("Pleito em reavaliação", Counter())
if pleitos_ob or conv_crit_n:
    partes = []
    if pleitos_ob:
        partes.append(("Obra: " if sum(pleitos_ob.values()) == 1 else "Obras: ") + lista_munic(pleitos_ob))
    if conv_crit_n:
        noms = sum((v[2] for k, v in conv_grupos.items() if k in CONV_CRIT), [])
        partes.append(("Convênio: " if conv_crit_n == 1 else "Convênios: ") + lista_conv(noms))
    munics = " &nbsp;·&nbsp; ".join(partes)
    pendencias += bloco_pend("Pleitos em reavaliação",
        sum(pleitos_ob.values()) + conv_crit_n,
        f"objeto/valor em rediscussão — {brl(conv_crit_val, 0)} em convênios aguardando a definição",
        munics, "rio")

tramitacao = ""
conv_tram_itens = sorted(((k, v) for k, v in conv_grupos.items() if k not in CONV_CRIT), key=lambda x: -x[1][1])
for k, (q, v, noms) in conv_tram_itens:
    tramitacao += f"""<tr><td class='t-nome'>{k}</td><td>Convênio</td><td class='t-num'>{q}</td>
      <td class='t-num'>{brl(v,0)}</td><td class='t-nota'>{lista_conv(noms)}</td></tr>"""
if ob_anda:
    for k, c in sorted(ob_anda.items(), key=lambda x: -sum(x[1].values())):
        tramitacao += f"""<tr><td class='t-nome'>{k}</td><td>Investimento — Obras</td><td class='t-num'>{sum(c.values())}</td>
          <td class='t-num'>—</td><td class='t-nota'>{lista_munic(c)}</td></tr>"""
if equip_at:
    tramitacao += f"""<tr><td class='t-nome'>Listas em análise</td><td>Investimento — Equipamentos</td><td class='t-num'>{equip_at}</td>
      <td class='t-num'>—</td><td class='t-nota'>{", ".join(equip_at_quem)}</td></tr>"""

# ── HTML ─────────────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Boletim de Acompanhamento — Plano de Ação em Saúde do Rio Doce</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@62..125,400..900&family=Source+Serif+4:ital,opsz,wght@0,8..60,400..700;1,8..60,400..700&display=swap" rel="stylesheet">
<style>
  :root {{
    --tinta:#1C2420; --mata:#1E6B47; --mata-escuro:#0F3D2A;
    --rio:#C2551B; --rio-claro:#FBEFE7;
    --ambar:#9A6A00; --ambar-claro:#FBF3DF;
    --verde-claro:#E8F2EC; --painel:#F2F4F0; --fio:#DCE3DB;
    --cinza:#5C6660; --cinza-claro:#8B948E;
  }}
  * {{ box-sizing:border-box; margin:0; }}
  html {{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }}
  body {{ font-family:'Source Serif 4', Georgia, serif; color:var(--tinta); background:#E6E9E4;
         font-size:11.5pt; line-height:1.62; }}
  .arch {{ font-family:'Archivo', 'Segoe UI', sans-serif; }}

  .pagina {{ background:#fff; width:210mm; min-height:297mm; margin:10mm auto; padding:16mm 17mm 14mm;
             box-shadow:0 2px 18px rgba(28,36,32,.14); position:relative; }}

  /* ── assinatura: o fio do rio (divisor mata→barro) ── */
  .fio-rio {{ height:2.5px; border:0; background:linear-gradient(90deg, var(--mata) 0%, var(--mata) 55%, var(--rio) 100%); margin:0 0 5mm; }}

  /* ── masthead ── */
  .cabo {{ display:flex; justify-content:space-between; align-items:flex-end; padding-bottom:4mm; }}
  .cabo-org {{ font-family:'Archivo'; font-size:8pt; font-weight:600; letter-spacing:.14em; text-transform:uppercase; color:var(--cinza); }}
  .cabo-num {{ font-family:'Archivo'; font-size:8pt; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--mata); text-align:right; }}
  .titulo {{ font-family:'Archivo'; font-weight:800; font-stretch:87%; font-size:24pt; line-height:1.04;
             letter-spacing:-.01em; color:var(--mata-escuro); margin:5mm 0 1.5mm; }}
  .titulo em {{ font-style:normal; color:var(--rio); }}
  .subtitulo {{ font-family:'Archivo'; font-size:10.5pt; font-weight:500; color:var(--cinza); margin-bottom:4mm; }}

  /* ── seções ── */
  .sec {{ margin-top:5.5mm; }}
  .sec-eyebrow {{ font-family:'Archivo'; font-size:8.5pt; font-weight:700; letter-spacing:.16em;
                  text-transform:uppercase; color:var(--mata); display:flex; align-items:center; gap:8px; margin-bottom:3.5mm; }}
  .sec-eyebrow::after {{ content:''; flex:1; height:1px; background:var(--fio); }}
  .lead {{ font-size:12.5pt; line-height:1.72; }}
  .lead b {{ font-weight:700; color:var(--mata-escuro); }}
  .lead .neg {{ color:var(--rio); font-weight:700; }}
  p + p {{ margin-top:3mm; }}
  .nota-analise {{ font-size:10.5pt; color:var(--cinza); border-left:2.5px solid var(--fio); padding-left:4mm; margin-top:4mm; }}

  /* ── o curso do recurso (hero) ── */
  .rio-wrap {{ margin:5mm 0 2mm; }}
  .rio-band {{ display:flex; height:15mm; border-radius:3px; overflow:hidden; }}
  .rio-seg {{ position:relative; display:flex; align-items:center; justify-content:center; }}
  .rio-seg .v {{ font-family:'Archivo'; font-weight:800; font-size:12.5pt; color:#fff; letter-spacing:-.01em; }}
  .rio-pago   {{ background:linear-gradient(180deg,#237A52,#1B5E3F); }}
  .rio-emp    {{ background:linear-gradient(180deg,#E0973B,#C97F22); }}
  .rio-dem    {{ background:linear-gradient(180deg,#CBD3CC,#B8C2BA); }}
  .rio-dem .v {{ color:var(--tinta); }}
  .rio-leg {{ display:flex; margin-top:2.5mm; font-family:'Archivo'; font-size:8.8pt; }}
  .rio-leg > div {{ padding-right:4mm; min-width:0; position:relative; }}
  .rio-leg > div::before {{ content:''; position:absolute; top:-2.5mm; left:0; width:6mm; height:1.5px; background:currentColor; opacity:.5; }}
  .rio-leg .l-t {{ font-weight:700; text-transform:uppercase; letter-spacing:.06em; font-size:7.5pt; }}
  .rio-leg .l-v {{ font-weight:600; color:var(--tinta); font-variant-numeric:tabular-nums; }}
  .rio-leg .l-t {{ line-height:1.25; }}
  .l-pago {{ color:var(--mata); }} .l-emp {{ color:var(--ambar); }} .l-dem {{ color:var(--cinza-claro); }}
  .l-pago .l-t {{ color:var(--mata); }} .l-emp .l-t {{ color:var(--ambar); }} .l-dem .l-t {{ color:var(--cinza-claro); }}
  .rio-marco {{ font-family:'Archivo'; font-size:8pt; color:var(--cinza); margin-top:2mm; }}
  .rio-marco b {{ color:var(--mata-escuro); }}

  /* ── KPIs ── */
  .kpis {{ display:flex; gap:3mm; margin-top:5mm; }}
  .kpi {{ flex:1; background:var(--painel); border-radius:3px; padding:3.2mm 3mm 2.8mm; text-align:center; }}
  .kpi-v {{ font-family:'Archivo'; font-weight:800; font-size:15pt; letter-spacing:-.02em; font-variant-numeric:tabular-nums; }}
  .kpi-l {{ font-family:'Archivo'; font-size:7.2pt; font-weight:600; letter-spacing:.05em; text-transform:uppercase; color:var(--cinza); margin-top:.8mm; line-height:1.3; }}

  /* ── tabelas ── */
  table {{ width:100%; border-collapse:collapse; font-family:'Archivo'; font-size:9.3pt; }}
  th {{ text-align:left; font-size:7.5pt; font-weight:700; letter-spacing:.09em; text-transform:uppercase;
       color:var(--cinza); padding:0 3mm 1.8mm 0; border-bottom:1.5px solid var(--tinta); }}
  td {{ padding:1.9mm 3mm 1.9mm 0; border-bottom:.5px solid var(--fio); vertical-align:top; }}
  tr:last-child td {{ border-bottom:none; }}
  .t-nome {{ font-weight:600; }}
  .t-num  {{ text-align:right; white-space:nowrap; font-variant-numeric:tabular-nums; }}
  th.t-num {{ text-align:right; }}
  .t-nota {{ font-family:'Source Serif 4'; font-size:9.3pt; color:var(--cinza); line-height:1.5; }}
  .t-exec {{ width:34mm; }}
  .t-pct {{ font-weight:700; font-variant-numeric:tabular-nums; display:block; margin-bottom:1mm; }}
  .mini {{ height:2.2mm; background:#E3E8E2; border-radius:2px; overflow:hidden; }}
  .mini-fill {{ height:100%; background:var(--mata); border-radius:2px; }}

  /* ── frentes (concluído / análise / pendente) ── */
  .frentes {{ display:grid; grid-template-columns:1fr 1fr; gap:3mm; margin-top:1mm; }}
  .frente {{ background:var(--painel); border-radius:3px; padding:3.2mm 4mm; }}
  .frente-t {{ font-family:'Archivo'; font-weight:700; font-size:10pt; margin-bottom:.5mm; }}
  .frente-s {{ font-family:'Archivo'; font-size:8pt; color:var(--cinza); margin-bottom:2.5mm; }}
  .stack {{ display:flex; height:4.5mm; border-radius:2px; overflow:hidden; margin-bottom:2mm; }}
  .st-ok {{ background:#237A52; }} .st-at {{ background:#E0973B; }} .st-cr {{ background:#C2551B; }} .st-vazio {{ background:#DDE3DC; }}
  .frente-leg {{ font-family:'Archivo'; font-size:8.4pt; color:var(--cinza); line-height:1.55; }}
  .frente-leg b {{ color:var(--tinta); font-variant-numeric:tabular-nums; }}
  .frente-obs {{ font-family:'Source Serif 4'; font-size:9.2pt; color:var(--tinta); margin-top:2.2mm;
                 padding-top:2.2mm; border-top:.5px solid var(--fio); line-height:1.55; }}

  /* ── pendências ── */
  .pend {{ border-radius:3px; padding:3.2mm 4.5mm; margin-bottom:2.5mm; }}
  .pend.rio {{ background:var(--rio-claro); border-left:3px solid var(--rio); }}
  .pend-head {{ display:flex; gap:4mm; align-items:baseline; }}
  .pend-num {{ font-family:'Archivo'; font-weight:800; font-size:17pt; color:var(--rio); min-width:9mm; font-variant-numeric:tabular-nums; }}
  .pend-t {{ font-family:'Archivo'; font-weight:700; font-size:10.3pt; }}
  .pend-s {{ font-family:'Archivo'; font-size:8.6pt; color:#8A4415; }}
  .pend-m {{ font-family:'Archivo'; font-size:8.6pt; color:var(--cinza); margin-top:1.8mm; padding-left:13mm; line-height:1.6; }}

  /* ── rodapé ── */
  .rodape {{ position:absolute; left:17mm; right:17mm; bottom:9mm; display:flex; justify-content:space-between;
             font-family:'Archivo'; font-size:7.5pt; color:var(--cinza-claro); border-top:.5px solid var(--fio); padding-top:2mm; }}

  /* ── impressão ── */
  .btn-print {{ position:fixed; top:14px; right:14px; z-index:9; font-family:'Archivo'; font-weight:700; font-size:13px;
                background:var(--mata); color:#fff; border:0; border-radius:22px; padding:10px 18px; cursor:pointer;
                box-shadow:0 3px 10px rgba(15,61,42,.35); }}
  .btn-print:hover {{ background:var(--mata-escuro); }}
  @page {{ size:A4; margin:0; }}
  @media print {{
    body {{ background:#fff; }}
    .pagina {{ margin:0; box-shadow:none; width:auto; min-height:auto; page-break-after:always; }}
    .pagina:last-child {{ page-break-after:auto; }}
    .btn-print {{ display:none; }}
  }}
  @media (max-width:820px) {{ .pagina {{ width:auto; min-height:auto; padding:8mm 6mm; }} .frentes {{ grid-template-columns:1fr; }} .kpis {{ flex-wrap:wrap; }} .kpi {{ min-width:28%; }} }}
</style>
</head>
<body>
<button class="btn-print" onclick="window.print()">🖨&nbsp; Imprimir / salvar PDF</button>

<!-- ═════════════════ PÁGINA 1 ═════════════════ -->
<div class="pagina">
  <div class="cabo">
    <div class="cabo-org">Secretaria de Estado de Saúde de Minas Gerais<br>Subsecretaria de Regionalização · CMIR</div>
    <div class="cabo-num">Boletim de acompanhamento<br>Edição {edicao} · {hoje.strftime("%d/%m/%Y")}</div>
  </div>
  <hr class="fio-rio">
  <div class="titulo">Plano de Ação em Saúde<br>do <em>Rio Doce</em></div>
  <div class="subtitulo arch">Acompanhamento da execução — {n_benef} beneficiários · {n_linhas} ações monitoradas · posição de {data_ext}</div>

  <div class="sec">
    <div class="sec-eyebrow arch">O curso do recurso</div>
    <div class="rio-wrap">
      <div class="rio-band">
        <div class="rio-seg rio-pago" style="width:{p_pago:.1f}%"><span class="v">{pct(p_pago)}</span></div>
        <div class="rio-seg rio-emp" style="width:{p_emp:.1f}%"><span class="v">{pct(p_emp)}</span></div>
        <div class="rio-seg rio-dem" style="flex:1"><span class="v">{pct(p_dem)}</span></div>
      </div>
      <div class="rio-leg">
        <div class="l-pago" style="width:{p_pago:.1f}%"><div class="l-t">Executado</div><div class="l-v">{brl(pago)}</div></div>
        <div class="l-emp" style="width:{p_emp:.1f}%"><div class="l-t">Empenhado · aguarda pagamento</div><div class="l-v">{brl(emp)}</div></div>
        <div class="l-dem" style="flex:1"><div class="l-t">Em etapas anteriores</div><div class="l-v">{brl(demais)}</div></div>
      </div>
      <div class="rio-marco">Do total de <b>{brl(tot, 0)}</b> do plano, <b>{brl(comprometido, 0)}</b> ({pct(p_compr)}) já estão pagos ou
      empenhados. Os pagamentos alcançam <b>{n_benef_pagos} dos {n_benef} beneficiários</b> — {len(benef_100)} deles com repasses
      integralmente quitados.</div>
    </div>
  </div>

  <div class="sec">
    <div class="sec-eyebrow arch">Execução por instrumento</div>
    <table>
      <thead><tr><th>Instrumento</th><th class="t-num">Previsto</th><th class="t-num">Pago</th><th>Execução</th><th>O que financia</th></tr></thead>
      <tbody>{linhas_instr}</tbody>
    </table>
  </div>

  <div class="sec">
    <div class="sec-eyebrow arch">Situação das etapas por frente</div>
    <div class="frentes">
      <div class="frente">
        <div class="frente-t">Obras <span style="color:var(--cinza);font-weight:500">· {n_obras} unidades</span></div>
        <div class="frente-s">análise pré-repasse, unidade a unidade</div>
        <div class="stack">
          <div class="st-at" style="width:{100*ob_anda_n/max(1,n_obras):.0f}%"></div>
          <div class="st-cr" style="flex:1"></div>
        </div>
        <div class="frente-leg"><b>0</b> liberadas · <b>{ob_anda_n}</b> em análise · <b style="color:var(--rio)">{ob_crit_n}</b> pendentes</div>
      </div>
      <div class="frente">
        <div class="frente-t">Equipamentos <span style="color:var(--cinza);font-weight:500">· {equip_ok + equip_at} listas</span></div>
        <div class="frente-s">validação das listas de aquisição</div>
        <div class="stack">
          <div class="st-ok" style="width:{100*equip_ok/max(1,equip_ok+equip_at):.0f}%"></div>
          <div class="st-at" style="flex:1"></div>
        </div>
        <div class="frente-leg"><b>{equip_ok}</b> validadas · <b>{equip_at}</b> em análise · <b>0</b> pendentes</div>
      </div>
      <div class="frente">
        <div class="frente-t">Convênios <span style="color:var(--cinza);font-weight:500">· {conv_total_n} instrumentos</span></div>
        <div class="frente-s">celebração pré-repasse</div>
        <div class="stack">
          <div class="st-at" style="width:{100*conv_tram_n/max(1,conv_total_n):.0f}%"></div>
          <div class="st-cr" style="flex:1"></div>
        </div>
        <div class="frente-leg"><b>0</b> publicados · <b>{conv_tram_n}</b> em tramitação · <b style="color:var(--rio)">{conv_crit_n}</b> em reavaliação</div>
      </div>
      <div class="frente">
        <div class="frente-t">Pagamentos <span style="color:var(--cinza);font-weight:500">· {n_linhas} ações</span></div>
        <div class="frente-s">todas as ações do plano</div>
        <div class="stack">
          <div class="st-ok" style="width:{100*(pg_int+pg_par)/n_linhas:.0f}%"></div>
          <div class="st-at" style="width:{100*pg_emp/n_linhas:.0f}%"></div>
          <div class="st-vazio" style="flex:1"></div>
        </div>
        <div class="frente-leg"><b>{pg_int}</b> pagas + <b>{pg_par}</b> parciais · <b>{pg_emp}</b> empenhadas · <b>{pg_sem}</b> sem repasse iniciado</div>
      </div>
    </div>
  </div>
  <div class="rodape"><span>Plano de Ação em Saúde do Rio Doce — Boletim de Acompanhamento</span><span>1 / 3</span></div>
</div>

<!-- ═════════════════ PÁGINA 2 ═════════════════ -->
<div class="pagina">
  <div class="sec" style="margin-top:0">
    <div class="sec-eyebrow arch">Em tramitação regular</div>
    <table>
      <thead><tr><th>Etapa</th><th>Instrumento</th><th class="t-num">Itens</th><th class="t-num">Valor</th><th>Beneficiários</th></tr></thead>
      <tbody>{tramitacao}</tbody>
    </table>
  </div>

  <div class="sec" >
    <div class="sec-eyebrow arch" style="color:var(--rio)">Onde o plano precisa de providência</div>
    {pendencias}
  </div>

  <div class="rodape"><span>Plano de Ação em Saúde do Rio Doce — Boletim de Acompanhamento</span><span>2 / 3</span></div>
</div>

<!-- ═════════════════ PÁGINA 3 ═════════════════ -->
<div class="pagina">
  <div class="sec" style="margin-top:0">
    <div class="sec-eyebrow arch">Pagamentos parciais em aberto</div>
    <table>
      <thead><tr><th>Beneficiário</th><th>Ação</th><th class="t-num">Previsto</th><th class="t-num">Pago</th><th class="t-num">Saldo</th></tr></thead>
      <tbody>{linhas_parciais}</tbody>
    </table>
  </div>
  <div class="sec">
    <div class="sec-eyebrow arch">Execução por unidade regional de saúde</div>
    <table>
      <thead><tr><th>URS</th><th class="t-num">Previsto</th><th class="t-num">Pago</th><th>Execução</th></tr></thead>
      <tbody>{linhas_urs}</tbody>
    </table>
  </div>

  <div class="sec">
    <div class="sec-eyebrow arch">Atualizações do plano</div>
    <p style="font-size:10.5pt">Foram registradas <b>{alt_fin + alt_and} atualizações</b> de pleito desde o início do plano —
    <b>{alt_fin} finalizadas</b> e <b>{alt_and} em andamento</b>{": " + "; ".join(alt_and_lista) if alt_and_lista else ""}.
    Os pleitos em reavaliação de Aimorés, Conselheiro Pena e Ponte Nova vinculam-se às atualizações em curso.</p>
  </div>

  <div class="rodape">
    <span>Fonte: Base Consolidada do Plano Rio Doce · Elaboração: CMIR / Subsecretaria de Regionalização — SES-MG</span>
    <span>3 / 3</span>
  </div>
</div>
</body>
</html>"""

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)
print(f"gerado: {OUT} ({len(html):,} chars)")
