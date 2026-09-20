#!/usr/bin/env python3
"""
update_dashboard.py
====================================================================
Atualiza o dashboard executivo da Porto com uma nova planilha de
resultados trimestrais divulgada pela área de RI.

FLUXO:
    1. Lê a planilha de resultados (.xlsx) informada
    2. Localiza cada série pelo RÓTULO da linha (coluna A), não pelo
       número da linha — resiliente a pequenos deslocamentos que a
       Porto às vezes introduz entre um trimestre e outro
    3. Extrai a série completa (1T23 em diante) para consolidado,
       as 4 verticais, indicadores operacionais, mercado e capital
    4. Realinha todas as séries por RÓTULO de trimestre (não por
       posição) — evita erros quando abas diferentes têm históricos
       de tamanhos diferentes
    5. Roda validações de sanidade contra a última publicação
    6. Grava data/dashboard.json e injeta no template HTML

USO:
    python update_dashboard.py Planilha-2T26.xlsx
    python update_dashboard.py Planilha-2T26.xlsx --template dashboard.html --out dashboard.html

REQUISITOS:
    pip install openpyxl --break-system-packages

MANUTENÇÃO — quando a Porto muda o layout da planilha:
    Se uma extração falhar com "rótulo não encontrado", a Porto
    normalmente renomeou/moveu a linha. Abra a planilha, ache o
    novo texto exato da linha (coluna A) e ajuste a constante
    SERIES abaixo — não é necessário mexer no resto do script.

    O schema de saída (chaves do JSON) é fixo — é o que o dashboard
    HTML espera. Não renomeie as chaves à esquerda em SERIES.
====================================================================
"""
import argparse
import json
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Faltou instalar a dependência: pip install openpyxl --break-system-packages")


# ============================================================
# 1. LOCALIZAÇÃO DAS SÉRIES NA PLANILHA
# ============================================================
# path_no_json : (aba, rótulo da linha, ocorrência [0-indexado, default 0], escala [default 1.0])
#
# "ocorrência" existe porque alguns rótulos se repetem em mais de um
# bloco da planilha (ex.: "Lucro Líquido" aparece no Consolidado E em
# cada vertical) — a ordem é sempre de cima para baixo na aba.

SERIES = {
    # ================= CONSOLIDADO =================
    "consolidado.receita_total":        ("DREs Verticais", "Receita Total (Prêmio Retido + Receitas Demais)", 0),
    "consolidado.sinistros":            ("DREs Verticais", "Sinistros Líquidos Retidos", 0),
    "consolidado.perdas_credito":       ("DREs Verticais", "Perdas de Crédito", 0),
    "consolidado.desp_comercializacao": ("DREs Verticais", "Despesa de Comercialização", 0),
    "consolidado.desp_tributos":        ("DREs Verticais", "Despesas com Tributos", 0),
    "consolidado.desp_operacionais":    ("DREs Verticais", "Despesas Operacionais", 0),
    "consolidado.desp_administrativas": ("DREs Verticais", "Despesas Administrativas", 0),
    "consolidado.lucro_operacional":    ("DREs Verticais", "Lucro Operacional", 0),
    "consolidado.resultado_financeiro": ("DREs Verticais", "Resultado Financeiro e Patrimonial", 0),
    "consolidado.lair":                 ("DREs Verticais", "LAIR", 0),
    "consolidado.ir_csll":              ("DREs Verticais", "Imposto de Renda e Contribuição Social", 0),
    "consolidado.participacao_resultados": ("DREs Verticais", "Participação nos Resultados", 0),
    "consolidado.lucro_liquido":        ("DREs Verticais", "Lucro Líquido", 0),
    "consolidado.pl_medio":             ("DREs Verticais", "Patrimônio Líquido Médio", 0),
    "consolidado.roae":                 ("DREs Verticais", "R.O.A.E. (%)", 0),

    # ================= PORTO SEGURO (vertical) =================
    "porto_seguro.receita_total":       ("DREs Verticais", "Receita Total (Prêmio Retido  +  Receitas)", 0),
    "porto_seguro.sinistros":           ("DREs Verticais", "Sinistros Líquidos Retidos", 1),
    "porto_seguro.lucro_liquido":       ("DREs Verticais", "Lucro Líquido", 1),
    "porto_seguro.roae":                ("DREs Verticais", "R.O.A.E. (%)", 1),
    "porto_seguro.pl_medio":            ("DREs Verticais", "Patrimônio Líquido Médio", 1),
    "porto_seguro.sinistralidade":      ("Indicadores Oper. e Fin.", "Sinistralidade", 0),
    "porto_seguro.resultado_operacional": ("DREs Verticais", "Resultado Operacional", 0),
    "porto_seguro.ga_indice":           ("DREs Verticais", "Índice G&A Vertical", 0),

    # ================= PORTO SAÚDE (vertical) =================
    "porto_saude.receita_total":        ("DREs Verticais", "Receita Total (Prêmio Retido  +  Receitas)", 1),
    "porto_saude.premios_emitidos":     ("DREs Verticais", "Receita Total (Prêmio Retido  +  Receitas)", 1),
    "porto_saude.sinistros":            ("DREs Verticais", "Sinistros Líquidos Retidos", 2),
    "porto_saude.lucro_liquido":        ("DREs Verticais", "Lucro Líquido", 2),
    "porto_saude.roae":                 ("DREs Verticais", "R.O.A.E. (%)", 2),
    "porto_saude.pl_medio":             ("DREs Verticais", "Patrimônio Líquido Médio", 2),
    "porto_saude.sinistralidade":       ("Indicadores Oper. e Fin.", "Índice de Sinistralidade", 0),
    "porto_saude.resultado_operacional": ("DREs Verticais", "Resultado Operacional", 1),

    # ================= PORTO BANK (vertical) =================
    "porto_bank.total_receitas":        ("DREs Verticais", "Total Receitas", 0),
    "porto_bank.receita_liquida":       ("DREs Verticais", "Receita Liquida", 0),
    "porto_bank.perdas_credito":        ("DREs Verticais", "Perdas de Crédito (ii)", 0),
    "porto_bank.lucro_liquido":         ("DREs Verticais", "Lucro Líquido", 3),
    "porto_bank.roae":                  ("DREs Verticais", "R.O.A.E. (%)", 3),
    "porto_bank.pl_medio":              ("DREs Verticais", "Patrimônio Líquido Médio", 3),
    "porto_bank.indice_eficiencia":     ("DREs Verticais", "Índice de Eficiência (%)", 0),
    "porto_bank.resultado_antes_impostos": ("DREs Verticais", "Resultado antes dos Impostos", 2),

    # ================= PORTO SERVIÇO (vertical) =================
    "porto_servico.receitas":           ("DREs Verticais", "Receitas com Serviços", 0),
    "porto_servico.lucro_liquido":      ("DREs Verticais", "Lucro (Prejuízo) Líquido", 0),
    "porto_servico.roae":               ("DREs Verticais", "ROAE (%)", 0),
    "porto_servico.resultado_operacional": ("DREs Verticais", "Resultado Operacional", 2),
    "porto_servico.pl_medio":           ("DREs Verticais", "Patrimônio Líquido Médio", 4),

    # ================= CONTROLADORA E DEMAIS =================
    "controladora.receita_total":       ("DREs Verticais", "Receita Total (Prêmio Retido + Receitas Demais)", 1),
    "controladora.lucro_liquido":       ("DREs Verticais", "Lucro Líquido", 4),

    # ================= RECEITA POR SEGMENTO (aba "Receitas") =================
    "receita_por_segmento.auto":            ("Receitas", "Auto", 0),
    "receita_por_segmento.patrimoniais":    ("Receitas", "Patrimoniais", 0),
    "receita_por_segmento.vida":            ("Receitas", "Vida", 0),
    "receita_por_segmento.uruguai_seguros": ("Receitas", "Uruguai Seguros", 0),
    "receita_por_segmento.uruguai_servicos": ("Receitas", "Uruguai Serviços", 0),
    "receita_por_segmento.outros_seguros":  ("Receitas", "Outros Seguros", 0),

    # ================= ÍNDICES — Porto Seguro / Saúde =================
    "porto_seguro_indices.sinistralidade":  ("Indicadores Oper. e Fin.", "Sinistralidade", 0),
    "porto_seguro_indices.comissionamento": ("Indicadores Oper. e Fin.", "Índice de Comissionamento", 0),
    "porto_seguro_indices.desp_administrativas": ("Indicadores Oper. e Fin.", "Índice de Despesas Administrativas", 0),
    "porto_seguro_indices.outras_op":       ("Indicadores Oper. e Fin.", "Índice de Outras Receitas e Despesas Operacionais", 0),
    "porto_seguro_indices.tributos":        ("Indicadores Oper. e Fin.", "Índice de Tributos", 0),
    "porto_seguro_indices.combinado":       ("Indicadores Oper. e Fin.", "Índice Combinado", 0),
    "porto_seguro_indices.combinado_ampliado": ("Indicadores Oper. e Fin.", "Índice Combinado Ampliado", 0),

    "saude_indices.sinistralidade":     ("Indicadores Oper. e Fin.", "Índice de Sinistralidade", 0),
    "saude_indices.combinado":          ("Indicadores Oper. e Fin.", "Índice Combinado (Saúde + Odonto)", 0),
    "saude_indices.combinado_ampliado": ("Indicadores Oper. e Fin.", "Índice Combinado Ampliado (Saúde + Odonto)", 0),

    # ================= BANK / SERVIÇO — índices =================
    "bank_indices.carteira_credito_total_bi": ("Indicadores Oper. e Fin.", "Carteira de Crédito (R$ bilhões)", 0),  # já consolida Cartão + Financiamento
    "bank_indices.over90":              ("Indicadores Oper. e Fin.", "Atrasos acima de 90 dias (Over 90 - 360d)", 0),
    "bank_indices.custo_risco":         ("Indicadores Oper. e Fin.", "Custo de Risco", 0),
    "bank_indices.npl_formation_pct":   ("Indicadores Oper. e Fin.", "NPL Formation", 0),  # versão em %; ha tambem uma em R$ milhoes com nome parecido — nao usar
    "servico_indices.margem_ebitda":    ("Indicadores Oper. e Fin.", "Margem Ebitda", 0),

    # ================= EFICIÊNCIA OPERACIONAL =================
    "eficiencia_operacional.desp_admin":     ("Indicadores Oper. e Fin.", "Despesas Administrativas", 0),
    "eficiencia_operacional.receita_total":  ("Indicadores Oper. e Fin.", "Receita Total", 0),
    "eficiencia_operacional.indice_eficiencia": ("Indicadores Oper. e Fin.", "Índice de Eficiência Operacional", 0),

    # ================= RESULTADO FINANCEIRO CONSOLIDADO =================
    "resultado_financeiro_consolidado.resultado_aplicacoes": ("Indicadores Oper. e Fin.", "Resultado de Aplicações Financeiras", 0),
    "resultado_financeiro_consolidado.total_ex_prev": ("Indicadores Oper. e Fin.", "Total (ex previdência)", 0),
    "resultado_financeiro_consolidado.total": ("Indicadores Oper. e Fin.", "Resultado Financeiro Total", 0),

    # ================= MERCADO =================
    "mercado.pl_medio_milhares":  ("Indicadores Oper. e Fin.", "Patrimônio Líquido Médio (R$ milhares)", 0),
    "mercado.roae":               ("Indicadores Oper. e Fin.", "Rentabilidade sobre o Patrimônio (ROAE)", 0),
    "mercado.lpa_sem_bc":         ("Indicadores Oper. e Fin.", "Lucro por Ação s/ Business Combination", 0),
    "mercado.lpa_com_bc":         ("Indicadores Oper. e Fin.", "Lucro por Ação c/ Business Combination", 0),
    "mercado.qtd_acoes_milhares": ("Indicadores Oper. e Fin.", 'Quantidade de Ações (milhares, considerados os efeitos do "split" em mar/08 e bonificação de out/21)', 0),
    "mercado.jcp_bruto":          ("Indicadores Oper. e Fin.", "JCP Bruto (R$ milhares)", 0, 0.001),
    "mercado.dividendos":         ("Indicadores Oper. e Fin.", "Dividendos (R$ milhares)", 0, 0.001),
    # mercado.vpa é CALCULADO (não vem de uma linha própria) — ver build_derived()

    # ================= CAPITAL =================
    "capital.capital_regulatorio_seguradoras": ("Indicadores Oper. e Fin.", "Capital Regulatório Seguradoras (R$ milhões)", 0),
    "capital.suf_capital_seguradoras":  ("Indicadores Oper. e Fin.", "Suficiência de Capital Seguradoras (R$ milhões)", 0),
    "capital.capital_regulatorio_financeiras": ("Indicadores Oper. e Fin.", "Capital Regulatório Financeiras (R$ milhões)", 0),
    "capital.suf_capital_financeiras":  ("Indicadores Oper. e Fin.", "Suficiência de Capital Financeiras (R$ milhões)", 0),
    "capital.necessidade_capital_total": ("Indicadores Oper. e Fin.", "Necessidade de Capital - Total (R$ milhão)", 0),
    "capital.suf_capital_total":        ("Indicadores Oper. e Fin.", "Suficiência de Capital - Total (R$ milhões)", 0),
    "capital.pla":                      ("Indicadores Oper. e Fin.", "PLA (R$ milhões)", 0),

    # ================= TOP-LEVEL (fora de qualquer bloco) =================
    "investimentos_capex":         ("Indicadores Oper. e Fin.", "Investimentos (R$ milhares)", 0),
    "itens_segurados_auto_total":  ("Negócios-Clientes", "Itens Segurados - Total (Porto Seguro e Azul) Auto (em mil)", 0),
    "vidas_seguradas_pessoas":     ("Negócios-Clientes", "Vidas Seguradas - Pessoas (em mil)", 0),
    "vidas_saude":                 ("Negócios-Clientes", "Número de Vidas Seguradas (em mil)", 0),
    "vidas_odonto":                ("Negócios-Clientes", "Número de Vidas Seguradas (em mil)", 1),
    "cartao_credito_unidades_mil": ("Negócios-Clientes", "Cartão de Crédito (milhares de unidades)", 0),
    "marketshare_total_auto":      ("Produtos", "Marketshare (Prêmios) - Total Automóvel (Porto + Azul + Itaú Auto e Residência + Mitsui)*", 0),
    "marketshare_porto_auto":      ("Produtos", "Marketshare (Prêmios) - Porto Seguro - Auto", 0),
    "sinistralidade_porto_auto_produto": ("Produtos", "Índice de Sinistralidade - Auto", 0),  # descontinuada pela Porto a partir do 1T25 (fica 0)
    "caixa_milhoes":               ("IFRS-4-Balanço", "Caixa e equivalentes de caixa", 0, 0.001),
    "total_ativo_milhoes":         ("IFRS-4-Balanço", "TOTAL DO ATIVO", 0, 0.001),
    "pl_total_milhoes":            ("IFRS-4-Balanço", "Total do patrimônio líquido", 0, 0.001),
}

# "total_negocios_ativos_consorcio" é a SOMA de imóveis + veículos — extraído à parte
CONSORCIO_IMOVEIS = ("Negócios-Clientes", "Negócios Ativos Imóveis (mil)", 0)
CONSORCIO_VEICULOS = ("Negócios-Clientes", "Negócios Ativos Veículos (mil)", 0)


# ============================================================
# 2. LEITURA DA PLANILHA
# ============================================================

def normalize(s):
    """Normaliza um rótulo para comparação: colapsa espaços internos e remove os das pontas."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def find_quarter_columns(ws):
    """Localiza a linha de cabeçalho (a que contém '1T23', '2T23'...) em
    QUALQUER lugar das primeiras 10 linhas e devolve {'1T23': col, ...}."""
    quarter_re = re.compile(r"^[1-4]T\d{2}$")
    for hr in range(1, 11):
        cols = {}
        for c in range(1, ws.max_column + 1):
            v = normalize(ws.cell(hr, c).value)
            if quarter_re.match(v):
                cols[v] = c
        if len(cols) >= 4:
            return cols
    raise RuntimeError(f"Não encontrei linha de cabeçalho de trimestres na aba '{ws.title}'.")


def find_label_row(ws, label, occurrence=0):
    """Encontra a occurrence-ésima linha (0-indexada, de cima para baixo)
    cuja coluna A normaliza para o mesmo texto que `label`."""
    target = normalize(label)
    hits = [r for r in range(1, ws.max_row + 1) if normalize(ws.cell(r, 1).value) == target]
    if not hits:
        raise RuntimeError(f"Rótulo não encontrado na aba '{ws.title}': '{label}'")
    if occurrence >= len(hits):
        raise RuntimeError(
            f"Rótulo '{label}' encontrado {len(hits)}x na aba '{ws.title}', "
            f"mas occurrence={occurrence} foi pedido (linhas: {hits})."
        )
    return hits[occurrence]


def extract_series_dict(wb, sheet, label, occurrence=0, scale=1.0):
    """Extrai uma linha como {'1T23': valor, ...} — dict, não lista, para
    permitir realinhamento correto por rótulo de trimestre depois."""
    ws = wb[sheet]
    quarters = find_quarter_columns(ws)
    row = find_label_row(ws, label, occurrence)
    out = {}
    for q, col in quarters.items():
        raw = ws.cell(row, col).value
        if raw is None or (isinstance(raw, str) and raw.strip() in ("-", "", "n/a", "N/A")):
            out[q] = None
        else:
            out[q] = float(raw) * scale
    return out


def load_all_series(xlsx_path, reference_start="1T23"):
    """Roda toda a extração e devolve (data_aninhado, lista_de_trimestres, workbook)."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    raw_by_path = {}
    all_quarters = set()
    errors = []

    def try_extract(path, spec):
        sheet, label = spec[0], spec[1]
        occurrence = spec[2] if len(spec) > 2 else 0
        scale = spec[3] if len(spec) > 3 else 1.0
        try:
            d = extract_series_dict(wb, sheet, label, occurrence, scale)
            raw_by_path[path] = d
            all_quarters.update(d.keys())
        except RuntimeError as e:
            errors.append(str(e))

    for path, spec in SERIES.items():
        try_extract(path, spec)

    try:
        imoveis = extract_series_dict(wb, *CONSORCIO_IMOVEIS)
        veiculos = extract_series_dict(wb, *CONSORCIO_VEICULOS)
        soma = {q: (imoveis.get(q) or 0) + (veiculos.get(q) or 0) for q in set(imoveis) | set(veiculos)}
        raw_by_path["total_negocios_ativos_consorcio"] = soma
        all_quarters.update(soma.keys())
    except RuntimeError as e:
        errors.append(str(e))

    if errors:
        print("\n⚠️  AVISOS DE EXTRAÇÃO (revise antes de publicar):", file=sys.stderr)
        for e in errors:
            print("   -", e, file=sys.stderr)

    def qkey(q):
        return (int("20" + q[-2:]), int(q[0]))
    ordered_all = sorted(all_quarters, key=qkey)
    if reference_start not in ordered_all:
        raise RuntimeError(f"Trimestre de referência '{reference_start}' não encontrado em nenhuma série.")
    quarters_ref = ordered_all[ordered_all.index(reference_start):]

    data = {}
    for path, series_dict in raw_by_path.items():
        values = [series_dict.get(q) for q in quarters_ref]
        node = data
        parts = path.split(".")
        for p in parts[:-1]:
            node = node.setdefault(p, {})
        node[parts[-1]] = values

    return data, quarters_ref, wb


def get(data, path, default=None):
    node = data
    for p in path.split("."):
        if not isinstance(node, dict) or p not in node:
            return default
        node = node[p]
    return node


# ============================================================
# 3. CAMPOS DERIVADOS (não vêm de uma linha — são calculados)
# ============================================================

def build_derived(data):
    """VPA (Valor Patrimonial por Ação) = Patrimônio Líquido Total / Qtd. de Ações.
    PL total em milhões, quantidade de ações em milhares na planilha —
    o resultado sai em R$/ação."""
    pl_total_mi = get(data, "pl_total_milhoes")
    qtd_acoes_mil = get(data, "mercado.qtd_acoes_milhares")
    vpa = []
    if pl_total_mi and qtd_acoes_mil:
        for pl, qtd in zip(pl_total_mi, qtd_acoes_mil):
            vpa.append(round(pl * 1000 / qtd, 4) if (pl is not None and qtd) else None)
    data.setdefault("mercado", {})["vpa"] = vpa


def build_receita_por_segmento_totais(data):
    """*_total do bloco receita_por_segmento reaproveitam as séries de
    receita já extraídas por vertical — evita reler a planilha."""
    seg = data.setdefault("receita_por_segmento", {})
    seg["porto_seguro_total"] = get(data, "porto_seguro.receita_total")
    seg["porto_saude_total"] = get(data, "porto_saude.receita_total")
    seg["porto_bank_total"] = get(data, "porto_bank.total_receitas")
    seg["porto_servico_total"] = get(data, "porto_servico.receitas")
    seg["controladora_total"] = get(data, "controladora.receita_total")


# ============================================================
# 4. VALIDAÇÕES DE SANIDADE
# ============================================================

def sanity_check(data, quarters):
    problems = []

    def last(path):
        v = get(data, path)
        return v[-1] if v else None

    rec, luc, roae, combined = (last("consolidado.receita_total"), last("consolidado.lucro_liquido"),
                                 last("consolidado.roae"), last("porto_seguro_indices.combinado"))

    if rec is None or not (3000 < rec < 30000):
        problems.append(f"Receita total consolidada fora do intervalo plausível: {rec}")
    if luc is None or not (-5000 < luc < 5000):
        problems.append(f"Lucro líquido consolidado fora do intervalo plausível: {luc}")
    if roae is None or not (-0.5 < roae < 0.6):
        problems.append(f"ROAE fora do intervalo plausível (fração, ex. 0.29): {roae}")
    if combined is None or not (0.5 < combined < 1.3):
        problems.append(f"Combined Ratio fora do intervalo plausível (fração, ex. 0.887): {combined}")

    soma_vert = sum(v for v in [last("porto_seguro.receita_total"), last("porto_saude.receita_total"),
                                 last("porto_bank.total_receitas"), last("porto_servico.receitas")] if v)
    if rec and soma_vert and not (0.5 * rec < soma_vert < 1.3 * rec):
        problems.append(f"Soma das verticais ({soma_vert:.0f}) muito distante da receita consolidada ({rec:.0f}).")

    if problems:
        print("\n🚫 VALIDAÇÃO FALHOU — revise antes de publicar:", file=sys.stderr)
        for p in problems:
            print("   -", p, file=sys.stderr)
        return False

    print(f"\n✅ Validação OK — último trimestre: {quarters[-1]}")
    print(f"   Receita: R$ {rec:,.0f} mi | Lucro: R$ {luc:,.0f} mi | ROAE: {roae*100:.1f}% | Combined: {combined*100:.1f}%")
    return True


# ============================================================
# 5. MONTAGEM DO JSON FINAL (schema fixo, igual ao já publicado) E INJEÇÃO NO HTML
# ============================================================

def build_json(data, quarters):
    build_derived(data)
    build_receita_por_segmento_totais(data)
    out = {"quarters": quarters}
    out.update(data)
    return out


def round_floats(obj, nd=4):
    if isinstance(obj, dict):
        return {k: round_floats(v, nd) for k, v in obj.items()}
    if isinstance(obj, list):
        return [round_floats(v, nd) for v in obj]
    if isinstance(obj, float):
        return round(obj, nd)
    return obj


def inject_into_template(template_path, json_data, out_path):
    html = Path(template_path).read_text(encoding="utf-8")
    payload = json.dumps(round_floats(json_data), ensure_ascii=False, separators=(",", ":"))
    pattern = re.compile(r'(<script id="raw-data" type="application/json">)(.*?)(</script>)', re.DOTALL)
    new_html, n = pattern.subn(lambda m: m.group(1) + payload + m.group(3), html)
    if n == 0:
        raise RuntimeError('Não encontrei <script id="raw-data"> no template. Confirme --template.')
    Path(out_path).write_text(new_html, encoding="utf-8")
    print(f"✅ Dashboard gerado: {out_path}")


# ============================================================
# 6. CLI
# ============================================================

def main():
    ap = argparse.ArgumentParser(description="Atualiza o dashboard Porto com uma nova planilha trimestral.")
    ap.add_argument("planilha", help="Caminho da planilha .xlsx de resultados")
    ap.add_argument("--template", default="dashboard.html", help="HTML-base a atualizar")
    ap.add_argument("--out", default="dashboard.html", help="Arquivo de saída")
    ap.add_argument("--json-out", default="data/dashboard.json", help="Onde salvar o JSON extraído")
    ap.add_argument("--skip-validation", action="store_true")
    args = ap.parse_args()

    print(f"📥 Lendo {args.planilha} ...")
    data, quarters, _wb = load_all_series(args.planilha)

    raw = build_json(data, quarters)

    ok = sanity_check(raw, quarters)
    if not ok and not args.skip_validation:
        sys.exit("\nExtração interrompida. Corrija os mapeamentos em SERIES ou rode com --skip-validation.")

    Path(args.json_out).parent.mkdir(parents=True, exist_ok=True)
    Path(args.json_out).write_text(json.dumps(round_floats(raw), ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"💾 JSON salvo em {args.json_out}")

    inject_into_template(args.template, raw, args.out)


if __name__ == "__main__":
    main()
