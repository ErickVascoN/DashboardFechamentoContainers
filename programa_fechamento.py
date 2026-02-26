# -*- coding: utf-8 -*-
"""
Programa de Fechamento de Containers
Interface gráfica para preencher dados e gerar dashboard HTML automaticamente.
"""

import csv
import os
import sys
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path

# Força UTF-8 no terminal Windows
if sys.stdout:
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

PASTA_BASE = Path(__file__).parent


# ═══════════════════════════════════════════════════════════════════════════════
#  FUNÇÕES AUXILIARES
# ═══════════════════════════════════════════════════════════════════════════════

def parse_numero(valor: str):
    """Converte string numérica brasileira (vírgula decimal) para float."""
    if valor is None:
        return None
    valor = str(valor).strip()
    if not valor or valor == "-":
        return None
    valor = valor.replace(" ", "").replace(".", "")
    valor = valor.replace(",", ".")
    valor = valor.replace("%", "")
    try:
        return float(valor)
    except ValueError:
        return None


def fmt_numero(valor, decimais=2):
    """Formata número para exibição brasileira."""
    if valor is None:
        return "-"
    return f"{valor:,.{decimais}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(valor):
    """Formata percentual."""
    if valor is None:
        return "-"
    return f"{valor:.2f}%"


def safe_float(entry_widget, default=None):
    """Extrai float seguro de um Entry widget."""
    texto = entry_widget.get().strip()
    if not texto or texto == "-":
        return default
    texto = texto.replace(".", "").replace(",", ".").replace("%", "").replace(" ", "")
    try:
        return float(texto)
    except ValueError:
        return default


def safe_int(entry_widget, default=None):
    """Extrai int seguro de um Entry widget."""
    val = safe_float(entry_widget)
    if val is None:
        return default
    return int(val)


# ═══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DO HTML (DASHBOARD)
# ═══════════════════════════════════════════════════════════════════════════════

def gerar_html(dados: dict) -> str:
    """Gera relatório HTML completo com tabelas estilizadas e gráficos."""

    diferenca_peso = (dados["peso_real"]["peso_unitario"] - dados["peso_teorico"]["peso_unitario"])
    diferenca_kgs = (dados["peso_real"]["total_kgs"] - dados["peso_teorico"]["total_kgs"])

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Fechamento De Container</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        :root {{
            --primary: #1a73e8;
            --primary-light: #e8f0fe;
            --success: #0d904f;
            --danger: #d93025;
            --warning: #f9ab00;
            --bg: #f8f9fa;
            --card-bg: #ffffff;
            --text: #202124;
            --text-secondary: #5f6368;
            --border: #dadce0;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: var(--bg);
            color: var(--text);
            padding: 20px;
            line-height: 1.6;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        .header {{
            background: linear-gradient(135deg, var(--primary), #1557b0);
            color: white;
            padding: 30px 40px;
            border-radius: 12px;
            margin-bottom: 24px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .header h1 {{ font-size: 28px; font-weight: 600; }}
        .header p {{ opacity: 0.9; margin-top: 4px; font-size: 14px; }}
        .grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(420px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
        }}
        .card {{
            background: var(--card-bg);
            border-radius: 10px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            border: 1px solid var(--border);
            overflow: hidden;
        }}
        .card-header {{
            padding: 16px 20px;
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        .card-header h2 {{ font-size: 16px; font-weight: 600; color: var(--text); }}
        .card-header .badge {{
            font-size: 11px;
            padding: 2px 8px;
            border-radius: 12px;
            font-weight: 600;
        }}
        .badge-blue {{ background: var(--primary-light); color: var(--primary); }}
        .badge-green {{ background: #e6f4ea; color: var(--success); }}
        .badge-red {{ background: #fce8e6; color: var(--danger); }}
        .card-body {{ padding: 16px 20px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
        th {{
            background: var(--bg);
            padding: 10px 12px;
            text-align: left;
            font-weight: 600;
            color: var(--text-secondary);
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            border-bottom: 2px solid var(--border);
        }}
        td {{ padding: 10px 12px; border-bottom: 1px solid #f0f0f0; }}
        tr:last-child td {{ border-bottom: none; }}
        tr:hover td {{ background: #fafbfc; }}
        .total-row {{ font-weight: 700; background: var(--primary-light) !important; }}
        .total-row td {{ border-top: 2px solid var(--primary); color: var(--primary); }}
        .text-right {{ text-align: right; }}
        .text-center {{ text-align: center; }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }}
        .kpi {{
            background: var(--card-bg);
            padding: 20px;
            border-radius: 10px;
            border: 1px solid var(--border);
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }}
        .kpi-label {{
            font-size: 12px;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 4px;
        }}
        .kpi-value {{ font-size: 28px; font-weight: 700; color: var(--text); }}
        .kpi-sub {{ font-size: 12px; color: var(--text-secondary); margin-top: 2px; }}
        .kpi-positive {{ color: var(--success) !important; }}
        .kpi-negative {{ color: var(--danger) !important; }}
        .chart-container {{ position: relative; height: 280px; padding: 10px; }}
        .full-width {{ grid-column: 1 / -1; }}
        .comparison-bar {{ display: flex; align-items: center; gap: 8px; margin: 8px 0; }}
        .comparison-bar .bar {{ height: 24px; border-radius: 4px; transition: width 0.5s ease; }}
        .bar-teorico {{ background: #a8c7fa; }}
        .bar-real {{ background: var(--primary); }}
        .comparison-label {{ font-size: 12px; min-width: 100px; color: var(--text-secondary); }}
        .comparison-value {{ font-size: 13px; font-weight: 600; min-width: 80px; text-align: right; }}
        .footer {{ text-align: center; padding: 20px; color: var(--text-secondary); font-size: 12px; }}
    </style>
</head>
<body>
<div class="container">

    <div class="header">
        <h1>📦 {dados['titulo']}</h1>
        <p>Relatório de análise de fechamento de container</p>
    </div>

    <div class="kpi-grid">
        <div class="kpi">
            <div class="kpi-label">Total de Peças Cortadas</div>
            <div class="kpi-value">{fmt_numero(dados['peso_real']['itens'][0]['pecas'], 0)}</div>
            <div class="kpi-sub">unidades processadas</div>
        </div>
        <div class="kpi">
            <div class="kpi-label">KGs Coletados</div>
            <div class="kpi-value">{fmt_numero(dados['resumo']['kgs_coletados'], 2)}</div>
            <div class="kpi-sub">quilogramas totais</div>
        </div>
        <div class="kpi">
            <div class="kpi-label">KGs Consumidos</div>
            <div class="kpi-value">{fmt_numero(dados['resumo']['kgs_consumidos'], 2)}</div>
            <div class="kpi-sub">quilogramas utilizados</div>
        </div>
        <div class="kpi">
            <div class="kpi-label">Saldo Devedor</div>
            <div class="kpi-value {'kpi-negative' if dados['resumo']['saldo_devedor'] and dados['resumo']['saldo_devedor'] > 0 else 'kpi-positive'}">{fmt_numero(dados['resumo']['saldo_devedor'], 2)} kg</div>
            <div class="kpi-sub">diferença pendente</div>
        </div>
        <div class="kpi">
            <div class="kpi-label">Diferença Peso Unit.</div>
            <div class="kpi-value {'kpi-negative' if diferenca_peso > 0 else 'kpi-positive'}">{'+' if diferenca_peso > 0 else ''}{fmt_numero(diferenca_peso, 4)}</div>
            <div class="kpi-sub">Real ({fmt_numero(dados['peso_real']['peso_unitario'], 4)}) vs Teórico ({fmt_numero(dados['peso_teorico']['peso_unitario'], 4)})</div>
        </div>
    </div>

    <div class="grid">

        <div class="card">
            <div class="card-header">
                <h2>⚖️ Peso Teórico</h2>
                <span class="badge badge-blue">Unit: {fmt_numero(dados['peso_teorico']['peso_unitario'], 4)}</span>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr><th>Item</th><th class="text-right">Peças</th><th class="text-right">KGs</th><th class="text-right">Peso Unit.</th><th class="text-right">%</th></tr>
                    </thead>
                    <tbody>"""

    for item in dados["peso_teorico"]["itens"]:
        html += f"""
                        <tr>
                            <td>{item['nome']}</td>
                            <td class="text-right">{fmt_numero(item['pecas'], 0) if item['pecas'] else '-'}</td>
                            <td class="text-right">{fmt_numero(item['kgs'], 2)}</td>
                            <td class="text-right">{fmt_numero(item['peso_unit'], 4) if item['peso_unit'] else '-'}</td>
                            <td class="text-right">{fmt_pct(item['pct'])}</td>
                        </tr>"""

    html += f"""
                        <tr class="total-row">
                            <td>TOTAL</td><td class="text-right"></td>
                            <td class="text-right">{fmt_numero(dados['peso_teorico']['total_kgs'], 2)}</td>
                            <td class="text-right"></td>
                            <td class="text-right">{fmt_pct(dados['peso_teorico']['total_pct'])}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div class="card">
            <div class="card-header">
                <h2>📐 Peso Real</h2>
                <span class="badge badge-green">Unit: {fmt_numero(dados['peso_real']['peso_unitario'], 4)}</span>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr><th>Item</th><th class="text-right">Peças</th><th class="text-right">KGs</th><th class="text-right">Peso Unit.</th><th class="text-right">%</th></tr>
                    </thead>
                    <tbody>"""

    for item in dados["peso_real"]["itens"]:
        html += f"""
                        <tr>
                            <td>{item['nome']}</td>
                            <td class="text-right">{fmt_numero(item['pecas'], 0) if item['pecas'] else '-'}</td>
                            <td class="text-right">{fmt_numero(item['kgs'], 2)}</td>
                            <td class="text-right">{fmt_numero(item['peso_unit'], 4) if item['peso_unit'] else '-'}</td>
                            <td class="text-right">{fmt_pct(item['pct'])}</td>
                        </tr>"""

    html += f"""
                        <tr class="total-row">
                            <td>TOTAL</td><td class="text-right"></td>
                            <td class="text-right">{fmt_numero(dados['peso_real']['total_kgs'], 2)}</td>
                            <td class="text-right"></td>
                            <td class="text-right">{fmt_pct(dados['peso_real']['total_pct'])}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div class="card">
            <div class="card-header">
                <h2>📋 OPs Utilizadas</h2>
                <span class="badge badge-blue">{len(dados['ops'])} OPs</span>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr><th>OP</th><th class="text-right">KGs</th><th class="text-right">Peças</th></tr>
                    </thead>
                    <tbody>"""

    for op in dados["ops"]:
        html += f"""
                        <tr>
                            <td>{op['op']}</td>
                            <td class="text-right">{fmt_numero(op['kgs'], 1)}</td>
                            <td class="text-right">{fmt_numero(op['pecas'], 0)}</td>
                        </tr>"""

    html += f"""
                        <tr class="total-row">
                            <td>TOTAL</td>
                            <td class="text-right">{fmt_numero(dados['ops_total']['kgs'], 1)}</td>
                            <td class="text-right">{fmt_numero(dados['ops_total']['pecas'], 0)}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div class="card">
            <div class="card-header">
                <h2>📊 Resumo Total</h2>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr><th>Indicador</th><th class="text-right">Valor (KG)</th></tr>
                    </thead>
                    <tbody>
                        <tr><td>KGs Coletados</td><td class="text-right">{fmt_numero(dados['resumo']['kgs_coletados'], 2)}</td></tr>
                        <tr><td>KGs Consumidos</td><td class="text-right">{fmt_numero(dados['resumo']['kgs_consumidos'], 2)}</td></tr>
                        <tr class="total-row"><td>Saldo Devedor</td><td class="text-right">{fmt_numero(dados['resumo']['saldo_devedor'], 2)}</td></tr>
                    </tbody>
                </table>
                <div style="margin-top: 20px;">
                    <div class="comparison-bar">
                        <span class="comparison-label">Coletados</span>
                        <div class="bar bar-teorico" style="width: 100%;"></div>
                        <span class="comparison-value">{fmt_numero(dados['resumo']['kgs_coletados'], 0)} kg</span>
                    </div>
                    <div class="comparison-bar">
                        <span class="comparison-label">Consumidos</span>
                        <div class="bar bar-real" style="width: {(dados['resumo']['kgs_consumidos'] / dados['resumo']['kgs_coletados'] * 100) if dados['resumo']['kgs_coletados'] else 0:.1f}%;"></div>
                        <span class="comparison-value">{fmt_numero(dados['resumo']['kgs_consumidos'], 0)} kg</span>
                    </div>
                </div>
            </div>
        </div>

        <div class="card full-width">
            <div class="card-header">
                <h2>🔬 Conferência de Gramatura</h2>
                <span class="badge badge-blue">{len(dados['gramatura'])} mantas</span>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr><th>Manta</th><th class="text-right">Largura (m)</th><th class="text-right">Comprimento (m)</th><th class="text-right">Peso (kg)</th></tr>
                    </thead>
                    <tbody>"""

    peso_teorico = dados['peso_teorico']['peso_unitario']
    for m in dados["gramatura"]:
        peso_val = m['peso']
        cor = ""
        if peso_val and peso_teorico:
            desvio = abs(peso_val - peso_teorico)
            if desvio >= 0.006:
                cor = ' style="color: var(--danger); font-weight: 600;"'
            elif desvio <= 0.005:
                cor = ' style="color: var(--success); font-weight: 600;"'

        html += f"""
                        <tr>
                            <td>{m['manta']}</td>
                            <td class="text-right">{fmt_numero(m['largura'], 2)}</td>
                            <td class="text-right">{fmt_numero(m['comprimento'], 2)}</td>
                            <td class="text-right"{cor}>{fmt_numero(m['peso'], 3)}</td>
                        </tr>"""

    html += f"""
                        <tr class="total-row">
                            <td>MÉDIA</td>
                            <td class="text-right">{fmt_numero(dados['gramatura_media']['largura'], 4)}</td>
                            <td class="text-right">{fmt_numero(dados['gramatura_media']['comprimento'], 2)}</td>
                            <td class="text-right">{fmt_numero(dados['gramatura_media']['peso'], 3)}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <div class="grid">
        <div class="card">
            <div class="card-header"><h2>📈 Distribuição de KGs por OP</h2></div>
            <div class="card-body"><div class="chart-container"><canvas id="chartOps"></canvas></div></div>
        </div>
        <div class="card">
            <div class="card-header"><h2>📉 Peso por Manta (Gramatura)</h2></div>
            <div class="card-body"><div class="chart-container"><canvas id="chartGramatura"></canvas></div></div>
        </div>
        <div class="card">
            <div class="card-header"><h2>⚖️ Peso Teórico vs Real (% composição)</h2></div>
            <div class="card-body"><div class="chart-container"><canvas id="chartComparacao"></canvas></div></div>
        </div>
        <div class="card">
            <div class="card-header"><h2>📊 Peças por OP</h2></div>
            <div class="card-body"><div class="chart-container"><canvas id="chartPecasOp"></canvas></div></div>
        </div>
    </div>

    <div class="footer">
        Análise de Fechamento de Containers <br> Por <br> Guedes e Erick
    </div>

</div>

<script>
    const cores = ['#1a73e8', '#34a853', '#fbbc04', '#ea4335', '#673ab7', '#ff6d00'];
    const coresAlpha = cores.map(c => c + '33');

    new Chart(document.getElementById('chartOps'), {{
        type: 'bar',
        data: {{
            labels: {[op['op'] for op in dados['ops']]},
            datasets: [{{
                label: 'KGs',
                data: {[op['kgs'] for op in dados['ops']]},
                backgroundColor: cores.slice(0, {len(dados['ops'])}),
                borderRadius: 6,
                barThickness: 40
            }}]
        }},
        options: {{
            responsive: true, maintainAspectRatio: false,
            plugins: {{ legend: {{ display: false }}, tooltip: {{ callbacks: {{ label: ctx => ctx.parsed.y.toLocaleString('pt-BR') + ' kg' }} }} }},
            scales: {{ y: {{ beginAtZero: true, ticks: {{ callback: v => v.toLocaleString('pt-BR') + ' kg' }} }} }}
        }}
    }});

    new Chart(document.getElementById('chartGramatura'), {{
        type: 'line',
        data: {{
            labels: {[m['manta'] for m in dados['gramatura']]},
            datasets: [{{
                label: 'Peso (kg)',
                data: {[m['peso'] for m in dados['gramatura']]},
                borderColor: '#1a73e8', backgroundColor: '#1a73e833',
                fill: true, tension: 0.3, pointRadius: 5, pointBackgroundColor: '#1a73e8'
            }},{{
                label: 'Média',
                data: Array({len(dados['gramatura'])}).fill({dados['gramatura_media']['peso']}),
                borderColor: '#ea4335', borderDash: [5, 5], pointRadius: 0, fill: false
            }},{{
                label: 'Peso Teórico ({fmt_numero(dados["peso_teorico"]["peso_unitario"], 4)} kg)',
                data: Array({len(dados['gramatura'])}).fill({dados['peso_teorico']['peso_unitario']}),
                borderColor: '#0d904f', borderDash: [10, 4], borderWidth: 2, pointRadius: 0, fill: false
            }}]
        }},
        options: {{
            responsive: true, maintainAspectRatio: false,
            plugins: {{ tooltip: {{ callbacks: {{ label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y.toFixed(4) + ' kg' }} }} }},
            scales: {{ y: {{ min: 0.60, max: 0.73 }} }}
        }}
    }});

    new Chart(document.getElementById('chartComparacao'), {{
        type: 'bar',
        data: {{
            labels: {[item['nome'] for item in dados['peso_teorico']['itens']]},
            datasets: [
                {{ label: 'Teórico (%)', data: {[item['pct'] if item['pct'] else 0 for item in dados['peso_teorico']['itens']]}, backgroundColor: '#a8c7fa', borderRadius: 4 }},
                {{ label: 'Real (%)', data: {[item['pct'] if item['pct'] else 0 for item in dados['peso_real']['itens']]}, backgroundColor: '#1a73e8', borderRadius: 4 }}
            ]
        }},
        options: {{
            responsive: true, maintainAspectRatio: false,
            plugins: {{ tooltip: {{ callbacks: {{ label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y.toFixed(2) + '%' }} }} }},
            scales: {{ y: {{ beginAtZero: true, ticks: {{ callback: v => v + '%' }} }} }}
        }}
    }});

    new Chart(document.getElementById('chartPecasOp'), {{
        type: 'doughnut',
        data: {{
            labels: {[op['op'] for op in dados['ops']]},
            datasets: [{{
                data: {[op['pecas'] for op in dados['ops']]},
                backgroundColor: cores, borderWidth: 2, borderColor: '#fff'
            }}]
        }},
        options: {{
            responsive: true, maintainAspectRatio: false,
            plugins: {{
                tooltip: {{ callbacks: {{
                    label: ctx => {{
                        const total = ctx.dataset.data.reduce((a,b) => a+b, 0);
                        const pct = (ctx.parsed / total * 100).toFixed(1);
                        return ctx.label + ': ' + ctx.parsed.toLocaleString('pt-BR') + ' peças (' + pct + '%)';
                    }}
                }} }}
            }}
        }}
    }});
</script>
</body>
</html>"""

    return html


# ═══════════════════════════════════════════════════════════════════════════════
#  SALVAR CSV
# ═══════════════════════════════════════════════════════════════════════════════

def gerar_csv(dados: dict) -> str:
    """Gera conteúdo CSV no formato original a partir do dicionário de dados."""
    def n(v, dec=2):
        if v is None:
            return "-"
        return f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def n4(v):
        return n(v, 4)

    def pct(v):
        if v is None:
            return "-"
        return f"{v:.2f}%".replace(".", ",")

    def ni(v):
        """Int com ponto de milhar."""
        if v is None:
            return "-"
        return f"{int(v):,}".replace(",", ".")

    linhas = []

    # Linha 0 - Título
    linhas.append([dados["titulo"]] + [""] * 14)

    # Linha 1 - Header PESO TEORICO
    linhas.append(["PESO TEORICO", "", "", "", n4(dados["peso_teorico"]["peso_unitario"]),
                    "", "OPS UTILIZADAS", "KGS UTILIZADOS", "PEÇAS", "", "",
                    "CONFERENCIA DE GRAMATURA", "", "", ""])

    # Linha 2 - Sub-header
    ops = dados["ops"]
    gram = dados["gramatura"]
    op0 = ops[0] if len(ops) > 0 else {"op": "", "kgs": None, "pecas": None}
    linhas.append(["", "PEÇAS", "KGS", "PESO UNIT", "%", "",
                    op0["op"], n(op0["kgs"], 1), ni(op0["pecas"]), "", "",
                    "MANTA", "LARGURA", "COMPRIMENTO", "PESO"])

    # Linhas 3-5: itens peso teórico + OPs + gramatura
    for idx in range(3):
        item_t = dados["peso_teorico"]["itens"][idx] if idx < len(dados["peso_teorico"]["itens"]) else {"nome": "", "pecas": None, "kgs": None, "peso_unit": None, "pct": None}
        op_i = ops[idx + 1] if idx + 1 < len(ops) else {"op": "", "kgs": None, "pecas": None}
        gram_i = gram[idx] if idx < len(gram) else {"manta": "", "largura": None, "comprimento": None, "peso": None}

        pecas_str = ni(item_t["pecas"]) if item_t["pecas"] else " -   "
        kgs_str = n(item_t["kgs"], 4) if item_t["kgs"] else ""
        pu_str = n(item_t["peso_unit"], 6) if item_t["peso_unit"] else " -   "
        pct_str = pct(item_t["pct"]) if item_t["pct"] else ""

        linhas.append([
            item_t["nome"], pecas_str, kgs_str, pu_str, pct_str, "",
            op_i["op"], n(op_i["kgs"], 1) if op_i["kgs"] else "",
            ni(op_i["pecas"]) if op_i["pecas"] else "", "", "",
            gram_i["manta"], n(gram_i["largura"], 2) if gram_i["largura"] else "",
            n(gram_i["comprimento"], 2) if gram_i["comprimento"] else "",
            n(gram_i["peso"], 3) if gram_i["peso"] else ""
        ])

    # Linha 6 - extra ops
    remaining_ops_count = max(0, len(ops) - 4)
    linhas.append(["", "", "", "", "", "",
                    "TOTAL", n(dados["ops_total"]["kgs"], 1),
                    ni(dados["ops_total"]["pecas"]),
                    n(dados["ops_total"]["peso_medio"], 9) if dados["ops_total"]["peso_medio"] else "", "",
                    gram[3]["manta"] if len(gram) > 3 else "",
                    n(gram[3]["largura"], 2) if len(gram) > 3 and gram[3]["largura"] else "",
                    n(gram[3]["comprimento"], 2) if len(gram) > 3 and gram[3]["comprimento"] else "",
                    n(gram[3]["peso"], 3) if len(gram) > 3 and gram[3]["peso"] else ""])

    # Linha 7 - Total peso teórico
    linhas.append(["", "", n(dados["peso_teorico"]["total_kgs"], 2), "", pct(dados["peso_teorico"]["total_pct"]),
                    "", "", "", "", "", "",
                    gram[4]["manta"] if len(gram) > 4 else "",
                    n(gram[4]["largura"], 2) if len(gram) > 4 and gram[4]["largura"] else "",
                    n(gram[4]["comprimento"], 2) if len(gram) > 4 and gram[4]["comprimento"] else "",
                    n(gram[4]["peso"], 3) if len(gram) > 4 and gram[4]["peso"] else ""])

    # Linha 8 - vazio + gramatura
    linhas.append(["", "", "", "", "", "", "", "", "", "", "",
                    gram[5]["manta"] if len(gram) > 5 else "",
                    n(gram[5]["largura"], 2) if len(gram) > 5 and gram[5]["largura"] else "",
                    n(gram[5]["comprimento"], 2) if len(gram) > 5 and gram[5]["comprimento"] else "",
                    n(gram[5]["peso"], 3) if len(gram) > 5 and gram[5]["peso"] else ""])

    # Linha 9 - PESO REAL header
    linhas.append(["PESO REAL", "", "", "", n4(dados["peso_real"]["peso_unitario"]),
                    "", "RESUMO TOTAL", "", "", "", "",
                    gram[6]["manta"] if len(gram) > 6 else "",
                    n(gram[6]["largura"], 2) if len(gram) > 6 and gram[6]["largura"] else "",
                    n(gram[6]["comprimento"], 2) if len(gram) > 6 and gram[6]["comprimento"] else "",
                    n(gram[6]["peso"], 3) if len(gram) > 6 and gram[6]["peso"] else ""])

    # Linha 10 - sub-header peso real + KGs coletados
    linhas.append(["", "PEÇAS", "KGS", "PESO UNIT", "%", "",
                    "KGS COLETADOS", n(dados["resumo"]["kgs_coletados"], 2), "", "", "",
                    gram[7]["manta"] if len(gram) > 7 else "",
                    n(gram[7]["largura"], 2) if len(gram) > 7 and gram[7]["largura"] else "",
                    n(gram[7]["comprimento"], 2) if len(gram) > 7 and gram[7]["comprimento"] else "",
                    n(gram[7]["peso"], 3) if len(gram) > 7 and gram[7]["peso"] else ""])

    # Linhas 11-13: itens peso real
    resumo_labels = ["KGS CONSUMIDOS", "", "SALDO DEVEDOR (KGS)"]
    resumo_vals = [dados["resumo"]["kgs_consumidos"], None, dados["resumo"]["saldo_devedor"]]
    resumo_extra = ["", "", "NF Perda"]
    for idx in range(3):
        item_r = dados["peso_real"]["itens"][idx] if idx < len(dados["peso_real"]["itens"]) else {"nome": "", "pecas": None, "kgs": None, "peso_unit": None, "pct": None}
        gram_idx = 8 + idx

        pecas_str = ni(item_r["pecas"]) if item_r["pecas"] else " -   "
        kgs_str = n(item_r["kgs"], 4) if item_r["kgs"] else ""
        pu_str = n(item_r["peso_unit"], 6) if item_r["peso_unit"] else " -   "
        pct_str = pct(item_r["pct"]) if item_r["pct"] else ""

        linhas.append([
            item_r["nome"], pecas_str, kgs_str, pu_str, pct_str, "",
            resumo_labels[idx],
            n(resumo_vals[idx], 2) if resumo_vals[idx] is not None else "",
            resumo_extra[idx], "", "",
            gram[gram_idx]["manta"] if gram_idx < len(gram) else "",
            n(gram[gram_idx]["largura"], 2) if gram_idx < len(gram) and gram[gram_idx]["largura"] else "",
            n(gram[gram_idx]["comprimento"], 2) if gram_idx < len(gram) and gram[gram_idx]["comprimento"] else "",
            n(gram[gram_idx]["peso"], 3) if gram_idx < len(gram) and gram[gram_idx]["peso"] else ""
        ])

    # Linha 14 - gramatura extra
    linhas.append(["", "", "", "", "", "", "", "", "", "", "",
                    gram[11]["manta"] if len(gram) > 11 else "",
                    n(gram[11]["largura"], 2) if len(gram) > 11 and gram[11]["largura"] else "",
                    n(gram[11]["comprimento"], 2) if len(gram) > 11 and gram[11]["comprimento"] else "",
                    n(gram[11]["peso"], 3) if len(gram) > 11 and gram[11]["peso"] else ""])

    # Linha 15 - Total peso real + gramatura
    linhas.append(["", "", n(dados["peso_real"]["total_kgs"], 2), "", pct(dados["peso_real"]["total_pct"]),
                    "", "", "", "", "", "",
                    gram[12]["manta"] if len(gram) > 12 else "",
                    n(gram[12]["largura"], 2) if len(gram) > 12 and gram[12]["largura"] else "",
                    n(gram[12]["comprimento"], 2) if len(gram) > 12 and gram[12]["comprimento"] else "",
                    n(gram[12]["peso"], 3) if len(gram) > 12 and gram[12]["peso"] else ""])

    # Linha 16 - vazio
    linhas.append([""] * 15)

    # Linha 17 - Média gramatura
    linhas.append([""] * 11 + [
        "MEDIA",
        n4(dados["gramatura_media"]["largura"]),
        n(dados["gramatura_media"]["comprimento"], 2),
        n(dados["gramatura_media"]["peso"], 3)
    ])

    # Converter para CSV string
    resultado = ""
    for linha in linhas:
        resultado += ";".join(str(c) for c in linha) + "\n"
    return resultado


# ═══════════════════════════════════════════════════════════════════════════════
#  CARREGAR CSV EXISTENTE
# ═══════════════════════════════════════════════════════════════════════════════

def ler_csv(caminho: Path) -> list[list[str]]:
    """Lê o CSV com delimitador ';'."""
    linhas = []
    with open(caminho, "r", encoding="utf-8") as f:
        leitor = csv.reader(f, delimiter=";")
        for linha in leitor:
            linhas.append(linha)
    return linhas


def extrair_dados_csv(linhas: list[list[str]]) -> dict:
    """Extrai dados do CSV para dicionário."""
    dados = {}
    dados["titulo"] = linhas[0][0].strip()

    peso_teorico_unit = parse_numero(linhas[1][4])
    dados["peso_teorico"] = {"peso_unitario": peso_teorico_unit, "itens": []}
    for i in range(3, 6):
        nome = linhas[i][0].strip()
        pecas = parse_numero(linhas[i][1])
        kgs = parse_numero(linhas[i][2])
        peso_unit = parse_numero(linhas[i][3])
        pct = parse_numero(linhas[i][4])
        dados["peso_teorico"]["itens"].append({
            "nome": nome, "pecas": pecas, "kgs": kgs,
            "peso_unit": peso_unit, "pct": pct
        })
    dados["peso_teorico"]["total_kgs"] = parse_numero(linhas[7][2])
    dados["peso_teorico"]["total_pct"] = parse_numero(linhas[7][4])

    peso_real_unit = parse_numero(linhas[9][4])
    dados["peso_real"] = {"peso_unitario": peso_real_unit, "itens": []}
    for i in range(11, 14):
        nome = linhas[i][0].strip()
        pecas = parse_numero(linhas[i][1])
        kgs = parse_numero(linhas[i][2])
        peso_unit = parse_numero(linhas[i][3])
        pct = parse_numero(linhas[i][4])
        dados["peso_real"]["itens"].append({
            "nome": nome, "pecas": pecas, "kgs": kgs,
            "peso_unit": peso_unit, "pct": pct
        })
    dados["peso_real"]["total_kgs"] = parse_numero(linhas[15][2])
    dados["peso_real"]["total_pct"] = parse_numero(linhas[15][4])

    dados["ops"] = []
    for i in range(2, 7):
        op = linhas[i][6].strip() if len(linhas[i]) > 6 and linhas[i][6].strip() else None
        kgs = parse_numero(linhas[i][7]) if len(linhas[i]) > 7 else None
        pecas = parse_numero(linhas[i][8]) if len(linhas[i]) > 8 else None
        if op and op != "TOTAL":
            dados["ops"].append({"op": op, "kgs": kgs, "pecas": pecas})

    dados["ops_total"] = {
        "kgs": parse_numero(linhas[6][7]) if len(linhas[6]) > 7 else None,
        "pecas": parse_numero(linhas[6][8]) if len(linhas[6]) > 8 else None,
        "peso_medio": parse_numero(linhas[6][9]) if len(linhas[6]) > 9 else None
    }

    dados["resumo"] = {
        "kgs_coletados": parse_numero(linhas[10][7]) if len(linhas[10]) > 7 else None,
        "kgs_consumidos": parse_numero(linhas[11][7]) if len(linhas[11]) > 7 else None,
        "saldo_devedor": parse_numero(linhas[13][7]) if len(linhas[13]) > 7 else None,
    }

    dados["gramatura"] = []
    for i in range(3, 16):
        manta = linhas[i][11].strip() if len(linhas[i]) > 11 and linhas[i][11].strip() else None
        if manta and manta.startswith("#"):
            largura = parse_numero(linhas[i][12]) if len(linhas[i]) > 12 else None
            comprimento = parse_numero(linhas[i][13]) if len(linhas[i]) > 13 else None
            peso = parse_numero(linhas[i][14]) if len(linhas[i]) > 14 else None
            dados["gramatura"].append({
                "manta": manta, "largura": largura,
                "comprimento": comprimento, "peso": peso
            })

    dados["gramatura_media"] = {
        "largura": parse_numero(linhas[17][12]) if len(linhas) > 17 and len(linhas[17]) > 12 else None,
        "comprimento": parse_numero(linhas[17][13]) if len(linhas) > 17 and len(linhas[17]) > 13 else None,
        "peso": parse_numero(linhas[17][14]) if len(linhas) > 17 and len(linhas[17]) > 14 else None,
    }

    return dados


# ═══════════════════════════════════════════════════════════════════════════════
#  INTERFACE GRÁFICA (TKINTER)
# ═══════════════════════════════════════════════════════════════════════════════

class AppFechamentoContainer:
    """Aplicação principal com formulário para preenchimento de dados."""

    def __init__(self, root):
        self.root = root
        self.root.title("Fechamento de Containers - Gerador de Dashboard")
        self.root.geometry("1100x750")
        self.root.configure(bg="#f0f2f5")

        # Estilo
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), background="#f0f2f5", foreground="#1a73e8")
        self.style.configure("Section.TLabelframe.Label", font=("Segoe UI", 11, "bold"), foreground="#1a73e8")
        self.style.configure("TLabel", font=("Segoe UI", 10), background="#f0f2f5")
        self.style.configure("TEntry", font=("Segoe UI", 10))
        self.style.configure("Action.TButton", font=("Segoe UI", 11, "bold"), padding=10)
        self.style.configure("Small.TButton", font=("Segoe UI", 9), padding=4)
        self.style.configure("TLabelframe", background="#f0f2f5")
        self.style.configure("TFrame", background="#f0f2f5")

        # ── Variáveis dinâmicas ──
        self.op_rows = []       # Lista de widgets de cada OP
        self.manta_rows = []    # Lista de widgets de cada manta

        self._build_ui()

    def _build_ui(self):
        """Constrói toda a interface."""

        # ── Menu de ações no topo ──
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=15, pady=(10, 0))

        ttk.Label(top_frame, text="Fechamento de Containers", style="Title.TLabel").pack(side="left")

        btn_frame = ttk.Frame(top_frame)
        btn_frame.pack(side="right")

        ttk.Button(btn_frame, text="📂 Carregar CSV", command=self.carregar_csv, style="Small.TButton").pack(side="left", padx=3)
        ttk.Button(btn_frame, text="💾 Salvar CSV", command=self.salvar_csv, style="Small.TButton").pack(side="left", padx=3)
        ttk.Button(btn_frame, text="🗑️ Limpar Tudo", command=self.limpar_tudo, style="Small.TButton").pack(side="left", padx=3)

        # ── Área com scroll ──
        container = ttk.Frame(self.root)
        container.pack(fill="both", expand=True, padx=15, pady=10)

        canvas = tk.Canvas(container, bg="#f0f2f5", highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)

        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Scroll com mouse
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        parent = self.scroll_frame

        # ══════════ TÍTULO ══════════
        sec = ttk.LabelFrame(parent, text="Informações Gerais", style="Section.TLabelframe", padding=10)
        sec.pack(fill="x", pady=5)

        ttk.Label(sec, text="Título do Container:").grid(row=0, column=0, sticky="w", padx=5)
        self.entry_titulo = ttk.Entry(sec, width=60)
        self.entry_titulo.grid(row=0, column=1, columnspan=3, sticky="w", padx=5)

        # ══════════ PESO TEÓRICO ══════════
        sec_pt = ttk.LabelFrame(parent, text="Peso Teórico", style="Section.TLabelframe", padding=10)
        sec_pt.pack(fill="x", pady=5)

        ttk.Label(sec_pt, text="Peso Unit. Teórico:").grid(row=0, column=0, sticky="w", padx=5)
        self.entry_pt_unit = ttk.Entry(sec_pt, width=15)
        self.entry_pt_unit.grid(row=0, column=1, sticky="w", padx=5)

        # Headers
        for j, h in enumerate(["Item", "Peças", "KGs", "Peso Unit."]):
            ttk.Label(sec_pt, text=h, font=("Segoe UI", 9, "bold")).grid(row=1, column=j, padx=5, pady=(8, 2))

        self.pt_items = []
        nomes_default = ["PEÇAS CORTADAS", "BABY RETIRADO", "RETALHOS"]
        for i in range(3):
            row = {}
            r = i + 2
            row["nome"] = ttk.Entry(sec_pt, width=20)
            row["nome"].grid(row=r, column=0, padx=3, pady=2)
            row["nome"].insert(0, nomes_default[i])
            row["pecas"] = ttk.Entry(sec_pt, width=12)
            row["pecas"].grid(row=r, column=1, padx=3, pady=2)
            row["kgs"] = ttk.Entry(sec_pt, width=15)
            row["kgs"].grid(row=r, column=2, padx=3, pady=2)
            row["peso_unit"] = ttk.Entry(sec_pt, width=12)
            row["peso_unit"].grid(row=r, column=3, padx=3, pady=2)
            self.pt_items.append(row)

        # ══════════ PESO REAL ══════════
        sec_pr = ttk.LabelFrame(parent, text="Peso Real", style="Section.TLabelframe", padding=10)
        sec_pr.pack(fill="x", pady=5)

        ttk.Label(sec_pr, text="Peso Unit. Real:").grid(row=0, column=0, sticky="w", padx=5)
        self.entry_pr_unit = ttk.Entry(sec_pr, width=15)
        self.entry_pr_unit.grid(row=0, column=1, sticky="w", padx=5)

        for j, h in enumerate(["Item", "Peças", "KGs", "Peso Unit."]):
            ttk.Label(sec_pr, text=h, font=("Segoe UI", 9, "bold")).grid(row=1, column=j, padx=5, pady=(8, 2))

        self.pr_items = []
        for i in range(3):
            row = {}
            r = i + 2
            row["nome"] = ttk.Entry(sec_pr, width=20)
            row["nome"].grid(row=r, column=0, padx=3, pady=2)
            row["nome"].insert(0, nomes_default[i])
            row["pecas"] = ttk.Entry(sec_pr, width=12)
            row["pecas"].grid(row=r, column=1, padx=3, pady=2)
            row["kgs"] = ttk.Entry(sec_pr, width=15)
            row["kgs"].grid(row=r, column=2, padx=3, pady=2)
            row["peso_unit"] = ttk.Entry(sec_pr, width=12)
            row["peso_unit"].grid(row=r, column=3, padx=3, pady=2)
            self.pr_items.append(row)

        # ══════════ OPS UTILIZADAS ══════════
        sec_ops = ttk.LabelFrame(parent, text="OPs Utilizadas", style="Section.TLabelframe", padding=10)
        sec_ops.pack(fill="x", pady=5)
        self.sec_ops = sec_ops

        btn_ops_frame = ttk.Frame(sec_ops)
        btn_ops_frame.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 5))
        ttk.Button(btn_ops_frame, text="+ Adicionar OP", command=self.adicionar_op, style="Small.TButton").pack(side="left", padx=3)
        ttk.Button(btn_ops_frame, text="- Remover Última", command=self.remover_op, style="Small.TButton").pack(side="left", padx=3)

        for j, h in enumerate(["OP", "KGs", "Peças"]):
            ttk.Label(sec_ops, text=h, font=("Segoe UI", 9, "bold")).grid(row=1, column=j, padx=5, pady=(2, 2))

        self.ops_start_row = 2

        # ══════════ GRAMATURA ══════════
        sec_gram = ttk.LabelFrame(parent, text="Conferência de Gramatura", style="Section.TLabelframe", padding=10)
        sec_gram.pack(fill="x", pady=5)
        self.sec_gram = sec_gram

        btn_gram_frame = ttk.Frame(sec_gram)
        btn_gram_frame.grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 5))
        ttk.Button(btn_gram_frame, text="+ Adicionar Manta", command=self.adicionar_manta, style="Small.TButton").pack(side="left", padx=3)
        ttk.Button(btn_gram_frame, text="- Remover Última", command=self.remover_manta, style="Small.TButton").pack(side="left", padx=3)

        for j, h in enumerate(["Manta", "Largura (m)", "Comprimento (m)", "Peso (kg)"]):
            ttk.Label(sec_gram, text=h, font=("Segoe UI", 9, "bold")).grid(row=1, column=j, padx=5, pady=(2, 2))

        self.gram_start_row = 2



        # ══════════ BOTÃO GERAR ══════════
        btn_gerar_frame = ttk.Frame(parent)
        btn_gerar_frame.pack(fill="x", pady=15)

        gerar_btn = tk.Button(
            btn_gerar_frame, text="🚀  GERAR DASHBOARD  🚀",
            command=self.gerar_dashboard,
            font=("Segoe UI", 14, "bold"),
            bg="#1a73e8", fg="white",
            activebackground="#1557b0", activeforeground="white",
            relief="flat", padx=30, pady=12, cursor="hand2"
        )
        gerar_btn.pack(pady=5)

        # Adicionar OPs default (4)
        for _ in range(4):
            self.adicionar_op()

        # Adicionar Mantas default (13)
        for _ in range(13):
            self.adicionar_manta()

    # ── Métodos: OPs dinâmicas ──

    def adicionar_op(self):
        """Adiciona uma linha de OP ao formulário."""
        idx = len(self.op_rows)
        r = self.ops_start_row + idx
        row = {}
        row["op"] = ttk.Entry(self.sec_ops, width=15)
        row["op"].grid(row=r, column=0, padx=3, pady=2)
        row["kgs"] = ttk.Entry(self.sec_ops, width=15)
        row["kgs"].grid(row=r, column=1, padx=3, pady=2)
        row["pecas"] = ttk.Entry(self.sec_ops, width=15)
        row["pecas"].grid(row=r, column=2, padx=3, pady=2)
        self.op_rows.append(row)

    def remover_op(self):
        """Remove a última OP."""
        if self.op_rows:
            row = self.op_rows.pop()
            for w in row.values():
                w.destroy()

    # ── Métodos: Mantas dinâmicas ──

    def adicionar_manta(self):
        """Adiciona uma linha de manta ao formulário."""
        idx = len(self.manta_rows)
        r = self.gram_start_row + idx
        row = {}
        row["manta"] = ttk.Entry(self.sec_gram, width=10)
        row["manta"].grid(row=r, column=0, padx=3, pady=2)
        row["manta"].insert(0, f"#{idx + 1}")
        row["largura"] = ttk.Entry(self.sec_gram, width=12)
        row["largura"].grid(row=r, column=1, padx=3, pady=2)
        row["comprimento"] = ttk.Entry(self.sec_gram, width=12)
        row["comprimento"].grid(row=r, column=2, padx=3, pady=2)
        row["peso"] = ttk.Entry(self.sec_gram, width=12)
        row["peso"].grid(row=r, column=3, padx=3, pady=2)
        self.manta_rows.append(row)

    def remover_manta(self):
        """Remove a última manta."""
        if self.manta_rows:
            row = self.manta_rows.pop()
            for w in row.values():
                w.destroy()



    # ── Coletar dados do formulário ──

    def coletar_dados(self) -> dict:
        """Coleta todos os dados dos campos do formulário."""
        dados = {}
        dados["titulo"] = self.entry_titulo.get().strip() or "FECHAMENTO CONTAINER"

        # Peso Teórico
        dados["peso_teorico"] = {
            "peso_unitario": safe_float(self.entry_pt_unit, 0),
            "itens": [],
            "total_kgs": None,  # calculado automaticamente
            "total_pct": None,
        }
        for row in self.pt_items:
            dados["peso_teorico"]["itens"].append({
                "nome": row["nome"].get().strip(),
                "pecas": safe_float(row["pecas"]),
                "kgs": safe_float(row["kgs"]),
                "peso_unit": safe_float(row["peso_unit"]),
                "pct": None,
            })

        # Peso Real
        dados["peso_real"] = {
            "peso_unitario": safe_float(self.entry_pr_unit, 0),
            "itens": [],
            "total_kgs": None,  # calculado automaticamente
            "total_pct": None,
        }
        for row in self.pr_items:
            dados["peso_real"]["itens"].append({
                "nome": row["nome"].get().strip(),
                "pecas": safe_float(row["pecas"]),
                "kgs": safe_float(row["kgs"]),
                "peso_unit": safe_float(row["peso_unit"]),
                "pct": None,
            })

        # OPs
        dados["ops"] = []
        total_kgs_ops = 0
        total_pecas_ops = 0
        for row in self.op_rows:
            op_nome = row["op"].get().strip()
            kgs = safe_float(row["kgs"])
            pecas = safe_float(row["pecas"])
            if op_nome:
                dados["ops"].append({"op": op_nome, "kgs": kgs, "pecas": pecas})
                if kgs:
                    total_kgs_ops += kgs
                if pecas:
                    total_pecas_ops += pecas

        peso_medio = total_kgs_ops / total_pecas_ops if total_pecas_ops else None
        dados["ops_total"] = {
            "kgs": total_kgs_ops,
            "pecas": total_pecas_ops,
            "peso_medio": peso_medio
        }

        # Resumo (calculado automaticamente)
        dados["resumo"] = {
            "kgs_coletados": None,
            "kgs_consumidos": None,
            "saldo_devedor": None,
        }

        # Gramatura
        dados["gramatura"] = []
        for row in self.manta_rows:
            manta = row["manta"].get().strip()
            if manta:
                dados["gramatura"].append({
                    "manta": manta,
                    "largura": safe_float(row["largura"]),
                    "comprimento": safe_float(row["comprimento"]),
                    "peso": safe_float(row["peso"]),
                })

        # Média gramatura (calculada automaticamente)
        dados["gramatura_media"] = {
            "largura": None,
            "comprimento": None,
            "peso": None,
        }

        return dados

    # ── Preencher formulário com dados ──

    def preencher_formulario(self, dados: dict):
        """Preenche o formulário a partir de um dicionário de dados."""
        self.limpar_tudo()

        def set_entry(entry, valor, fmt="{}"):
            entry.delete(0, "end")
            if valor is not None:
                texto = fmt.format(valor).replace(".", ",")
                entry.insert(0, texto)

        set_entry(self.entry_titulo, dados.get("titulo", ""))
        if dados["titulo"]:
            self.entry_titulo.delete(0, "end")
            self.entry_titulo.insert(0, dados["titulo"])

        # Peso Teórico
        set_entry(self.entry_pt_unit, dados["peso_teorico"]["peso_unitario"], "{:.4f}")
        for i, item in enumerate(dados["peso_teorico"]["itens"]):
            if i < len(self.pt_items):
                self.pt_items[i]["nome"].delete(0, "end")
                self.pt_items[i]["nome"].insert(0, item["nome"])
                set_entry(self.pt_items[i]["pecas"], item["pecas"], "{:.0f}")
                set_entry(self.pt_items[i]["kgs"], item["kgs"], "{:.4f}")
                set_entry(self.pt_items[i]["peso_unit"], item["peso_unit"], "{:.6f}")

        # Peso Real
        set_entry(self.entry_pr_unit, dados["peso_real"]["peso_unitario"], "{:.4f}")
        for i, item in enumerate(dados["peso_real"]["itens"]):
            if i < len(self.pr_items):
                self.pr_items[i]["nome"].delete(0, "end")
                self.pr_items[i]["nome"].insert(0, item["nome"])
                set_entry(self.pr_items[i]["pecas"], item["pecas"], "{:.0f}")
                set_entry(self.pr_items[i]["kgs"], item["kgs"], "{:.4f}")
                set_entry(self.pr_items[i]["peso_unit"], item["peso_unit"], "{:.6f}")

        # OPs
        while self.op_rows:
            self.remover_op()
        for op in dados["ops"]:
            self.adicionar_op()
            idx = len(self.op_rows) - 1
            self.op_rows[idx]["op"].insert(0, op["op"])
            set_entry(self.op_rows[idx]["kgs"], op["kgs"], "{:.1f}")
            set_entry(self.op_rows[idx]["pecas"], op["pecas"], "{:.0f}")

        # Gramatura
        while self.manta_rows:
            self.remover_manta()
        for m in dados["gramatura"]:
            self.adicionar_manta()
            idx = len(self.manta_rows) - 1
            self.manta_rows[idx]["manta"].delete(0, "end")
            self.manta_rows[idx]["manta"].insert(0, m["manta"])
            set_entry(self.manta_rows[idx]["largura"], m["largura"], "{:.2f}")
            set_entry(self.manta_rows[idx]["comprimento"], m["comprimento"], "{:.2f}")
            set_entry(self.manta_rows[idx]["peso"], m["peso"], "{:.3f}")



    # ── Ações ──

    def limpar_tudo(self):
        """Limpa todos os campos do formulário."""
        self.entry_titulo.delete(0, "end")
        self.entry_pt_unit.delete(0, "end")
        self.entry_pr_unit.delete(0, "end")


        for row in self.pt_items:
            for key in ["pecas", "kgs", "peso_unit"]:
                row[key].delete(0, "end")

        for row in self.pr_items:
            for key in ["pecas", "kgs", "peso_unit"]:
                row[key].delete(0, "end")

    def carregar_csv(self):
        """Abre diálogo para selecionar CSV e preenche o formulário."""
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo CSV",
            initialdir=str(PASTA_BASE),
            filetypes=[("CSV", "*.csv"), ("Todos", "*.*")]
        )
        if not caminho:
            return

        try:
            linhas = ler_csv(Path(caminho))
            dados = extrair_dados_csv(linhas)
            self.preencher_formulario(dados)
            messagebox.showinfo("Sucesso", f"Dados carregados de:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro ao carregar", f"Não foi possível ler o CSV:\n{e}")

    def salvar_csv(self):
        """Salva os dados atuais em um arquivo CSV."""
        caminho = filedialog.asksaveasfilename(
            title="Salvar CSV",
            initialdir=str(PASTA_BASE),
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Todos", "*.*")]
        )
        if not caminho:
            return

        try:
            dados = self.coletar_dados()
            conteudo = gerar_csv(dados)
            with open(caminho, "w", encoding="utf-8") as f:
                f.write(conteudo)
            messagebox.showinfo("Sucesso", f"CSV salvo em:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", f"Não foi possível salvar o CSV:\n{e}")

    @staticmethod
    def _calcular_media_gramatura(dados):
        """Calcula a média de gramatura a partir dos dados das mantas."""
        larguras, comprimentos, pesos = [], [], []
        for m in dados["gramatura"]:
            if m.get("largura") is not None:
                larguras.append(m["largura"])
            if m.get("comprimento") is not None:
                comprimentos.append(m["comprimento"])
            if m.get("peso") is not None:
                pesos.append(m["peso"])

        dados["gramatura_media"] = {
            "largura": round(sum(larguras) / len(larguras), 4) if larguras else None,
            "comprimento": round(sum(comprimentos) / len(comprimentos), 2) if comprimentos else None,
            "peso": round(sum(pesos) / len(pesos), 3) if pesos else None,
        }

    @staticmethod
    def _calcular_total_kgs(dados):
        """Calcula o Total de KGs de Peso Teórico e Peso Real somando os KGs dos itens."""
        for secao in ("peso_teorico", "peso_real"):
            total = 0
            for item in dados[secao]["itens"]:
                kgs = item.get("kgs") or 0
                total += kgs
            dados[secao]["total_kgs"] = round(total, 2)

    @staticmethod
    def _calcular_resumo(dados):
        """Calcula KGs Coletados, KGs Consumidos e Saldo Devedor automaticamente."""
        # KGs Coletados = soma dos KGs de todas as OPs
        kgs_coletados = dados["ops_total"]["kgs"] or 0

        # KGs Consumidos = Total KGs do Peso Teórico
        kgs_consumidos = dados["peso_teorico"].get("total_kgs") or 0

        # Saldo Devedor = Coletados - Consumidos
        saldo_devedor = kgs_coletados - kgs_consumidos

        dados["resumo"] = {
            "kgs_coletados": round(kgs_coletados, 2),
            "kgs_consumidos": round(kgs_consumidos, 2),
            "saldo_devedor": round(saldo_devedor, 2),
        }

    @staticmethod
    def _calcular_percentuais(dados):
        """Calcula as % de cada item de Peso Teórico e Peso Real com base nos KGs coletados."""
        kgs_coletados = dados["resumo"].get("kgs_coletados") or 0

        for secao in ("peso_teorico", "peso_real"):
            total_pct = 0
            for item in dados[secao]["itens"]:
                kgs = item.get("kgs") or 0
                if kgs_coletados and kgs:
                    pct = (kgs / kgs_coletados) * 100
                else:
                    pct = 0
                item["pct"] = round(pct, 2)
                total_pct += pct

            total_kgs = dados[secao].get("total_kgs") or 0
            if kgs_coletados and total_kgs:
                dados[secao]["total_pct"] = round((total_kgs / kgs_coletados) * 100, 2)
            else:
                dados[secao]["total_pct"] = round(total_pct, 2)

    def gerar_dashboard(self):
        """Coleta dados e gera o dashboard HTML."""
        try:
            dados = self.coletar_dados()

            # Calcular automaticamente: total KGs, média, resumo e %
            self._calcular_total_kgs(dados)
            self._calcular_media_gramatura(dados)
            self._calcular_resumo(dados)
            self._calcular_percentuais(dados)

            # Validação básica
            if not dados["titulo"]:
                messagebox.showwarning("Atenção", "Preencha o título do container.")
                return
            if not dados["ops"]:
                messagebox.showwarning("Atenção", "Adicione pelo menos uma OP.")
                return
            if not dados["gramatura"]:
                messagebox.showwarning("Atenção", "Adicione pelo menos uma manta na gramatura.")
                return

            # Gerar HTML
            html = gerar_html(dados)

            # Salvar
            nome_arquivo = dados["titulo"].replace(" ", "_").replace("/", "-") + ".html"
            caminho_html = PASTA_BASE / nome_arquivo

            with open(caminho_html, "w", encoding="utf-8") as f:
                f.write(html)

            messagebox.showinfo(
                "Dashboard Gerado!",
                f"Arquivo gerado:\n\n"
                f"HTML: {caminho_html.name}\n\n"
                f"O dashboard será aberto no navegador."
            )

            webbrowser.open(str(caminho_html))

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar dashboard:\n{e}")


# ═══════════════════════════════════════════════════════════════════════════════
#  TELA INICIAL DE ESCOLHA
# ═══════════════════════════════════════════════════════════════════════════════

class TelaInicial:
    """Tela de boas-vindas com opção de preencher manualmente ou carregar CSV."""

    def __init__(self, root):
        self.root = root
        self.root.title("Fechamento de Containers")
        self.root.geometry("600x480")
        self.root.configure(bg="#1a73e8")
        self.root.resizable(False, False)
        self.escolha = None  # "manual" ou "csv"
        self.caminho_csv = None

        self._centralizar_janela(600, 480)
        self._build_ui()

    def _centralizar_janela(self, w, h):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        bg = "#1a73e8"

        # Ícone / Emoji grande
        tk.Label(
            self.root, text="📦", font=("Segoe UI Emoji", 48),
            bg=bg, fg="white"
        ).pack(pady=(35, 5))

        # Título
        tk.Label(
            self.root, text="Fechamento de Containers",
            font=("Segoe UI", 22, "bold"), bg=bg, fg="white"
        ).pack(pady=(0, 5))

        tk.Label(
            self.root, text="Gerador de Dashboard",
            font=("Segoe UI", 13), bg=bg, fg="#a8c7fa"
        ).pack(pady=(0, 30))

        # Separador
        tk.Frame(self.root, bg="#a8c7fa", height=1).pack(fill="x", padx=80)

        tk.Label(
            self.root, text="Como deseja iniciar?",
            font=("Segoe UI", 14), bg=bg, fg="white"
        ).pack(pady=(25, 15))

        # ── Botão: Preencher Manualmente ──
        btn_manual = tk.Button(
            self.root,
            text="✏️  Preencher Manualmente",
            font=("Segoe UI", 13, "bold"),
            bg="white", fg="#1a73e8",
            activebackground="#e8f0fe", activeforeground="#1557b0",
            relief="flat", padx=30, pady=12, cursor="hand2",
            width=30,
            command=self._escolher_manual
        )
        btn_manual.pack(pady=(0, 12))

        # ── Botão: Carregar CSV ──
        btn_csv = tk.Button(
            self.root,
            text="📂  Carregar Arquivo CSV",
            font=("Segoe UI", 13, "bold"),
            bg="#0d904f", fg="white",
            activebackground="#0a7a42", activeforeground="white",
            relief="flat", padx=30, pady=12, cursor="hand2",
            width=30,
            command=self._escolher_csv
        )
        btn_csv.pack(pady=(0, 20))

        # Rodapé
        tk.Label(
            self.root, text="Por Guedes e Erick",
            font=("Segoe UI", 10), bg=bg, fg="#a8c7fa"
        ).pack(side="bottom", pady=15)

    def _escolher_manual(self):
        self.escolha = "manual"
        self.root.destroy()

    def _escolher_csv(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo CSV",
            initialdir=str(PASTA_BASE),
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.escolha = "csv"
            self.caminho_csv = caminho
            self.root.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  EXECUÇÃO
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    # ── Tela inicial: escolher modo ──
    root_inicial = tk.Tk()
    try:
        root_inicial.iconbitmap(default="")
    except Exception:
        pass

    tela = TelaInicial(root_inicial)
    root_inicial.mainloop()

    # Se o usuário fechou sem escolher, sair
    if tela.escolha is None:
        return

    # ── Abrir formulário principal ──
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass

    app = AppFechamentoContainer(root)

    if tela.escolha == "csv" and tela.caminho_csv:
        try:
            linhas = ler_csv(Path(tela.caminho_csv))
            dados = extrair_dados_csv(linhas)
            app.preencher_formulario(dados)
        except Exception as e:
            messagebox.showerror("Erro ao carregar", f"Não foi possível ler o CSV:\n{e}")

    root.mainloop()


if __name__ == "__main__":
    main()
