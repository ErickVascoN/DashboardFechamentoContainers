# -*- coding: utf-8 -*-
"""
Análise de Fechamento de Containers
Lê o CSV e gera relatório visual em HTML com tabelas e gráficos.
"""

import csv
import os
import re
import sys
from pathlib import Path

# Força UTF-8 no terminal Windows para suportar emojis
sys.stdout.reconfigure(encoding='utf-8')

# ─── Configuração ────────────────────────────────────────────────────────────
ARQUIVO_CSV = Path(__file__).parent / "FECHAMENTO CONTAINERS.csv"
ARQUIVO_HTML = Path(__file__).parent / "Dashboard Fechamento De Container.html"


# ─── Funções auxiliares ──────────────────────────────────────────────────────
def parse_numero(valor: str) -> float | None:
    """Converte string numérica brasileira (vírgula decimal) para float."""
    if valor is None:
        return None
    valor = valor.strip()
    if not valor or valor == "-":
        return None
    # Remove pontos de milhar e espaços
    valor = valor.replace(" ", "").replace(".", "")
    # Troca vírgula decimal por ponto
    valor = valor.replace(",", ".")
    # Remove % se houver
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


# ─── Leitura e parsing do CSV ────────────────────────────────────────────────
def ler_csv(caminho: Path) -> list[list[str]]:
    """Lê o CSV com delimitador ';'."""
    linhas = []
    with open(caminho, "r", encoding="utf-8") as f:
        leitor = csv.reader(f, delimiter=";")
        for linha in leitor:
            linhas.append(linha)
    return linhas


def extrair_dados(linhas: list[list[str]]) -> dict:
    """Extrai todas as seções de dados do CSV."""
    dados = {}

    # ── Título ──
    dados["titulo"] = linhas[0][0].strip()

    # ── Peso Teórico ──
    peso_teorico_unit = parse_numero(linhas[1][4])  # 0,6447
    dados["peso_teorico"] = {
        "peso_unitario": peso_teorico_unit,
        "itens": []
    }
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

    # ── Peso Real ──
    peso_real_unit = parse_numero(linhas[9][4])  # 0,678
    dados["peso_real"] = {
        "peso_unitario": peso_real_unit,
        "itens": []
    }
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

    # ── OPs Utilizadas ──
    dados["ops"] = []
    for i in range(2, 7):
        op = linhas[i][6].strip() if linhas[i][6].strip() else None
        kgs = parse_numero(linhas[i][7])
        pecas = parse_numero(linhas[i][8])
        if op and op != "TOTAL":
            dados["ops"].append({"op": op, "kgs": kgs, "pecas": pecas})
    # Total
    dados["ops_total"] = {
        "kgs": parse_numero(linhas[6][7]),
        "pecas": parse_numero(linhas[6][8]),
        "peso_medio": parse_numero(linhas[6][9])
    }

    # ── Resumo Total ──
    dados["resumo"] = {
        "kgs_coletados": parse_numero(linhas[10][7]),
        "kgs_consumidos": parse_numero(linhas[11][7]),
        "saldo_devedor": parse_numero(linhas[13][7]),
    }

    # ── Conferência de Gramatura ──
    dados["gramatura"] = []
    for i in range(3, 16):
        manta = linhas[i][11].strip() if len(linhas[i]) > 11 and linhas[i][11].strip() else None
        if manta and manta.startswith("#"):
            largura = parse_numero(linhas[i][12])
            comprimento = parse_numero(linhas[i][13])
            peso = parse_numero(linhas[i][14])
            dados["gramatura"].append({
                "manta": manta, "largura": largura,
                "comprimento": comprimento, "peso": peso
            })
    # Média
    dados["gramatura_media"] = {
        "largura": parse_numero(linhas[17][12]),
        "comprimento": parse_numero(linhas[17][13]),
        "peso": parse_numero(linhas[17][14]),
    }

    return dados


# ─── Geração do relatório HTML ───────────────────────────────────────────────
def gerar_html(dados: dict) -> str:
    """Gera relatório HTML completo com tabelas estilizadas e gráficos."""

    # Cálculos para gráficos
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
        .container {{
            max-width: 1400px;
            margin: 0 auto;
        }}
        .header {{
            background: linear-gradient(135deg, var(--primary), #1557b0);
            color: white;
            padding: 30px 40px;
            border-radius: 12px;
            margin-bottom: 24px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .header h1 {{
            font-size: 28px;
            font-weight: 600;
        }}
        .header p {{
            opacity: 0.9;
            margin-top: 4px;
            font-size: 14px;
        }}
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
        .card-header h2 {{
            font-size: 16px;
            font-weight: 600;
            color: var(--text);
        }}
        .card-header .badge {{
            font-size: 11px;
            padding: 2px 8px;
            border-radius: 12px;
            font-weight: 600;
        }}
        .badge-blue {{ background: var(--primary-light); color: var(--primary); }}
        .badge-green {{ background: #e6f4ea; color: var(--success); }}
        .badge-red {{ background: #fce8e6; color: var(--danger); }}
        .card-body {{
            padding: 16px 20px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }}
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
        td {{
            padding: 10px 12px;
            border-bottom: 1px solid #f0f0f0;
        }}
        tr:last-child td {{
            border-bottom: none;
        }}
        tr:hover td {{
            background: #fafbfc;
        }}
        .total-row {{
            font-weight: 700;
            background: var(--primary-light) !important;
        }}
        .total-row td {{
            border-top: 2px solid var(--primary);
            color: var(--primary);
        }}
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
        .kpi-value {{
            font-size: 28px;
            font-weight: 700;
            color: var(--text);
        }}
        .kpi-sub {{
            font-size: 12px;
            color: var(--text-secondary);
            margin-top: 2px;
        }}
        .kpi-positive {{ color: var(--success) !important; }}
        .kpi-negative {{ color: var(--danger) !important; }}
        .chart-container {{
            position: relative;
            height: 280px;
            padding: 10px;
        }}
        .full-width {{
            grid-column: 1 / -1;
        }}
        .comparison-bar {{
            display: flex;
            align-items: center;
            gap: 8px;
            margin: 8px 0;
        }}
        .comparison-bar .bar {{
            height: 24px;
            border-radius: 4px;
            transition: width 0.5s ease;
        }}
        .bar-teorico {{ background: #a8c7fa; }}
        .bar-real {{ background: var(--primary); }}
        .comparison-label {{
            font-size: 12px;
            min-width: 100px;
            color: var(--text-secondary);
        }}
        .comparison-value {{
            font-size: 13px;
            font-weight: 600;
            min-width: 80px;
            text-align: right;
        }}
        .footer {{
            text-align: center;
            padding: 20px;
            color: var(--text-secondary);
            font-size: 12px;
        }}
    </style>
</head>
<body>
<div class="container">

    <!-- Header -->
    <div class="header">
        <h1>📦 {dados['titulo']}</h1>
        <p>Relatório de análise de fechamento de container</p>
    </div>

    <!-- KPIs -->
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

        <!-- Peso Teórico -->
        <div class="card">
            <div class="card-header">
                <h2>⚖️ Peso Teórico</h2>
                <span class="badge badge-blue">Unit: {fmt_numero(dados['peso_teorico']['peso_unitario'], 4)}</span>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr>
                            <th>Item</th>
                            <th class="text-right">Peças</th>
                            <th class="text-right">KGs</th>
                            <th class="text-right">Peso Unit.</th>
                            <th class="text-right">%</th>
                        </tr>
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
                            <td>TOTAL</td>
                            <td class="text-right"></td>
                            <td class="text-right">{fmt_numero(dados['peso_teorico']['total_kgs'], 2)}</td>
                            <td class="text-right"></td>
                            <td class="text-right">{fmt_pct(dados['peso_teorico']['total_pct'])}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Peso Real -->
        <div class="card">
            <div class="card-header">
                <h2>📐 Peso Real</h2>
                <span class="badge badge-green">Unit: {fmt_numero(dados['peso_real']['peso_unitario'], 4)}</span>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr>
                            <th>Item</th>
                            <th class="text-right">Peças</th>
                            <th class="text-right">KGs</th>
                            <th class="text-right">Peso Unit.</th>
                            <th class="text-right">%</th>
                        </tr>
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
                            <td>TOTAL</td>
                            <td class="text-right"></td>
                            <td class="text-right">{fmt_numero(dados['peso_real']['total_kgs'], 2)}</td>
                            <td class="text-right"></td>
                            <td class="text-right">{fmt_pct(dados['peso_real']['total_pct'])}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <!-- OPs Utilizadas -->
        <div class="card">
            <div class="card-header">
                <h2>📋 OPs Utilizadas</h2>
                <span class="badge badge-blue">{len(dados['ops'])} OPs</span>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr>
                            <th>OP</th>
                            <th class="text-right">KGs</th>
                            <th class="text-right">Peças</th>
                        </tr>
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

        <!-- Resumo Total -->
        <div class="card">
            <div class="card-header">
                <h2>📊 Resumo Total</h2>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr>
                            <th>Indicador</th>
                            <th class="text-right">Valor (KG)</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>KGs Coletados</td>
                            <td class="text-right">{fmt_numero(dados['resumo']['kgs_coletados'], 2)}</td>
                        </tr>
                        <tr>
                            <td>KGs Consumidos</td>
                            <td class="text-right">{fmt_numero(dados['resumo']['kgs_consumidos'], 2)}</td>
                        </tr>
                        <tr class="total-row">
                            <td>Saldo Devedor</td>
                            <td class="text-right">{fmt_numero(dados['resumo']['saldo_devedor'], 2)}</td>
                        </tr>
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

        <!-- Conferência de Gramatura -->
        <div class="card full-width">
            <div class="card-header">
                <h2>🔬 Conferência de Gramatura</h2>
                <span class="badge badge-blue">{len(dados['gramatura'])} mantas</span>
            </div>
            <div class="card-body">
                <table>
                    <thead>
                        <tr>
                            <th>Manta</th>
                            <th class="text-right">Largura (m)</th>
                            <th class="text-right">Comprimento (m)</th>
                            <th class="text-right">Peso (kg)</th>
                        </tr>
                    </thead>
                    <tbody>"""

    peso_teorico = dados['peso_teorico']['peso_unitario']
    for m in dados["gramatura"]:
        # Marcar em vermelho se diferença >= 0,006 do teórico, verde se <= 0,005
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

    <!-- Gráficos -->
    <div class="grid">
        <div class="card">
            <div class="card-header">
                <h2>📈 Distribuição de KGs por OP</h2>
            </div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="chartOps"></canvas>
                </div>
            </div>
        </div>
        <div class="card">
            <div class="card-header">
                <h2>📉 Peso por Manta (Gramatura)</h2>
            </div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="chartGramatura"></canvas>
                </div>
            </div>
        </div>
        <div class="card">
            <div class="card-header">
                <h2>⚖️ Peso Teórico vs Real (% composição)</h2>
            </div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="chartComparacao"></canvas>
                </div>
            </div>
        </div>
        <div class="card">
            <div class="card-header">
                <h2>📊 Peças por OP</h2>
            </div>
            <div class="card-body">
                <div class="chart-container">
                    <canvas id="chartPecasOp"></canvas>
                </div>
            </div>
        </div>
    </div>

    <div class="footer">
        Análise de Fechamento de Containers <br> Por <br>
          Guedes e Erick
    </div>

</div>

<script>
    // Cores
    const cores = ['#1a73e8', '#34a853', '#fbbc04', '#ea4335', '#673ab7', '#ff6d00'];
    const coresAlpha = cores.map(c => c + '33');

    // Gráfico OPs (barras)
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
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
                legend: {{ display: false }},
                tooltip: {{
                    callbacks: {{
                        label: ctx => ctx.parsed.y.toLocaleString('pt-BR') + ' kg'
                    }}
                }}
            }},
            scales: {{
                y: {{
                    beginAtZero: true,
                    ticks: {{ callback: v => v.toLocaleString('pt-BR') + ' kg' }}
                }}
            }}
        }}
    }});

    // Gráfico Gramatura (linha)
    new Chart(document.getElementById('chartGramatura'), {{
        type: 'line',
        data: {{
            labels: {[m['manta'] for m in dados['gramatura']]},
            datasets: [{{
                label: 'Peso (kg)',
                data: {[m['peso'] for m in dados['gramatura']]},
                borderColor: '#1a73e8',
                backgroundColor: '#1a73e833',
                fill: true,
                tension: 0.3,
                pointRadius: 5,
                pointBackgroundColor: '#1a73e8'
            }},
            {{
                label: 'Média',
                data: Array({len(dados['gramatura'])}).fill({dados['gramatura_media']['peso']}),
                borderColor: '#ea4335',
                borderDash: [5, 5],
                pointRadius: 0,
                fill: false
            }},
            {{
                label: 'Peso Teórico ({fmt_numero(dados["peso_teorico"]["peso_unitario"], 4)} kg)',
                data: Array({len(dados['gramatura'])}).fill({dados['peso_teorico']['peso_unitario']}),
                borderColor: '#0d904f',
                borderDash: [10, 4],
                borderWidth: 2,
                pointRadius: 0,
                fill: false
            }}]
        }},
        options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
                tooltip: {{
                    callbacks: {{
                        label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y.toFixed(4) + ' kg'
                    }}
                }}
            }},
            scales: {{
                y: {{
                    min: 0.60,
                    max: 0.73
                }}
            }}
        }}
    }});

    // Gráfico Comparação Teórico vs Real (barras agrupadas)
    new Chart(document.getElementById('chartComparacao'), {{
        type: 'bar',
        data: {{
            labels: {[item['nome'] for item in dados['peso_teorico']['itens']]},
            datasets: [
                {{
                    label: 'Teórico (%)',
                    data: {[item['pct'] if item['pct'] else 0 for item in dados['peso_teorico']['itens']]},
                    backgroundColor: '#a8c7fa',
                    borderRadius: 4
                }},
                {{
                    label: 'Real (%)',
                    data: {[item['pct'] if item['pct'] else 0 for item in dados['peso_real']['itens']]},
                    backgroundColor: '#1a73e8',
                    borderRadius: 4
                }}
            ]
        }},
        options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
                tooltip: {{
                    callbacks: {{
                        label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y.toFixed(2) + '%'
                    }}
                }}
            }},
            scales: {{
                y: {{
                    beginAtZero: true,
                    ticks: {{ callback: v => v + '%' }}
                }}
            }}
        }}
    }});

    // Gráfico Peças por OP (donut)
    new Chart(document.getElementById('chartPecasOp'), {{
        type: 'doughnut',
        data: {{
            labels: {[op['op'] for op in dados['ops']]},
            datasets: [{{
                data: {[op['pecas'] for op in dados['ops']]},
                backgroundColor: cores,
                borderWidth: 2,
                borderColor: '#fff'
            }}]
        }},
        options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{
                tooltip: {{
                    callbacks: {{
                        label: ctx => {{
                            const total = ctx.dataset.data.reduce((a,b) => a+b, 0);
                            const pct = (ctx.parsed / total * 100).toFixed(1);
                            return ctx.label + ': ' + ctx.parsed.toLocaleString('pt-BR') + ' peças (' + pct + '%)';
                        }}
                    }}
                }}
            }}
        }}
    }});
</script>

</body>
</html>"""

    return html


# ─── Execução principal ─────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  ANÁLISE DE FECHAMENTO DE CONTAINERS")
    print("=" * 60)

    # Ler CSV
    print(f"\n📂 Lendo arquivo: {ARQUIVO_CSV.name}")
    linhas = ler_csv(ARQUIVO_CSV)
    print(f"   ✅ {len(linhas)} linhas lidas")

    # Extrair dados
    dados = extrair_dados(linhas)
    print(f"   ✅ Dados extraídos com sucesso")

    # Exibir resumo no terminal
    print(f"\n{'─' * 60}")
    print(f"  📦 {dados['titulo']}")
    print(f"{'─' * 60}")

    print(f"\n  🔹 Peso Teórico Unitário: {fmt_numero(dados['peso_teorico']['peso_unitario'], 4)} kg")
    print(f"  🔹 Peso Real Unitário:    {fmt_numero(dados['peso_real']['peso_unitario'], 4)} kg")
    dif = dados['peso_real']['peso_unitario'] - dados['peso_teorico']['peso_unitario']
    print(f"  🔹 Diferença:             {'+' if dif > 0 else ''}{fmt_numero(dif, 4)} kg")

    print(f"\n  📋 OPs Utilizadas:")
    for op in dados['ops']:
        print(f"     OP {op['op']}: {fmt_numero(op['kgs'], 1)} kg | {fmt_numero(op['pecas'], 0)} peças")

    print(f"\n  📊 Resumo:")
    print(f"     KGs Coletados:  {fmt_numero(dados['resumo']['kgs_coletados'], 2)}")
    print(f"     KGs Consumidos: {fmt_numero(dados['resumo']['kgs_consumidos'], 2)}")
    print(f"     Saldo Devedor:  {fmt_numero(dados['resumo']['saldo_devedor'], 2)} kg")

    print(f"\n  🔬 Gramatura Média:")
    print(f"     Largura:      {fmt_numero(dados['gramatura_media']['largura'], 4)} m")
    print(f"     Comprimento:  {fmt_numero(dados['gramatura_media']['comprimento'], 2)} m")
    print(f"     Peso:         {fmt_numero(dados['gramatura_media']['peso'], 3)} kg")

    # Gerar relatório HTML
    html = gerar_html(dados)
    with open(ARQUIVO_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n{'─' * 60}")
    print(f"  ✅ Relatório HTML gerado: {ARQUIVO_HTML.name}")
    print(f"  📂 Caminho: {ARQUIVO_HTML}")
    print(f"{'─' * 60}")

    # Abrir no navegador
    import webbrowser
    webbrowser.open(str(ARQUIVO_HTML))
    print("  🌐 Abrindo no navegador...")


if __name__ == "__main__":
    main()
