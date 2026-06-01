#!/usr/bin/env python3
"""
RODAR DASH MAGALU DIRETORIA
============================
Script chamado automaticamente pelo Enchanté quando o usuário escreve:
"RODAR DASH MAGALU DIRETORIA"

O que faz:
1. Lê o arquivo TESTE BASE MAGALU.xlsx do Box
2. Filtra PARTNER=MAGALU HQ, exclui SubLOB=Total
3. Agrupa por meses: Abril(W1-W5), Maio(W6-W9), Junho(W10-W13)
4. Gera magalu_data.json com todos os KPIs
5. Gera magalu_dashboard.html com gráficos interativos
6. Salva no Box: GLAUCO HD/MYMALL GT/Projeto Magalu Glauco AI/
7. Faz push para GitHub: gtamega/magalu-diretoria-dash
"""

import pandas as pd
import json
import pathlib
import subprocess
import sys
import os

# ── Caminhos ─────────────────────────────────────────────────────
BOX_EXCEL = os.path.expanduser(
    "~/Library/CloudStorage/Box-Box/GLAUCO HD/MYMALL GT/Projeto Magalu Glauco AI/TESTE BASE MAGALU.xlsx"
)
BOX_OUT = os.path.expanduser(
    "~/Library/CloudStorage/Box-Box/GLAUCO HD/MYMALL GT/Projeto Magalu Glauco AI"
)
GH_REPO_DIR = "/tmp/magalu-diretoria-dash"
GH_TOKEN = "[GITHUB_TOKEN]"

# ── Ler dados ─────────────────────────────────────────────────────
print("📊 Lendo TESTE BASE MAGALU.xlsx...")
df = pd.read_excel(BOX_EXCEL, sheet_name=0)

# Filtrar MAGALU HQ e excluir SubLOB=Total
mag = df[(df['PARTNER'] == 'MAGALU HQ') & (df['SubLOB'] != 'Total')].copy()

# Meses
mag['ABRIL'] = mag[['W1','W2','W3','W4','W5']].sum(axis=1)
mag['MAIO']  = mag[['W6','W7','W8','W9']].sum(axis=1)
mag['JUNHO'] = mag[['W10','W11','W12','W13']].sum(axis=1)

MESES   = ['ABRIL','MAIO','JUNHO']
TABS    = ['MAGALU BRASIL'] + sorted(mag['DIRETORIA'].dropna().unique().tolist())
NPI     = ['iPhone 17','iPhone 17 Pro','iPhone 17 Pro Max','iPhone Air','iPhone 17e']
COLORS  = ['#0071e3','#34c759','#ff9500','#ff3b30','#af52de','#5ac8fa','#ffcc00','#ff6b35','#8e8e93','#30b0c7']

# ── Processar dados ───────────────────────────────────────────────
print("⚙️  Processando KPIs...")
data = {}
for tab in TABS:
    sub = mag.copy() if tab == 'MAGALU BRASIL' else mag[mag['DIRETORIA'] == tab].copy()
    data[tab] = {}
    brasil_total = {m: mag[m].sum() for m in MESES}

    for mes in MESES:
        total_mes = sub[mes].sum()
        top_model = sub.groupby('SubLOB')[mes].sum().sort_values(ascending=False)
        top_model = top_model[top_model > 0]
        modelo_top     = top_model.index[0] if len(top_model) > 0 else 'N/A'
        modelo_top_val = int(top_model.iloc[0]) if len(top_model) > 0 else 0
        npi_val    = sub[sub['SubLOB'].isin(NPI)][mes].sum()
        pct_npi    = round(npi_val / total_mes * 100, 1) if total_mes > 0 else 0
        pct_brasil = round(total_mes / brasil_total[mes] * 100, 1) if brasil_total[mes] > 0 else 0
        territ_v   = sub.groupby('TERRIT')[mes].sum().sort_values(ascending=False)
        territ_v   = territ_v[territ_v > 0]
        top_sublobs = sub.groupby('SubLOB')[mes].sum().sort_values(ascending=False).head(8).index.tolist()
        mix_data = {}
        for t in territ_v.index[:20]:
            t_sub   = sub[sub['TERRIT'] == t]
            t_total = t_sub[mes].sum()
            if t_total > 0:
                mix_data[t] = {sl: round(t_sub[t_sub['SubLOB']==sl][mes].sum()/t_total*100,1) for sl in top_sublobs}
        top15 = sub.groupby('Name')[mes].sum().sort_values(ascending=False).head(15)
        top15 = top15[top15 > 0]
        data[tab][mes] = {
            'total': int(total_mes), 'modelo_top': modelo_top,
            'modelo_top_val': modelo_top_val, 'pct_npi': pct_npi,
            'pct_brasil': pct_brasil,
            'territ': {k: int(v) for k,v in territ_v.items()},
            'mix': mix_data, 'top_sublobs': top_sublobs,
            'top15': {k: int(v) for k,v in top15.items()},
        }

# ── Gerar JSON ────────────────────────────────────────────────────
data_path = pathlib.Path(BOX_OUT) / 'magalu_data.json'
with open(data_path, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
print(f"✅ JSON salvo: {data_path}")

# ── Gerar HTML ────────────────────────────────────────────────────
print("🎨 Gerando HTML...")
DATA_JS   = json.dumps(data, ensure_ascii=False)
TABS_JS   = json.dumps(TABS)
COLORS_JS = json.dumps(COLORS)

ICONS = {'MAGALU BRASIL':'BR','MAGALU-MG/CO':'MG','MAGALU-N/NE':'NNE',
         'MAGALU-SP/RJ':'SP','MAGALU-SUL':'SUL','MAGALU-VIRTUAL':'VRT'}
ICON_COLORS = {'MAGALU BRASIL':'#0071e3','MAGALU-MG/CO':'#34c759',
               'MAGALU-N/NE':'#ff9500','MAGALU-SP/RJ':'#af52de',
               'MAGALU-SUL':'#ff3b30','MAGALU-VIRTUAL':'#5ac8fa'}

# [O HTML completo é gerado dinamicamente — ver magalu_dashboard.html]
# Importar e reusar o gerador HTML do arquivo original
# Por simplicidade, copiar o HTML existente e atualizar apenas o DATA_JS
html_path_src = pathlib.Path(BOX_OUT) / 'magalu_dashboard.html'
if html_path_src.exists():
    import re
    html = html_path_src.read_text(encoding='utf-8')
    # Substituir o bloco de dados
    html = re.sub(r'const DATA = \{.*?\};', f'const DATA = {DATA_JS};', html, flags=re.DOTALL, count=1)
    html_path_src.write_text(html, encoding='utf-8')
    print(f"✅ HTML atualizado: {html_path_src}")
else:
    print("⚠️  HTML base não encontrado, regenerar manualmente")

# ── Push GitHub ───────────────────────────────────────────────────
print("🐙 Fazendo push no GitHub...")
gh_dir = pathlib.Path(GH_REPO_DIR)
if gh_dir.exists():
    import shutil
    shutil.copy(str(html_path_src), str(gh_dir / 'magalu_dashboard.html'))
    shutil.copy(str(html_path_src), str(gh_dir / 'index.html'))
    shutil.copy(str(data_path), str(gh_dir / 'magalu_data.json'))
    os.environ['GH_TOKEN'] = GH_TOKEN
    result = subprocess.run(
        ['git', '-C', str(gh_dir), 'add', '.'],
        capture_output=True, text=True
    )
    result = subprocess.run(
        ['git', '-C', str(gh_dir), 'commit', '-m', '🔄 RODAR DASH MAGALU DIRETORIA — atualização automática'],
        capture_output=True, text=True
    )
    remote = f"https://gtamega:{GH_TOKEN}@github.com/gtamega/magalu-diretoria-dash.git"
    result = subprocess.run(
        ['git', '-C', str(gh_dir), 'push', remote, 'main'],
        capture_output=True, text=True
    )
    print("✅ GitHub atualizado!" if result.returncode == 0 else f"⚠️  GitHub: {result.stderr[:100]}")

print("\n🎉 RODAR DASH MAGALU DIRETORIA — concluído!")
print(f"   📦 Box: {BOX_OUT}/magalu_dashboard.html")
print(f"   🌐 GitHub Pages: https://gtamega.github.io/magalu-diretoria-dash")
