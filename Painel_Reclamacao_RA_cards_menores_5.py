import json
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ============================================================
# CONFIGURAÇÕES
# ============================================================
ARQUIVO_EXCEL = Path(r"C:/Users/GabriellyOliveira/Desktop/Processo CSOPS/07. Teste/Power Bi Reclamação/Analitico_Reclamação.xlsx")
PASTA_SAIDA = ARQUIVO_EXCEL.parent
ARQUIVO_SAIDA = PASTA_SAIDA / "Painel_Reclamação.html"
ARQUIVO_BASE_COMPLETA = PASTA_SAIDA / "Base_Analitica_Reclamacao_Completa.csv"

# Página pública do Reclame Aqui da Digisac.
# Não usa login, senha nem API contratada.
URL_RECLAME_AQUI_PUBLICO = "https://www.reclameaqui.com.br/empresa/e-m-solucoes-integradas-de-sistemas/"
ARQUIVO_RECLAME_AQUI_CSV = PASTA_SAIDA / "reclame_aqui_digisac_indicadores.csv"
ARQUIVO_RECLAME_AQUI_DEBUG = PASTA_SAIDA / "debug_reclame_aqui_texto.txt"

MESES_PT = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
MESES_LONGO = {1:"jan",2:"fev",3:"mar",4:"abr",5:"mai",6:"jun",7:"jul",8:"ago",9:"set",10:"out",11:"nov",12:"dez"}

# ============================================================
# FUNÇÕES AUXILIARES PYTHON
# ============================================================
def limpar_texto(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def normalizar_vazio(x, padrao="Não informado"):
    x = limpar_texto(x)
    return x if x else padrao

def motivo_final(row, motivo_cols):
    dep = limpar_texto(row.get("Departamento responsável", ""))
    if dep:
        alvo = f"Motivo da reclamação - {dep}"
        if alvo in motivo_cols and limpar_texto(row.get(alvo, "")):
            return limpar_texto(row.get(alvo, ""))
    for col in motivo_cols:
        val = limpar_texto(row.get(col, ""))
        if val:
            return val
    return "Sem motivo"


def extrair_numero_apos_linha(linhas, texto_alvo):
    """
    Procura uma linha que contenha o texto_alvo e retorna o primeiro número encontrado
    nas próximas linhas.
    """
    texto_alvo = texto_alvo.lower()
    for i, linha in enumerate(linhas):
        if texto_alvo in linha.lower():
            for prox in linhas[i + 1:i + 8]:
                m = re.search(r"\d+(?:[.,]\d+)?%?", prox)
                if m:
                    return m.group(0).replace(",", ".")
    return ""


def coletar_reclame_aqui_publico():
    """
    Coleta automaticamente os indicadores públicos da página da Digisac no Reclame Aqui.
    Se o site mudar ou bloquear a coleta, o painel continua gerando e os cards aparecem vazios.
    """
    dados = {
        "url": URL_RECLAME_AQUI_PUBLICO,
        "periodo": "",
        "reputacao": "",
        "nota_media_reputacao": "",
        "qtd_reclamacoes": "",
        "reclamacoes_respondidas_pct": "",
        "nao_respondidas": "",
        "avaliadas": "",
        "nota_consumidor": "",
        "voltariam_fazer_negocio_pct": "",
        "indice_solucao_pct": "",
        "tempo_medio_resposta": "",
        "coletado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
            "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
        }

        resp = requests.get(URL_RECLAME_AQUI_PUBLICO, headers=headers, timeout=40)
        resp.raise_for_status()

        soup = BeautifulSoup(resp.text, "html.parser")
        texto = soup.get_text("\n", strip=True)
        ARQUIVO_RECLAME_AQUI_DEBUG.write_text(texto, encoding="utf-8")

        linhas = [l.strip() for l in texto.splitlines() if l.strip()]
        texto_unico = "\n".join(linhas)

        # Reputação: normalmente aparece logo depois da palavra "Reputação".
        for i, linha in enumerate(linhas):
            if linha.strip().lower() == "reputação" and i + 1 < len(linhas):
                dados["reputacao"] = linhas[i + 1].strip()
                break

        # Nota média de reputação nos últimos 6 meses.
        m = re.search(
            r"nota média nos últimos 6 meses é\s*\n?\s*(\d+(?:[.,]\d+)?)\s*\n?\s*/10",
            texto_unico,
            flags=re.I
        )
        if m:
            dados["nota_media_reputacao"] = m.group(1).replace(",", ".")

        # Quantidade de reclamações recebidas.
        m = re.search(r"recebeu\s*\n?\s*(\d+)\s+reclamaç", texto_unico, flags=re.I)
        if m:
            dados["qtd_reclamacoes"] = m.group(1)

        # Percentual respondido.
        m = re.search(r"Respondeu\s*\n?\s*(\d+(?:[.,]\d+)?%)", texto_unico, flags=re.I)
        if m:
            dados["reclamacoes_respondidas_pct"] = m.group(1).replace(",", ".")

        # Não respondidas.
        m = re.search(r"Há\s*\n?\s*(\d+)\s+reclamaç[aã]o\s*\n?\s*aguardando resposta", texto_unico, flags=re.I)
        if m:
            dados["nao_respondidas"] = m.group(1)

        # Avaliadas e nota do consumidor.
        m = re.search(
            r"Há\s*\n?\s*(\d+)\s+reclamações\s*\n?\s*avaliadas.*?é\s*\n?\s*(\d+(?:[.,]\d+)?)",
            texto_unico,
            flags=re.I | re.S
        )
        if m:
            dados["avaliadas"] = m.group(1)
            dados["nota_consumidor"] = m.group(2).replace(",", ".")

        # Voltariam a fazer negócio.
        m = re.search(r"Dos que avaliaram,\s*\n?\s*(\d+(?:[.,]\d+)?%)\s+voltariam", texto_unico, flags=re.I)
        if m:
            dados["voltariam_fazer_negocio_pct"] = m.group(1).replace(",", ".")

        # Índice de solução.
        m = re.search(r"resolveu\s*\n?\s*(\d+(?:[.,]\d+)?%)\s+das reclamações", texto_unico, flags=re.I)
        if m:
            dados["indice_solucao_pct"] = m.group(1).replace(",", ".")

        # Tempo médio de resposta.
        m = re.search(r"O tempo médio de resposta é\s*\n?\s*([^\n]+)", texto_unico, flags=re.I)
        if m:
            dados["tempo_medio_resposta"] = m.group(1).strip()

        # Período.
        m = re.search(
            r"Os dados correspondem ao período de\s*\n?\s*(\d{2}/\d{2}/\d{4}\s+a\s+\d{2}/\d{2}/\d{4})",
            texto_unico,
            flags=re.I
        )
        if m:
            dados["periodo"] = m.group(1)

    except Exception as e:
        print(f"[AVISO] Não foi possível coletar os dados públicos do Reclame Aqui: {e}")

    pd.DataFrame([dados]).to_csv(ARQUIVO_RECLAME_AQUI_CSV, index=False, sep=";", encoding="utf-8-sig")
    return dados

# ============================================================
# LEITURA E TRATAMENTO DA BASE
# ============================================================
df = pd.read_excel(ARQUIVO_EXCEL)
df.columns = [str(c).strip() for c in df.columns]

colunas_esperadas = [
    "Título", "Fase atual", "Data da manifestação", "Finalizado em", "Etiquetas", "URL",
    "Departamento responsável", "Problema resolvido?", "Voltaria a fazer negócio?",
    "Nota do atendimento", "Observações", "Canal da reclamação",
]

faltantes = [c for c in colunas_esperadas if c not in df.columns]
if faltantes:
    raise ValueError(f"Colunas não encontradas na planilha: {faltantes}")

motivo_cols = [c for c in df.columns if c.startswith("Motivo da reclamação - ")]
if not motivo_cols:
    raise ValueError("Não encontrei colunas de motivo. Elas precisam começar com: 'Motivo da reclamação - '")

df["Data da manifestação"] = pd.to_datetime(df["Data da manifestação"], errors="coerce", dayfirst=True)
df = df[~df["Data da manifestação"].isna()].copy()
df["Ano"] = df["Data da manifestação"].dt.year.astype(int)
df["MesNum"] = df["Data da manifestação"].dt.month.astype(int)
df = df[df["Ano"] >= 2025].copy()

if df.empty:
    raise ValueError("Após filtrar somente anos a partir de 2025, a base ficou vazia.")

df["Mês"] = df["MesNum"].map(MESES_PT)
df["AnoMes"] = df["Ano"].astype(str) + "-" + df["MesNum"].astype(str).str.zfill(2)
df["PeriodoLabel"] = df["MesNum"].map(MESES_LONGO) + "/" + df["Ano"].astype(str)

# Tratamento da coluna Finalizado em
# Regra: para Reclame Aqui pode ficar como "Não se aplica";
# para os demais canais, quando vier vazio, aparece como "Não informado" para facilitar correção da base.
df["Finalizado em"] = pd.to_datetime(df["Finalizado em"], errors="coerce", dayfirst=True)
canal_temp = df["Canal da reclamação"].apply(lambda x: limpar_texto(x).lower())
df["Finalizado em"] = np.where(
    df["Finalizado em"].isna() & canal_temp.str.contains("reclame", na=False),
    "Não se aplica",
    np.where(
        df["Finalizado em"].isna(),
        "Não informado",
        df["Finalizado em"].dt.strftime("%d/%m/%Y")
    )
)

for col in ["Título", "Fase atual", "Etiquetas", "URL", "Departamento responsável", "Problema resolvido?", "Voltaria a fazer negócio?", "Observações", "Canal da reclamação"]:
    df[col] = df[col].apply(normalizar_vazio)

df["Nota do atendimento"] = pd.to_numeric(df["Nota do atendimento"], errors="coerce")
df["Motivo Final"] = df.apply(lambda r: motivo_final(r, motivo_cols), axis=1)

# Base que será salva automaticamente ao rodar o Python
base_export = df.copy()
base_export["Data da manifestação"] = base_export["Data da manifestação"].dt.strftime("%d/%m/%Y")
base_export.to_csv(ARQUIVO_BASE_COMPLETA, index=False, sep=";", encoding="utf-8-sig")

# Etiquetas expandidas
regs_etiquetas = []
for _, r in df.iterrows():
    txt = limpar_texto(r["Etiquetas"])
    if not txt or txt == "Não informado":
        continue
    partes = [p.strip() for p in txt.replace(";", ",").split(",") if p.strip()]
    for et in partes:
        regs_etiquetas.append({"Ano": int(r.Ano), "MesNum": int(r.MesNum), "Mês": r["Mês"], "AnoMes": r.AnoMes, "Departamento responsável": r["Departamento responsável"], "Fase atual": r["Fase atual"], "Etiqueta": et})
df_etiquetas = pd.DataFrame(regs_etiquetas)
if df_etiquetas.empty:
    df_etiquetas = pd.DataFrame(columns=["Ano","MesNum","Mês","AnoMes","Departamento responsável","Fase atual","Etiqueta"])

# Motivos expandidos por departamento de motivo
regs_motivos = []
for _, r in df.iterrows():
    for col in motivo_cols:
        val = limpar_texto(r[col])
        if not val:
            continue
        dep_motivo = col.replace("Motivo da reclamação - ", "").strip()
        partes = [p.strip() for p in val.replace(";", ",").split(",") if p.strip()]
        for mot in partes:
            regs_motivos.append({
                "Ano": int(r.Ano), "MesNum": int(r.MesNum), "Mês": r["Mês"], "AnoMes": r.AnoMes,
                "Departamento responsável": r["Departamento responsável"], "Departamento do motivo": dep_motivo,
                "Motivo": mot, "Canal da reclamação": r["Canal da reclamação"], "Fase atual": r["Fase atual"]
            })
df_motivos = pd.DataFrame(regs_motivos)
if df_motivos.empty:
    df_motivos = pd.DataFrame(columns=["Ano","MesNum","Mês","AnoMes","Departamento responsável","Departamento do motivo","Motivo","Canal da reclamação","Fase atual"])

base_principal = df[[
    "Título", "Fase atual", "Data da manifestação", "Finalizado em", "Etiquetas", "URL", "Departamento responsável",
    "Problema resolvido?", "Voltaria a fazer negócio?", "Nota do atendimento", "Observações",
    "Canal da reclamação", "Motivo Final", "Preencha o outro motivo:", "Ano", "MesNum", "Mês", "AnoMes", "PeriodoLabel"
]].copy()
base_principal["Data da manifestação"] = base_principal["Data da manifestação"].dt.strftime("%d/%m/%Y")

# Renomeia a coluna para aparecer no analítico
base_principal.rename(
    columns={"Preencha o outro motivo:": "Outro Motivo"},
    inplace=True
)

base_principal = base_principal.replace({np.nan: None})

DATA_ATUALIZACAO = datetime.now().strftime("%d/%m/%Y as %H:%M")
dados_reclame_aqui = coletar_reclame_aqui_publico()

payload = {
    "atualizado": DATA_ATUALIZACAO,
    "reclame_aqui": dados_reclame_aqui,
    "principal": base_principal.to_dict(orient="records"),
    "motivos": df_motivos.replace({np.nan: None}).to_dict(orient="records"),
    "etiquetas": df_etiquetas.replace({np.nan: None}).to_dict(orient="records"),
    "anos": sorted(df["Ano"].dropna().unique().astype(int).tolist()),
    "departamentos": sorted(df["Departamento responsável"].dropna().unique().tolist()),
    "canais": sorted(df["Canal da reclamação"].dropna().unique().tolist()),
    "fases": sorted(df["Fase atual"].dropna().unique().tolist()),
    "meses": [{"num": k, "nome": v, "longo": MESES_LONGO[k]} for k, v in MESES_PT.items()],
}
payload_json = json.dumps(payload, ensure_ascii=False)

html = r'''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Painel Reclamação</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<script src="https://unpkg.com/lucide@latest"></script>
<style>
:root{
  --bg:#f6f9fb; --card:#fff; --border:#d7e4ee; --text:#1f2937; --muted:#71829a;
  --cyan:#19b5cf; --cyan2:#dff8fb; --cyan3:#eefcff; --blue:#1d8ff2; --navy:#182f9f;
  --greenSoft:#dcfce3; --shadow:0 2px 8px rgba(20,40,60,.05); --radius:16px;
}
*{box-sizing:border-box} body{margin:0;background:var(--bg);font-family:Segoe UI,Arial,sans-serif;color:var(--text)}
.app{display:flex;min-height:100vh}.side{width:88px;background:#fff;border-right:1px solid #e6eef5;display:flex;flex-direction:column;align-items:center;padding-top:170px;gap:42px;position:fixed;left:0;top:0;bottom:0}.nav{width:70px;height:70px;border-radius:50%;display:flex;align-items:center;justify-content:center;color:#8fa0ad;cursor:pointer}.nav.active{background:var(--cyan2);color:var(--cyan);}.nav.green.active{background:var(--greenSoft)}
.lucide{width:22px;height:22px;stroke-width:2.2}.nav .lucide{width:24px;height:24px}.icon-badge .lucide{width:19px;height:19px;color:var(--cyan)}.home-card .icon-badge .lucide{width:30px;height:30px}.return .lucide{width:30px;height:30px;color:#3898db;stroke-width:2.4}
.main{margin-left:88px;width:calc(100% - 88px);padding:8px 16px 24px}.app.home-mode .side{display:none}.app.home-mode .main{margin-left:0;width:100%;padding:0}.app.home-mode .home-cover{min-height:100vh}.page{display:none}.page.active{display:block}.hero{background:#fff;border:1px solid var(--border);border-radius:20px;box-shadow:var(--shadow);padding:18px 22px;margin-bottom:18px}.hero-title{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}.h-left{display:flex;align-items:center;gap:12px}.icon-badge{width:34px;height:34px;border-radius:50%;background:var(--cyan2);display:flex;align-items:center;justify-content:center;color:var(--cyan);font-weight:800;font-size:18px}.hero h1{margin:0;font-size:30px;line-height:1;font-weight:800;color:#222}.return{font-size:34px;color:#3898db}.filters{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.filters.three{grid-template-columns:repeat(3,1fr)}.fg label{display:block;font-weight:700;font-size:16px;margin:0 0 4px;color:#3d4650}.fg select{width:100%;height:45px;border:1px solid #cbd6e2;border-radius:8px;background:white;padding:0 12px;font-size:15px;color:#666;outline:none}
.fg select[multiple]{display:none}
.ms-wrap{position:relative;width:100%}
.ms-display{height:45px;border:1px solid #cbd6e2;border-radius:8px;background:#fff;padding:0 38px 0 12px;font-size:15px;color:#1f2937;display:flex;align-items:center;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ms-display:after{content:"⌄";position:absolute;right:14px;top:10px;color:#1f2937;font-size:16px;pointer-events:none}
.ms-wrap.open .ms-display{border-color:var(--cyan);box-shadow:0 0 0 2px rgba(25,181,207,.18)}
.ms-panel{display:none;position:absolute;z-index:30;top:50px;left:0;width:min(220px,100%);background:#fff;border:1px solid #cbd6e2;border-radius:12px;box-shadow:0 8px 22px rgba(15,35,55,.16);padding:10px}
.ms-wrap.open .ms-panel{display:block}
.ms-actions{display:flex;gap:8px;margin-bottom:8px}
.ms-btn{border:1px solid #d7e4ee;background:#f8fbff;color:#063f78;border-radius:8px;padding:8px 12px;font-size:12px;font-weight:800;cursor:pointer}
.ms-option{display:flex;align-items:center;gap:8px;padding:7px 4px;font-size:14px;color:#1f2937;cursor:pointer}
.ms-option input{width:14px;height:14px;accent-color:var(--cyan)}
.home-subtitle{max-width:650px;margin-top:-36px;margin-bottom:44px;font-size:17px;line-height:1.55;color:#eafcff}
.home-updated{position:absolute;right:40px;bottom:36px;font-size:16px;color:#fff;display:flex;align-items:center;gap:10px}.home-updated:before{content:"";width:8px;height:8px;border-radius:50%;background:#a9f3ff;display:inline-block}
.home-cover{position:relative}
.base-analitica{margin-top:34px}
.status-card{background:#fff}.card{background:#fff;border:1px solid var(--border);border-radius:8px;box-shadow:var(--shadow);padding:12px;margin-bottom:20px}.card.round{border-radius:16px}.kpi-row{display:flex;gap:12px;flex-wrap:wrap}.kpi{width:218px;border:1px solid var(--border);border-radius:16px;background:#fff;padding:24px 20px}.kpi .label{font-size:13px;color:#56677d;margin-bottom:10px}.kpi .value{font-size:34px;font-weight:800;color:var(--cyan)}.kpi .sub{font-size:14px;color:var(--cyan);font-weight:700;margin-top:12px}
.geral-top{display:grid;grid-template-columns:420px 1fr;gap:20px;align-items:start}
.ra-fonte-info{display:flex;align-items:center;gap:7px;font-size:11px;color:#7890ad;background:#f4faff;border:1px solid #d7eaf8;border-radius:8px;padding:7px 12px;margin-bottom:10px;line-height:1.4}
.ra-fonte-info a{color:var(--cyan);font-weight:700;text-decoration:none}
.ra-fonte-info a:hover{text-decoration:underline}
.ra-cards-all{display:grid;grid-template-columns:repeat(8,minmax(0,1fr));gap:8px;width:100%}
.ra-cards{display:grid;grid-template-columns:repeat(3,minmax(150px,1fr));gap:10px;width:100%}
.ra-card{border:1px solid var(--border);border-radius:12px;background:#fff;padding:10px 12px;min-height:80px;box-shadow:var(--shadow)}
.ra-card.ra-card-kpi{border-left:3px solid var(--cyan)}
.ra-card.destaque{background:linear-gradient(135deg,#eefcff,#ffffff)}
.ra-label{font-size:11px;color:#56677d;margin-bottom:5px}
.ra-label-row{display:flex;align-items:center;justify-content:space-between;gap:4px;margin-bottom:5px}
.ra-value{font-size:20px;font-weight:900;color:#071f3f;line-height:1.05}
.ra-value.cyan{color:var(--cyan)}
.ra-sub{font-size:11px;color:#7890ad;margin-top:5px;line-height:1.2}.section-title{font-size:20px;font-weight:700;margin:0 0 6px;color:#333}.chart-grid-2{display:grid;grid-template-columns:1fr 1fr;gap:34px 40px}.chart-box h3{margin:0 0 8px;font-size:18px;color:#333}.plot{height:180px}.plot.tall{height:250px}.dept-main{height:145px}.month-dept{height:160px}#chartDeptoAno,#chartDeptoMes{border-radius:18px;overflow:hidden;background:#fff}#chartDeptoAno .main-svg,#chartDeptoMes .main-svg{border-radius:18px}.summary{font-size:15px;line-height:1.55;color:#24364b;padding:22px}.summary h3{margin:0 0 14px;font-size:22px;color:#061d3a}.summary b{font-weight:800}.motivos-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}.motivo-card{border:1px solid var(--border);border-radius:16px;background:#fff;padding:18px}.motivo-head{display:flex;justify-content:space-between;align-items:flex-start}.motivo-head h3{font-size:18px;margin:0}.motivo-total{color:var(--cyan);font-size:28px;font-weight:800;text-align:right}.small-muted{font-size:13px;color:#7890ad;margin-top:4px}.mini-table{width:100%;border-collapse:collapse;margin-top:18px;font-size:13px}.mini-table th{text-align:left;color:#5c6f88;font-size:13px;border-bottom:1px solid #d7e4ee;padding:9px}.mini-table td{border-bottom:1px solid #edf2f6;padding:9px}.mini-table th:nth-child(2),.mini-table td:nth-child(2){text-align:center;font-weight:700}.mini-table th:nth-child(3),.mini-table td:nth-child(3){text-align:right;color:#5e7390}.analytics-head{display:flex;justify-content:space-between;gap:14px;align-items:center}.btns{display:flex;gap:8px;flex-wrap:wrap}.btn{border:1px solid var(--border);background:#fff;color:#1d5d88;border-radius:10px;padding:10px 12px;font-weight:700;cursor:pointer}.btn.primary{background:var(--cyan);border-color:var(--cyan);color:#fff}.table-wrap{max-height:360px;overflow:auto;border:1px solid var(--border);border-radius:8px;background:#fff}.detail-table{width:100%;border-collapse:collapse;font-size:12px}.detail-table th{position:sticky;top:0;background:#f6f9fc;color:#60738f;text-align:left;padding:12px;border-bottom:1px solid #e7edf3}.detail-table td{padding:12px;border-bottom:1px solid #edf2f6;vertical-align:top}.pill{display:inline-block;background:var(--cyan2);color:#0e91a7;border-radius:999px;padding:5px 10px;font-weight:800;font-size:12px}.quality-grid{display:grid;grid-template-columns:minmax(0,.85fr) minmax(0,1fr) minmax(0,1.25fr);gap:14px;align-items:start;width:100%;overflow:hidden}.quality-grid>div{min-width:0}.status-card{border:1px solid var(--border);border-radius:16px;padding:18px;margin-bottom:16px}.status-card h3{margin:0 0 2px;font-size:21px}.status-num{float:right;color:var(--cyan);font-size:30px;font-weight:800}.bar-bg{height:8px;background:#e5e9ef;border-radius:999px;margin:14px 0 8px}.bar{height:8px;background:var(--cyan);border-radius:999px}.mini-kpis{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-bottom:18px;align-items:stretch;width:100%}.mini-card{border:1px solid var(--border);border-radius:16px;padding:18px 20px;background:#fff;min-height:128px;height:128px;display:grid;grid-template-rows:28px 1fr 20px;align-items:center;overflow:visible}.mini-title{font-weight:700;color:#6b7280;font-size:13px;margin:0;display:flex;align-items:center;justify-content:space-between;gap:8px;min-width:0}.info-tip{position:relative;display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;background:var(--cyan2);color:var(--cyan);font-weight:900;font-size:14px;cursor:help;flex:0 0 auto}.info-tip::after{content:attr(data-tip);display:none;position:absolute;right:-8px;top:30px;width:300px;background:#1f3f68;color:#fff;border-radius:12px;padding:12px 14px;font-size:13px;line-height:1.45;font-weight:600;box-shadow:0 10px 24px rgba(0,0,0,.18);z-index:80;white-space:normal}.info-tip::before{content:"";display:none;position:absolute;right:3px;top:23px;border-left:8px solid transparent;border-right:8px solid transparent;border-bottom:8px solid #1f3f68;z-index:81}.info-tip:hover::after,.info-tip:hover::before{display:block}.yesno{display:grid;grid-template-columns:1fr 1fr;gap:18px;align-items:end;width:100%;grid-row:2 / 4}.yesno>div{display:grid;grid-template-rows:20px 42px;align-items:end}.big-cyan{font-size:34px;color:var(--cyan);font-weight:800;line-height:1}.mini-card>.big-cyan{grid-row:2;align-self:end}.mini-card>.small-muted{grid-row:3;align-self:end;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.note-row{display:flex;justify-content:space-between;font-size:14px;margin:14px 0 5px}.toast{position:fixed;right:22px;bottom:22px;background:#122235;color:#fff;padding:12px 16px;border-radius:12px;display:none;z-index:99}.return{cursor:pointer}.home-cover{min-height:720px;background:var(--cyan);border-radius:0;padding:38px 34px;color:white}.brand{font-size:34px;font-weight:800;letter-spacing:-1px}.home-title{font-size:78px;line-height:1.02;font-weight:300;margin:70px 0 55px;color:#cff6ff}.home-title b{font-weight:700;color:#fff}.home-cards{display:grid;grid-template-columns:repeat(4,minmax(210px,1fr));gap:30px;max-width:1180px}.home-card{background:#fff;color:#202124;border-radius:22px;min-height:235px;padding:52px 26px 24px;cursor:pointer;box-shadow:0 10px 24px rgba(0,0,0,.08);transition:.18s}.home-card:hover{transform:translateY(-3px)}.home-card .icon-badge{width:64px;height:64px;font-size:28px;margin-bottom:26px}.home-card h2{font-size:25px;line-height:1.05;margin:0}.home-card p{margin:12px 0 0;font-size:15px;line-height:1.35;color:#5f6f83;max-width:210px}.home-footer{font-size:34px;font-weight:800;margin-top:55px}.csvbox{display:none;width:100%;height:180px;margin-top:10px;border:1px solid var(--border);border-radius:10px;padding:10px;font-family:Consolas,monospace;font-size:12px}.tag-scroll{max-height:680px;overflow-y:auto;padding-right:4px;scrollbar-width:thin;scrollbar-color:var(--cyan) #e5e9ef}.tag-scroll::-webkit-scrollbar{width:6px}.tag-scroll::-webkit-scrollbar-track{background:#e5e9ef;border-radius:999px}.tag-scroll::-webkit-scrollbar-thumb{background:var(--cyan);border-radius:999px}
@media(max-width:1100px){.filters,.filters.three,.chart-grid-2,.motivos-grid,.quality-grid,.geral-top,.ra-cards,.ra-cards-all{grid-template-columns:1fr}.main{margin-left:0;width:100%}.side{display:none}.hero h1{font-size:24px}.kpi{width:100%}}
</style>
</head>
<body>
<div class="app">
  <aside class="side">
    <div class="nav active" onclick="showPage('home',this)" title="Capa"><i data-lucide="home"></i></div>
    <div class="nav" onclick="showPage('geral',this)" title="Visão Geral"><i data-lucide="chart-line"></i></div>
    <div class="nav green" onclick="showPage('depto',this)" title="Participação das Áreas nas Ocorrências"><i data-lucide="building-2"></i></div>
    <div class="nav" onclick="showPage('motivos',this)" title="Motivos & Ocorrências"><i data-lucide="clipboard-list"></i></div>
    <div class="nav" onclick="showPage('qualidade',this)" title="Qualidade & Status"><i data-lucide="headphones"></i></div>
  </aside>
  <main class="main">
    <section id="page-home" class="page active">
      <div class="home-cover">
        <div class="brand">●digisac</div>
        <div class="home-title">Relatório de<br><b>Reclamação</b></div>
        <div class="home-subtitle">Painel executivo para acompanhar reclamações, canais de entrada, departamentos responsáveis, motivos recorrentes, status de tratativa, qualidade do atendimento e base analítica do período.</div>
        <div class="home-cards">
          <div class="home-card" onclick="showPage('geral')">
            <span class="icon-badge"><i data-lucide="chart-line"></i></span>
            <h2>Visão Geral</h2>
            <p>KPIs, evolução mensal e canais de entrada.</p>
          </div>
          <div class="home-card" onclick="showPage('depto')">
            <span class="icon-badge"><i data-lucide="building-2"></i></span>
            <h2>Participação das<br>Áreas</h2>
            <p>Distribuição das ocorrências por área responsável.</p>
          </div>
          <div class="home-card" onclick="showPage('motivos')">
            <span class="icon-badge"><i data-lucide="clipboard-list"></i></span>
            <h2>Motivos & Ocorrências</h2>
            <p>Principais motivos, padrões e base analítica.</p>
          </div>
          <div class="home-card" onclick="showPage('qualidade')">
            <span class="icon-badge"><i data-lucide="headphones"></i></span>
            <h2>Qualidade & Status</h2>
            <p>Status das tratativas, satisfação e avaliações.</p>
          </div>
        </div>
        <div class="home-footer">ikatec</div>
        <div class="home-updated">Atualizado: <span id="homeAtualizado"></span></div>
      </div>
    </section>
    <section id="page-geral" class="page">
      <div class="hero"><div class="hero-title"><div class="h-left"><span class="icon-badge"><i data-lucide="chart-line"></i></span><h1>Visão Geral</h1></div><div class="return" title="Limpar filtros" onclick="resetFiltrosPagina()"><i data-lucide="rotate-ccw"></i></div></div><div class="filters three"><div class="fg"><label>Ano</label><select id="gAno" multiple></select></div><div class="fg"><label>Mês</label><select id="gMes" multiple></select></div><div class="fg"><label>Departamento</label><select id="gDepto" multiple></select></div></div></div>
      <div class="card round">
        <div class="ra-cards-all">
          <div class="ra-card ra-card-kpi">
            <div class="ra-label">Total de Reclamações</div>
            <div class="ra-value" id="kTotal" style="color:#071f3f">0</div>
          </div>
          <div class="ra-card ra-card-kpi">
            <div class="ra-label">Canal com Mais Reclamações</div>
            <div class="ra-value" id="kCanal" style="font-size:18px;color:#071f3f">-</div>
            <div class="ra-sub" id="kCanalQtd">0 reclamações</div>
          </div>
          <div class="ra-card destaque">
            <div class="ra-label-row"><span class="ra-label">Reputação Reclame Aqui</span><span class="info-tip" data-tip="Reputação pública da empresa no Reclame Aqui. Dados coletados automaticamente do site reclameaqui.com.br e podem não refletir atualizações em tempo real.">i</span></div>
            <div class="ra-value cyan" id="raReputacao">-</div>
            <div class="ra-sub"><b id="raNotaRep">-</b>/10 • últimos 6 meses</div>
          </div>
          <div class="ra-card">
            <div class="ra-label-row"><span class="ra-label">Qtd. Reclamações</span><span class="info-tip" data-tip="Total de reclamações recebidas pela empresa no Reclame Aqui no período indicado. Fonte: reclameaqui.com.br.">i</span></div>
            <div class="ra-value" id="raQtd">-</div>
            <div class="ra-sub">Período: <span id="raPeriodo">-</span></div>
          </div>
          <div class="ra-card">
            <div class="ra-label-row"><span class="ra-label">Respondidas</span><span class="info-tip" data-tip="Percentual de reclamações que receberam resposta da empresa no Reclame Aqui. Fonte: reclameaqui.com.br.">i</span></div>
            <div class="ra-value cyan" id="raRespondidas">-</div>
            <div class="ra-sub">Não respondidas: <b id="raNaoResp">-</b></div>
          </div>
          <div class="ra-card">
            <div class="ra-label-row"><span class="ra-label">Avaliadas</span><span class="info-tip" data-tip="Quantidade de reclamações avaliadas pelo consumidor após a resposta da empresa. A nota reflete a satisfação média. Fonte: reclameaqui.com.br.">i</span></div>
            <div class="ra-value" id="raAvaliadas">-</div>
            <div class="ra-sub">Nota consumidor: <b id="raNotaConsumidor">-</b></div>
          </div>
          <div class="ra-card">
            <div class="ra-label-row"><span class="ra-label">Voltariam a fazer negócio</span><span class="info-tip" data-tip="Percentual de clientes que avaliaram a reclamação e indicaram que voltariam a fazer negócio com a empresa. Fonte: reclameaqui.com.br.">i</span></div>
            <div class="ra-value cyan" id="raVoltariam">-</div>
            <div class="ra-sub">Clientes que avaliaram</div>
          </div>
          <div class="ra-card">
            <div class="ra-label-row"><span class="ra-label">Índice de solução</span><span class="info-tip" data-tip="Percentual de reclamações marcadas como resolvidas pelo consumidor no Reclame Aqui. Fonte: reclameaqui.com.br.">i</span></div>
            <div class="ra-value cyan" id="raSolucao">-</div>
            <div class="ra-sub">Reclamações resolvidas</div>
          </div>
        </div>
      </div>
      <div class="card round"><div class="chart-grid-2"><div class="chart-box"><h3>Reclamações por Mês e Ano</h3><div id="chartMesAno" class="plot"></div></div><div class="chart-box"><h3>Reclamações Instagram</h3><div id="chartInstagram" class="plot"></div></div><div class="chart-box"><h3>Reclamações Reclame Aqui</h3><div id="chartRA" class="plot"></div></div><div class="chart-box"><h3>Reclamações Google</h3><div id="chartGoogle" class="plot"></div></div></div></div>
    </section>

    <section id="page-depto" class="page">
      <div class="hero"><div class="hero-title"><div class="h-left"><span class="icon-badge"><i data-lucide="building-2"></i></span><h1>Participação das Áreas nas Ocorrências</h1></div><div class="return" title="Limpar filtros" onclick="resetFiltrosPagina()"><i data-lucide="rotate-ccw"></i></div></div><div class="filters three"><div class="fg"><label>Ano</label><select id="dAno" multiple></select></div><div class="fg"><label>Mês</label><select id="dMes" multiple></select></div><div class="fg"><label>Canal de Reclamação</label><select id="dCanal" multiple></select></div></div></div>
      <div class="card round"><h3 class="section-title">Participação das Áreas nas Ocorrências</h3><div id="chartDeptoAno" class="dept-main"></div></div>
      <div class="card round"><h3 class="section-title" id="tituloDeptoMes">Reclamações 2026</h3><div id="chartDeptoMes" class="month-dept"></div></div>
      <div class="card round summary"><h3>Resumo Executivo</h3><div id="resumoDepto"></div></div>
    </section>

    <section id="page-motivos" class="page">
      <div class="hero"><div class="hero-title"><div class="h-left"><span class="icon-badge"><i data-lucide="clipboard-list"></i></span><h1>Motivos & Ocorrências</h1></div><div class="return" title="Limpar filtros" onclick="resetFiltrosPagina()"><i data-lucide="rotate-ccw"></i></div></div><div class="filters"><div class="fg"><label>Ano</label><select id="mAno" multiple></select></div><div class="fg"><label>Mês</label><select id="mMes" multiple></select></div><div class="fg"><label>Departamento</label><select id="mDepto" multiple></select></div><div class="fg"><label>Canal Reclamação</label><select id="mCanal" multiple></select></div></div></div>
      <h2 style="margin:0 0 4px;font-size:24px">Motivos das Participação das Áreas nas Ocorrências</h2><div class="small-muted" id="periodoMotivos"></div><br>
      <div class="motivos-grid" id="motivoCards"></div>
      <div class="card round base-analitica"><div class="analytics-head"><div><h2 style="margin:0">Base Analítica</h2><div class="small-muted" id="periodoBase"></div></div><div class="btns"><button class="btn primary" onclick="baixarCSV()">Baixar CSV filtrado</button><button class="btn" onclick="copiarCSV()">Copiar CSV</button><button class="btn" onclick="mostrarCSV()">Mostrar CSV</button></div></div><textarea id="csvText" class="csvbox"></textarea><br><div class="table-wrap" id="tabelaMotivos"></div></div>
    </section>

    <section id="page-qualidade" class="page">
      <div class="hero"><div class="hero-title"><div class="h-left"><span class="icon-badge"><i data-lucide="headphones"></i></span><h1>Qualidade & Status</h1></div><div class="return" title="Limpar filtros" onclick="resetFiltrosPagina()"><i data-lucide="rotate-ccw"></i></div></div><div class="filters"><div class="fg"><label>Ano</label><select id="qAno" multiple></select></div><div class="fg"><label>Mês</label><select id="qMes" multiple></select></div><div class="fg"><label>Departamento</label><select id="qDepto" multiple></select></div><div class="fg"><label>Canal Reclamação</label><select id="qCanal" multiple></select></div></div></div>
      <div class="quality-grid"><div><h2>Fases Atuais</h2><div class="small-muted" id="periodoQ1"></div><div class="tag-scroll"><div id="statusCards"></div></div></div><div><h2>Etiquetas por Fase Atual</h2><div class="small-muted" id="periodoQ2"></div><div class="tag-scroll"><div id="tagCards"></div></div></div><div><h2>Satisfação do Cliente</h2><div class="small-muted" id="periodoQ4"></div><div class="mini-kpis"><div class="mini-card"><div class="mini-title"><span>Problema Resolvido</span><span class="info-tip" data-tip="Conta as respostas da coluna 'Problema resolvido?'. Sim = reclamações marcadas como resolvidas. Não = reclamações marcadas como não resolvidas. Respostas vazias ou diferentes de Sim/Não não entram nesse card.">i</span></div><div class="yesno"><div>Sim<br><span class="big-cyan" id="qSim">0</span></div><div>Não<br><span class="big-cyan" id="qNao">0</span></div></div></div><div class="mini-card"><div class="mini-title"><span>% Resolvidas</span><span class="info-tip" data-tip="Cálculo: quantidade de respostas 'Sim' em Problema resolvido ÷ total de respostas Sim + Não em Problema resolvido × 100.">i</span></div><div class="big-cyan" id="qPercResolvido">0%</div><div class="small-muted" id="qTxtResolvido"></div></div><div class="mini-card"><div class="mini-title"><span>Voltaria a Fazer Negócio</span><span class="info-tip" data-tip="Conta as respostas da coluna 'Voltaria a fazer negócio?'. Sim = clientes que voltariam a fazer negócio. Não = clientes que não voltariam. Respostas vazias ou diferentes de Sim/Não não entram nesse card.">i</span></div><div class="yesno"><div>Sim<br><span class="big-cyan" id="vSim">0</span></div><div>Não<br><span class="big-cyan" id="vNao">0</span></div></div></div><div class="mini-card"><div class="mini-title"><span>% Voltaria a Fazer Negócio</span><span class="info-tip" data-tip="Cálculo: quantidade de respostas 'Sim' em Voltaria a fazer negócio ÷ total de respostas Sim + Não em Voltaria a fazer negócio × 100.">i</span></div><div class="big-cyan" id="vPerc">0%</div><div class="small-muted" id="vTxt"></div></div></div><div class="card round"><h2>Notas das Avaliações</h2><div class="small-muted" id="periodoQ3"></div><br><div id="notasBox"></div></div></div></div>
    </section>
  </main>
</div>
<div id="toast" class="toast"></div>
<script>
const DATA=__PAYLOAD_JSON__; const MESES=DATA.meses;
const C1="#1d8ff2", C2="#182f9f", CY="#19b5cf", MUT="#71829a";
function id(x){return document.getElementById(x)}
function toast(t){id('toast').textContent=t;id('toast').style.display='block';setTimeout(()=>id('toast').style.display='none',2600)}
function showPage(p,el){document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));id('page-'+p).classList.add('active');document.querySelector('.app').classList.toggle('home-mode',p==='home');document.querySelectorAll('.nav').forEach(x=>x.classList.remove('active'));let navMap={home:0,geral:1,depto:2,motivos:3,qualidade:4};let navs=document.querySelectorAll('.nav');if(el)el.classList.add('active');else if(navs[navMap[p]])navs[navMap[p]].classList.add('active');renderAll();if(window.lucide){lucide.createIcons();}}
function fill(sel,vals,all='Todos'){sel.innerHTML=`<option value="Todos">${all}</option>`; vals.forEach(v=>sel.innerHTML+=`<option value="${String(v).replace(/"/g,'&quot;')}">${v}</option>`)}
function fillMes(sel){sel.innerHTML='<option value="Todos">Todos</option>'; MESES.forEach(m=>sel.innerHTML+=`<option value="${m.num}">${m.nome}</option>`)}
function getVals(sel){let vals=[...sel.selectedOptions].map(o=>o.value);if(!vals.length||vals.includes('Todos'))return ['Todos'];return vals}
function isAll(vals){return !vals||vals.includes('Todos')}
function matchVal(value, vals){return isAll(vals)||vals.map(String).includes(String(value))}
function num(n){return new Intl.NumberFormat('pt-BR').format(n||0)} function dec(n){return new Intl.NumberFormat('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1}).format(n||0)}
function uniq(a){return [...new Set(a.filter(x=>x!==null&&x!==undefined&&String(x).trim()!==''))]}
function group(rows,field){let m=new Map();rows.forEach(r=>{let k=r[field]||'Não informado';m.set(k,(m.get(k)||0)+1)});return [...m].map(([key,value])=>({key,value})).sort((a,b)=>b.value-a.value)}
function monthLong(v){let m=MESES.find(x=>String(x.num)===String(v));return m?m.longo:''}
function filtros(prefix){return {ano:getVals(id(prefix+'Ano')),mes:getVals(id(prefix+'Mes')),depto:id(prefix+'Depto')?getVals(id(prefix+'Depto')):['Todos'],canal:id(prefix+'Canal')?getVals(id(prefix+'Canal')):['Todos']}}

let syncingFilters=false;
function filtroTipoPorId(selectId){
  if(selectId.endsWith('Ano')) return 'Ano';
  if(selectId.endsWith('Mes')) return 'Mes';
  if(selectId.endsWith('Depto')) return 'Depto';
  if(selectId.endsWith('Canal')) return 'Canal';
  return null;
}
function idsPorTipoFiltro(tipo){
  const mapa={
    Ano:['gAno','dAno','mAno','qAno'],
    Mes:['gMes','dMes','mMes','qMes'],
    Depto:['gDepto','mDepto','qDepto'],
    Canal:['dCanal','mCanal','qCanal']
  };
  return mapa[tipo]||[];
}
function aplicarValoresFiltro(sel, valores){
  [...sel.options].forEach(o=>o.selected=valores.includes(o.value));
  if(![...sel.selectedOptions].length && sel.options[0]) sel.options[0].selected=true;
}
function sincronizarFiltrosGlobais(sel){
  if(syncingFilters || !sel || !sel.id) return;
  const tipo=filtroTipoPorId(sel.id);
  if(!tipo) return;
  syncingFilters=true;
  const valores=getVals(sel);
  idsPorTipoFiltro(tipo).forEach(selectId=>{
    const alvo=id(selectId);
    if(alvo && alvo!==sel) aplicarValoresFiltro(alvo,valores);
  });
  syncingFilters=false;
  refreshMultiSelectDisplays();
}
function atualizarFiltroERenderizar(sel){
  sincronizarFiltrosGlobais(sel);
  refreshMultiSelectDisplays();
  renderAll();
}
function filterRows(rows,f){return rows.filter(r=>matchVal(r.Ano,f.ano)&&matchVal(r.MesNum,f.mes)&&matchVal(r['Departamento responsável'],f.depto)&&matchVal(r['Canal da reclamação'],f.canal))}
function selectedYears(f){return isAll(f.ano)?DATA.anos.map(String):f.ano.map(String)}
function selectedMonths(f, rows=null){if(!isAll(f.mes))return MESES.filter(m=>f.mes.map(String).includes(String(m.num)));let base=rows||DATA.principal;let nums=uniq(base.map(r=>Number(r.MesNum))).sort((a,b)=>a-b);return MESES.filter(m=>nums.includes(Number(m.num)))}
function periodoLabel(f){let anos=selectedYears(f).sort();let meses=selectedMonths(f);if(!anos.length)anos=DATA.anos.map(String);if(!meses.length)meses=MESES;let minAno=anos[0], maxAno=anos[anos.length-1];if(meses.length===1&&anos.length>1)return `${meses[0].longo}/${minAno} a ${meses[0].longo}/${maxAno}`;if(meses.length===1&&anos.length===1)return `${meses[0].longo}/${minAno}`;let mIni=meses[0].longo,mFim=meses[meses.length-1].longo;if(anos.length===1)return `${mIni} a ${mFim}/${minAno}`;return `${mIni}/${minAno} a ${mFim}/${maxAno}`}
function layout(extra={}){return {paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',margin:{l:10,r:45,t:42,b:34},font:{family:'Segoe UI,Arial',size:12,color:'#111827'},xaxis:{showgrid:false,zeroline:false,automargin:true},yaxis:{showgrid:false,zeroline:false,automargin:true,rangemode:'tozero',showticklabels:false,ticks:''},legend:{orientation:'h',x:0,y:1.22},...extra}}
function lineMonth(div,rows,canal=null,f=null){let base=canal?rows.filter(r=>String(r['Canal da reclamação']).toLowerCase().includes(canal)):rows;let years=f?selectedYears(f):uniq(base.map(r=>r.Ano)).sort();if(!years.length)years=DATA.anos.map(String);let colors=[C1,C2,'#59bfd7','#209fb8'],allVals=[];let traces=years.map((y,i)=>{let vals=MESES.map(m=>base.filter(r=>String(r.Ano)===String(y)&&Number(r.MesNum)===Number(m.num)).length);allVals.push(...vals);return {x:MESES.map(m=>m.nome),y:vals,type:'scatter',mode:'lines+markers+text',name:String(y),text:vals.map(v=>String(v||0)),textposition:'top center',cliponaxis:false,line:{color:colors[i%colors.length],width:3},marker:{color:colors[i%colors.length],size:8},hovertemplate:'%{x}<br>Qtd.: %{y}<extra></extra>'}});let mx=Math.max(...allVals,1);Plotly.newPlot(div,traces,layout({margin:{l:10,r:50,t:42,b:30},yaxis:{showgrid:false,zeroline:false,showticklabels:false,ticks:'',range:[0,Math.ceil(mx*1.55)]}}),{responsive:true,displayModeBar:false})}
function renderReclameAqui(){
  const ra = DATA.reclame_aqui || {};
  id('raReputacao').textContent = ra.reputacao || '-';
  id('raNotaRep').textContent = ra.nota_media_reputacao || '-';
  id('raQtd').textContent = ra.qtd_reclamacoes || '-';
  id('raPeriodo').textContent = ra.periodo || '-';
  id('raRespondidas').textContent = ra.reclamacoes_respondidas_pct || '-';
  id('raNaoResp').textContent = ra.nao_respondidas || '-';
  id('raAvaliadas').textContent = ra.avaliadas || '-';
  id('raNotaConsumidor').textContent = ra.nota_consumidor || '-';
  id('raVoltariam').textContent = ra.voltariam_fazer_negocio_pct || '-';
  id('raSolucao').textContent = ra.indice_solucao_pct || '-';
}
function renderGeral(){renderReclameAqui();let f=filtros('g');let rows=filterRows(DATA.principal,f);id('kTotal').textContent=num(rows.length);let gc=group(rows,'Canal da reclamação')[0]||{key:'-',value:0};id('kCanal').textContent=gc.key;id('kCanalQtd').textContent=num(gc.value)+' reclamações';lineMonth('chartMesAno',rows,null,f);lineMonth('chartInstagram',rows,'instagram',f);lineMonth('chartRA',rows,'reclame',f);lineMonth('chartGoogle',rows,'google',f)}
function labelMesesSelecionados(f){
  let meses = selectedMonths(f);

  if(!meses.length){
    return "Todos";
  }

  if(meses.length === 1){
    return meses[0].nome;
  }

  return meses[0].nome + " a " + meses[meses.length - 1].nome;
}


function arredondarBarrasPlotly(divId){
  const alvo = document.getElementById(divId);
  if(!alvo) return;
  const paths = alvo.querySelectorAll('.barlayer .trace.bars path');
  paths.forEach(path=>{
    const d = path.getAttribute('d') || '';
    const nums = d.match(/-?\d+(?:\.\d+)?/g);
    if(!nums || nums.length < 4) return;

    // Formato padrão do Plotly em gráfico de barra vertical:
    // M x0,yBase V yTopo H x1 V yBase Z
    const x0 = parseFloat(nums[0]);
    const yBase = parseFloat(nums[1]);
    const yTopo = parseFloat(nums[2]);
    const x1 = parseFloat(nums[3]);

    if([x0,yBase,yTopo,x1].some(v=>Number.isNaN(v))) return;

    const largura = Math.abs(x1 - x0);
    const altura = Math.abs(yBase - yTopo);
    if(largura < 4 || altura < 4) return;

    const r = Math.min(9, largura / 2, altura / 2);
    const esquerda = Math.min(x0, x1);
    const direita = Math.max(x0, x1);
    const topo = Math.min(yBase, yTopo);
    const base = Math.max(yBase, yTopo);

    const novoD = `M${esquerda},${base} V${topo + r} Q${esquerda},${topo} ${esquerda + r},${topo} H${direita - r} Q${direita},${topo} ${direita},${topo + r} V${base} Z`;
    path.setAttribute('d', novoD);
  });
}

function renderDepto(){
  let f = {
    ano:getVals(id('dAno')),
    mes:getVals(id('dMes')),
    depto:['Todos'],
    canal:getVals(id('dCanal'))
  };

  // ===============================
  // GRÁFICO 1: respeita filtros
  // ===============================
  let rows = filterRows(DATA.principal,f);
  let periodoMes = labelMesesSelecionados(f);

  let depts = group(rows,'Departamento responsável').map(x=>x.key);

  if(!depts.length){
    depts = uniq(DATA.principal.map(r=>r['Departamento responsável'])).sort();
  }

  let anosSel = selectedYears(f);

  let traces = anosSel.map((ano,i)=>{
    let vals = depts.map(d =>
      DATA.principal.filter(r =>
        String(r.Ano) === String(ano) &&
        matchVal(r.MesNum,f.mes) &&
        matchVal(r['Canal da reclamação'],f.canal) &&
        r['Departamento responsável'] === d
      ).length
    );

    return {
      x:depts,
      y:vals,
      type:'bar',
      name: ano + " | " + periodoMes,
      text:vals.map(v=>String(v||0)),
      textposition:'outside',
      cliponaxis:false,
      marker:{
        color:[C1,C2,'#59bfd7','#209fb8'][i%4]
      }
    };
  });

  let max1 = Math.max(1,...traces.flatMap(t=>t.y));

  Plotly.newPlot(
    'chartDeptoAno',
    traces,
    layout({
      barmode:'group',
      height:170,
      margin:{l:10,r:35,t:42,b:52},
      showlegend:true,
      yaxis:{
        showgrid:false,
        zeroline:false,
        showticklabels:false,
        ticks:'',
        range:[0,Math.ceil(max1*1.7)]
      }
    }),
    {responsive:true,displayModeBar:false}
  ).then(()=>arredondarBarrasPlotly('chartDeptoAno'));

  // ===============================
  // GRÁFICO 2: TRAVADO EM 2026
  // Não respeita filtro de Ano nem filtro de Mês.
  // Ele serve para mostrar a evolução mensal de 2026.
  // O filtro de Canal continua funcionando.
  // ===============================
  const ANO_EVOLUCAO = 2026;

  id('tituloDeptoMes').textContent = 'Evolução Mensal das Ocorrências por Área - 2026';

  let rows2026 = DATA.principal.filter(r =>
    String(r.Ano) === String(ANO_EVOLUCAO) &&
    matchVal(r['Canal da reclamação'],f.canal)
  );

  let depts2026 = group(rows2026,'Departamento responsável').map(x=>x.key);

  if(!depts2026.length){
    depts2026 = uniq(
      DATA.principal
        .filter(r => String(r.Ano) === String(ANO_EVOLUCAO))
        .map(r => r['Departamento responsável'])
    ).sort();
  }

  let mesesMostrar = MESES.filter(m =>
    rows2026.some(r => Number(r.MesNum) === Number(m.num))
  );

  if(!mesesMostrar.length){
    mesesMostrar = MESES;
  }

  let tracesMes = mesesMostrar.map((m,i)=>{
    let vals = depts2026.map(d =>
      DATA.principal.filter(r =>
        String(r.Ano) === String(ANO_EVOLUCAO) &&
        Number(r.MesNum) === Number(m.num) &&
        matchVal(r['Canal da reclamação'],f.canal) &&
        r['Departamento responsável'] === d
      ).length
    );

    return {
      x:depts2026,
      y:vals,
      type:'bar',
      name:m.nome,
      text:vals.map(v=>String(v||0)),
      textposition:'outside',
      cliponaxis:false,
      hovertemplate:'<b>%{x}</b><br>Mês: <b>'+m.nome+'</b><br>Reclamações: <b>%{y}</b><extra></extra>',
      marker:{
        color:['#7dcce2','#59bfd7','#37abc3','#209fb8','#188ea7','#69cfe1'][i%6]
      }
    };
  });

  let max2 = Math.max(1,...tracesMes.flatMap(t=>t.y));

  Plotly.newPlot(
    'chartDeptoMes',
    tracesMes,
    layout({
      barmode:'group',
      height:185,
      margin:{l:10,r:35,t:42,b:52},
      showlegend:true,
      yaxis:{
        showgrid:false,
        zeroline:false,
        showticklabels:false,
        ticks:'',
        range:[0,Math.ceil(max2*1.7)]
      }
    }),
    {responsive:true,displayModeBar:false}
  ).then(()=>arredondarBarrasPlotly('chartDeptoMes'));

  let total = rows.length;
  let topD = group(rows,'Departamento responsável')[0] || {key:'-',value:0};
  let topC = group(rows,'Canal da reclamação')[0] || {key:'-',value:0};
  let topM = group(rows,'Motivo Final')[0] || {key:'-',value:0};

  id('resumoDepto').innerHTML =
    `No período <b>${periodoLabel(f)}</b>, foram registradas <b>${num(total)}</b> ocorrências. A área com maior participação nas ocorrências foi <b>${topD.key}</b>, com <b>${num(topD.value)}</b> registros associados. O canal com maior incidência foi <b>${topC.key}</b>, totalizando <b>${num(topC.value)}</b> ocorrências. O motivo mais recorrente no período foi <b>${topM.key}</b>, representando <b>${num(topM.value)}</b> registros.`;
}

function motivoRows(){let f=filtros('m');return DATA.motivos.filter(r=>matchVal(r.Ano,f.ano)&&matchVal(r.MesNum,f.mes)&&(isAll(f.depto)||matchVal(r['Departamento do motivo'],f.depto)||matchVal(r['Departamento responsável'],f.depto))&&matchVal(r['Canal da reclamação'],f.canal))}
function renderMotivos(){let f=filtros('m');let rows=motivoRows();id('periodoMotivos').textContent=periodoLabel(f);id('periodoBase').textContent=periodoLabel(f);let depts=group(rows,'Departamento do motivo');let html='';if(!depts.length)html='<div class="card">Sem motivos para os filtros selecionados.</div>';depts.forEach(d=>{let rs=rows.filter(r=>r['Departamento do motivo']===d.key);let g=group(rs,'Motivo').slice(0,3);html+=`<div class="motivo-card"><div class="motivo-head"><div><h3>${d.key}</h3><div class="small-muted">Top 3 motivos do departamento</div></div><div class="motivo-total">${num(d.value)}<div class="small-muted">reclamações</div></div></div><table class="mini-table"><thead><tr><th>MOTIVO</th><th>QTD</th><th>% DEPTO</th></tr></thead><tbody>${g.map(x=>`<tr><td><b>${x.key}</b></td><td>${num(x.value)}</td><td>${dec(x.value/d.value*100)}%</td></tr>`).join('')}</tbody></table></div>`});id('motivoCards').innerHTML=html;renderTabelaAnalitica()}
function getAnaliticaFiltrada(){let f=filtros('m');return filterRows(DATA.principal,f)}
function renderTabelaAnalitica(){let rows=getAnaliticaFiltrada();let cols=['Título','Data da manifestação','Finalizado em','Departamento responsável','Motivo Final','Outro Motivo','Canal da reclamação','Fase atual','Etiquetas','Problema resolvido?','Voltaria a fazer negócio?','Nota do atendimento','URL','Observações'];let html='<table class="detail-table"><thead><tr>'+cols.map(c=>`<th>${c}</th>`).join('')+'</tr></thead><tbody>';if(!rows.length)html+='<tr><td colspan="14">Sem registros para os filtros selecionados.</td></tr>';rows.slice(0,1000).forEach(r=>{html+='<tr>'+cols.map(c=>`<td>${c==='Fase atual'?'<span class="pill">'+(r[c]||'')+'</span>':(r[c]??'')}</td>`).join('')+'</tr>'});html+='</tbody></table>';id('tabelaMotivos').innerHTML=html;prepararCSV()}
function csvEscape(v){v=(v??'').toString().replace(/\r?\n/g,' ');return '"'+v.replace(/"/g,'""')+'"'}
function prepararCSV(){let rows=getAnaliticaFiltrada();let cols=['Título','Data da manifestação','Finalizado em','Departamento responsável','Motivo Final','Canal da reclamação','Fase atual','Etiquetas','Problema resolvido?','Voltaria a fazer negócio?','Nota do atendimento','URL','Observações'];let csv=cols.join(';')+'\n'+rows.map(r=>cols.map(c=>csvEscape(r[c])).join(';')).join('\n');id('csvText').value=csv;return csv}
function baixarCSV(){let csv=prepararCSV();let blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8;'});let url=URL.createObjectURL(blob);let a=document.createElement('a');a.href=url;a.download='Base_Analitica_Reclamacao_Filtrada.csv';document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);toast('CSV filtrado gerado. Se estiver dentro do Power BI e não baixar, use Copiar CSV.')} 
function copiarCSV(){let csv=prepararCSV();navigator.clipboard?.writeText(csv).then(()=>toast('CSV copiado. Cole no Excel/Bloco de notas.')).catch(()=>{id('csvText').style.display='block';id('csvText').select();document.execCommand('copy');toast('CSV selecionado/copiadado.')})}
function mostrarCSV(){prepararCSV();id('csvText').style.display=id('csvText').style.display==='block'?'none':'block'}
function normalResp(v){let t=String(v||'').trim().toLowerCase();if(t==='sim')return 'sim';if(t==='não'||t==='nao')return 'nao';return 'outro'}
function renderQualidade(){let f=filtros('q');let rows=filterRows(DATA.principal,f);id('periodoQ1').textContent=periodoLabel(f);id('periodoQ2').textContent=periodoLabel(f);id('periodoQ3').textContent=periodoLabel(f);id('periodoQ4').textContent=periodoLabel(f);let fases=group(rows,'Fase atual');id('statusCards').innerHTML=fases.map(x=>`<div class="status-card"><span class="status-num">${num(x.value)}</span><h3>${x.key}</h3><div class="small-muted">Fase atual da reclamação</div><div class="bar-bg"><div class="bar" style="width:${rows.length?x.value/rows.length*100:0}%"></div></div><div class="yesno"><span class="small-muted">% do total de reclamações</span><b>${dec(rows.length?x.value/rows.length*100:0)}%</b></div></div>`).join('')||'<div class="status-card">Sem dados.</div>';let tagHtml='';fases.forEach(fa=>{let ets=DATA.etiquetas.filter(e=>rows.some(r=>String(r.Ano)===String(e.Ano)&&String(r.MesNum)===String(e.MesNum)&&r['Departamento responsável']===e['Departamento responsável']&&r['Fase atual']===fa.key)&&e['Fase atual']===fa.key);let g=group(ets,'Etiqueta').slice(0,5);tagHtml+=`<div class="motivo-card" style="margin-bottom:16px"><div class="motivo-head"><div><h3>${fa.key}</h3><div class="small-muted">Top 5 etiquetas da fase</div></div><div class="motivo-total">${fa.value}<div class="small-muted">reclamações</div></div></div><table class="mini-table"><thead><tr><th>ETIQUETA</th><th>QTD</th><th>% FASE</th></tr></thead><tbody>${g.map(x=>`<tr><td><b>${x.key}</b></td><td>${x.value}</td><td>${dec(x.value/fa.value*100)}%</td></tr>`).join('')||'<tr><td colspan="3">Sem etiquetas</td></tr>'}</tbody></table></div>`});id('tagCards').innerHTML=tagHtml;let sim=rows.filter(r=>normalResp(r['Problema resolvido?'])==='sim').length,nao=rows.filter(r=>normalResp(r['Problema resolvido?'])==='nao').length;id('qSim').textContent=sim;id('qNao').textContent=nao;id('qPercResolvido').textContent=dec((sim+nao)?sim/(sim+nao)*100:0)+'%';id('qTxtResolvido').textContent=`${num(sim)} resolvidas • ${periodoLabel(f)}`;let vs=rows.filter(r=>normalResp(r['Voltaria a fazer negócio?'])==='sim').length,vn=rows.filter(r=>normalResp(r['Voltaria a fazer negócio?'])==='nao').length;id('vSim').textContent=vs;id('vNao').textContent=vn;id('vPerc').textContent=dec((vs+vn)?vs/(vs+vn)*100:0)+'%';id('vTxt').textContent=`${num(vs)} respostas positivas • ${periodoLabel(f)}`;let notas=rows.filter(r=>r['Nota do atendimento']!==null&&r['Nota do atendimento']!==undefined&&!isNaN(Number(r['Nota do atendimento'])));let gn=group(notas,'Nota do atendimento').sort((a,b)=>Number(b.key)-Number(a.key));id('notasBox').innerHTML=`<div>Total de avaliações com nota: <b>${num(notas.length)}</b></div>`+gn.map(x=>`<div class="note-row"><span>Nota ${x.key}</span><b>${x.value} avaliações · ${dec(notas.length?x.value/notas.length*100:0)}%</b></div><div class="bar-bg"><div class="bar" style="width:${notas.length?x.value/notas.length*100:0}%"></div></div>`).join('')}

function enhanceMultiSelects(){
  document.querySelectorAll('select[multiple]').forEach(sel=>{
    if(sel.dataset.enhanced==='1') return;
    sel.dataset.enhanced='1';
    const wrap=document.createElement('div');
    wrap.className='ms-wrap';
    const display=document.createElement('div');
    display.className='ms-display';
    const panel=document.createElement('div');
    panel.className='ms-panel';
    const actions=document.createElement('div');
    actions.className='ms-actions';
    const bAll=document.createElement('button'); bAll.type='button'; bAll.className='ms-btn'; bAll.textContent='Todos';
    const bClear=document.createElement('button'); bClear.type='button'; bClear.className='ms-btn'; bClear.textContent='Limpar';
    actions.appendChild(bAll); actions.appendChild(bClear); panel.appendChild(actions);
    function optionLabel(opt){
      const lab=document.createElement('label'); lab.className='ms-option';
      const cb=document.createElement('input'); cb.type='checkbox'; cb.value=opt.value;
      cb.checked=opt.selected || opt.value==='Todos';
      lab.appendChild(cb); lab.appendChild(document.createTextNode(opt.textContent));
      cb.addEventListener('change',()=>{
        if(cb.value==='Todos' && cb.checked){
          panel.querySelectorAll('input[type="checkbox"]').forEach(x=>{x.checked=(x.value==='Todos')});
        }else{
          const todos=panel.querySelector('input[value="Todos"]');
          if(todos) todos.checked=false;
          const checked=[...panel.querySelectorAll('input[type="checkbox"]')].filter(x=>x.checked && x.value!=='Todos');
          if(!checked.length && todos) todos.checked=true;
        }
        syncToSelect(); atualizarFiltroERenderizar(sel);
      });
      panel.appendChild(lab);
    }
    [...sel.options].forEach(optionLabel);
    function syncToSelect(){
      const checked=[...panel.querySelectorAll('input[type="checkbox"]')].filter(x=>x.checked).map(x=>x.value);
      [...sel.options].forEach(o=>o.selected=checked.includes(o.value));
      updateDisplay();
    }
    function updateFromSelect(){
      const selected=[...sel.selectedOptions].map(o=>o.value);
      panel.querySelectorAll('input[type="checkbox"]').forEach(cb=>cb.checked=selected.includes(cb.value));
      if(!selected.length){ const todos=panel.querySelector('input[value="Todos"]'); if(todos) todos.checked=true; }
      updateDisplay();
    }
    function updateDisplay(){
      const selected=[...sel.selectedOptions].map(o=>o.textContent);
      display.textContent=(!selected.length || selected.includes('Todos')) ? 'Todos' : selected.join(', ');
    }
    bAll.addEventListener('click',()=>{panel.querySelectorAll('input[type="checkbox"]').forEach(x=>x.checked=(x.value==='Todos'));syncToSelect();atualizarFiltroERenderizar(sel);});
    bClear.addEventListener('click',()=>{panel.querySelectorAll('input[type="checkbox"]').forEach(x=>x.checked=false);const todos=panel.querySelector('input[value="Todos"]');if(todos)todos.checked=true;syncToSelect();atualizarFiltroERenderizar(sel);});
    display.addEventListener('click',(e)=>{e.stopPropagation();document.querySelectorAll('.ms-wrap.open').forEach(w=>{if(w!==wrap)w.classList.remove('open')});wrap.classList.toggle('open');});
    panel.addEventListener('click',e=>e.stopPropagation());
    sel.parentNode.insertBefore(wrap, sel.nextSibling);
    wrap.appendChild(display); wrap.appendChild(panel);
    sel.addEventListener('change',updateFromSelect);
    syncToSelect();
  });
  document.addEventListener('click',()=>document.querySelectorAll('.ms-wrap.open').forEach(w=>w.classList.remove('open')),{once:false});
}
function refreshMultiSelectDisplays(){
  document.querySelectorAll('select[multiple]').forEach(sel=>{
    const wrap=sel.nextElementSibling;
    if(!wrap||!wrap.classList.contains('ms-wrap')) return;
    const selected=[...sel.selectedOptions].map(o=>o.value);
    wrap.querySelectorAll('input[type="checkbox"]').forEach(cb=>cb.checked=selected.includes(cb.value));
    if(!selected.length){ const todos=wrap.querySelector('input[value="Todos"]'); if(todos) todos.checked=true; }
    const labels=[...sel.selectedOptions].map(o=>o.textContent);
    const display=wrap.querySelector('.ms-display');
    if(display) display.textContent=(!labels.length || labels.includes('Todos')) ? 'Todos' : labels.join(', ');
  });
}

function renderAll(){if(id('page-home').classList.contains('active')){if(window.lucide){lucide.createIcons();}return;}if(id('page-geral').classList.contains('active'))renderGeral();if(id('page-depto').classList.contains('active'))renderDepto();if(id('page-motivos').classList.contains('active'))renderMotivos();if(id('page-qualidade').classList.contains('active'))renderQualidade();if(window.lucide){lucide.createIcons();}}
function resetFiltrosPagina(){let p=document.querySelector('.page.active').id.replace('page-','');if(p==='home')return;document.querySelectorAll('select').forEach(sel=>{[...sel.options].forEach(o=>o.selected=false);if(sel.options[0])sel.options[0].selected=true});refreshMultiSelectDisplays();renderAll();toast('Filtros limpos em todas as páginas.')}
function setSelectOnly(sel,val){[...sel.options].forEach(o=>o.selected=String(o.value)===String(val));refreshMultiSelectDisplays();}
function init(){document.querySelector('.app').classList.add('home-mode'); if(id('homeAtualizado')) id('homeAtualizado').textContent=DATA.atualizado;['g', 'm', 'q'].forEach(p=>{fill(id(p+'Ano'),DATA.anos);fillMes(id(p+'Mes'));fill(id(p+'Depto'),DATA.departamentos);if(id(p+'Canal'))fill(id(p+'Canal'),DATA.canais)});fill(id('dAno'),DATA.anos);fillMes(id('dMes'));fill(id('dCanal'),DATA.canais);document.querySelectorAll('select').forEach(s=>s.addEventListener('change',()=>atualizarFiltroERenderizar(s)));enhanceMultiSelects();if(DATA.anos.includes(2026)){['mAno', 'qAno'].forEach(x=>setSelectOnly(id(x),'2026'))}renderAll()}
init();
if(window.lucide){lucide.createIcons();}
</script>
</body>
</html>
'''

html = html.replace("__PAYLOAD_JSON__", payload_json)
ARQUIVO_SAIDA.write_text(html, encoding="utf-8")

print(f"HTML gerado com sucesso em: {ARQUIVO_SAIDA}")
print(f"Base analítica completa exportada em: {ARQUIVO_BASE_COMPLETA}")
print("Observação: dentro do Power BI alguns visuais HTML bloqueiam download. Nesse caso, use o botão 'Copiar CSV' ou 'Mostrar CSV'.")
