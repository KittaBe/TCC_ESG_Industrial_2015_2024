# ==============================================================
# ANÁLISE ESTATÍSTICA - IMPACTO ESG NAS EMPRESAS INDUSTRIAIS
# ==============================================================
# - Lê a base Excel gerada pelo pipeline
# - Correlação, boxplots e regressões OLS com erros robustos (HC3)
# - Exporta: Tabelas no Excel (com interpretação), resumos .txt e log
# ==============================================================

import os, io, shutil, time
from typing import Optional

import numpy as np
import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import seaborn as sns

# ------------------ CONFIG ------------------
BASE_PATH = r"C:\Users\Kitta\OneDrive\Desktop\TCC\Python\base_industrial_2019_2024.xlsx"
OUT_DIR   = r"C:\Users\Kitta\OneDrive\Desktop\TCC\Python"
OUT_XLSX  = os.path.join(OUT_DIR, "resultados_estatisticos.xlsx")
LOG_PATH  = os.path.join(OUT_DIR, "analise_log.txt")
FIG_DIR   = os.path.join(OUT_DIR, "outputs")
os.makedirs(FIG_DIR, exist_ok=True)

# ------------- Helpers: leitura segura -------------
def safe_read_excel(path, sheet_name=0, tries=3, wait_sec=1.0):
    """Tenta ler o Excel; se estiver bloqueado, lê de uma cópia temporária."""
    last_err = None
    for i in range(tries):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
        except PermissionError as e:
            last_err = e
            try:
                tmp = os.path.join(os.path.dirname(path), f"~tmp_read_{int(time.time())}_{i}.xlsx")
                shutil.copy2(path, tmp)
                df = pd.read_excel(tmp, sheet_name=sheet_name, engine="openpyxl")
                try:
                    os.remove(tmp)
                except Exception:
                    pass
                return df
            except PermissionError as e2:
                last_err = e2
                time.sleep(wait_sec)
    raise last_err

# ------------ Helpers: interpretação ------------
def estrelas(p):
    if p < 0.01: return "***"
    if p < 0.05: return "**"
    if p < 0.10: return "*"
    return ""

def nivel_sig(p):
    if p < 0.01: return "muito forte (p<0,01)"
    if p < 0.05: return "forte (p<0,05)"
    if p < 0.10: return "moderada (p<0,10)"
    return "não significativo"

def interpreta_linha(dep_var, var, coef, p):
    direcao = "positivo" if coef is not None and coef > 0 else "negativo"
    sig = nivel_sig(p) if p is not None else "—"
    if var == "Dummy ISE":
        return f"Efeito {direcao} {sig} do ESG (ISE) sobre {dep_var} mantendo controles constantes."
    elif var == "const":
        return f"Intercepto do modelo de {dep_var}."
    else:
        return f"Efeito {direcao} {sig} de {var} sobre {dep_var}."

# ------------- Início do log (captura terminal) -------------
class Tee:
    def __init__(self, *streams): self.streams = streams
    def write(self, data): 
        for s in self.streams: 
            try: s.write(data)
            except: pass
    def flush(self): 
        for s in self.streams:
            try: s.flush()
            except: pass

orig_stdout = os.sys.stdout
log_file = open(LOG_PATH, "w", encoding="utf-8")
os.sys.stdout = Tee(orig_stdout, log_file)
print(">>> Iniciando análise…")

# ------------------ 1) Leitura e preparo ------------------
df = safe_read_excel(BASE_PATH, sheet_name="Base")  # ou remova sheet_name se o nome for diferente
df.columns = df.columns.str.strip()
df = df.apply(pd.to_numeric, errors="ignore")
num_cols = df.select_dtypes(include=[np.number]).columns
df[num_cols] = df[num_cols].apply(pd.to_numeric, errors='coerce')

# Filtra observações válidas
df = df.dropna(subset=["Receita", "Lucro Líquido", "Patrimônio", "Ativos"], how="any").copy()

# Derivadas (reforça caso as fórmulas não existam no arquivo)
df["ROE"] = df["Lucro Líquido"] / df["Patrimônio"]
df["ROA"] = df["Lucro Líquido"] / df["Ativos"]
df["Margem_Liquida"] = df["Lucro Líquido"] / df["Receita"]
df["Divida_PL"] = df["Passivo Total"] / df["Patrimônio"]
df["EV_EBITDA"] = df["EV (calc)"] / df["EBITDA (aprox)"]

print("✅ Base carregada. Registros:", len(df))
print(df.head())

# ------------------ 2) Descritivas ------------------
print("\n📊 Estatísticas descritivas (principais métricas):")
print(df[["ROE","ROA","Margem_Liquida","EV_EBITDA","Divida_PL"]].describe().T[["mean","std","min","max"]])

# ------------------ 3) Correlações ------------------
corr_vars = ["Dummy ISE","ROE","ROA","Margem_Liquida","EV_EBITDA","Divida_PL"]
corr = df[corr_vars].corr()
print("\n🔍 Correlação ESG x Indicadores:")
print(corr["Dummy ISE"].sort_values(ascending=False))

# Heatmap
plt.figure(figsize=(8,5))
sns.heatmap(corr, annot=True, fmt=".2f", cmap="coolwarm", vmin=-1, vmax=1)
plt.title("Correlação entre ESG (ISE) e Indicadores Financeiros")
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR, "figura1_heatmap_correlacao.png"), dpi=220)
plt.close()

# ------------------ 4) Boxplots ------------------
sns.set(style="whitegrid")
for var, fig_name, titulo in [
    ("ROE", "figura2_boxplot_ROE.png", "ROE por Grupo ESG (ISE=1) vs Não ESG (ISE=0)"),
    ("Margem_Liquida", "figura3_boxplot_Margem.png", "Margem Líquida por Grupo ESG"),
    ("EV_EBITDA", "figura4_boxplot_EV_EBITDA.png", "EV/EBITDA por Grupo ESG"),
]:
    plt.figure(figsize=(6,4))
    sns.boxplot(x="Dummy ISE", y=var, data=df)
    plt.title(titulo)
    plt.xlabel("Participação no ISE (0=Não, 1=Sim)")
    plt.ylabel(var)
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR, fig_name), dpi=220)
    plt.close()

print(f"🖼️ Figuras salvas em: {FIG_DIR}")

# ------------------ 5) Regressões OLS ------------------
def rodar_regressao(dep_var: str):
    X = df[["Dummy ISE","Ativos","Passivo Total","Patrimônio"]].copy()
    X = sm.add_constant(X)
    y = df[dep_var]
    model = sm.OLS(y, X, missing="drop").fit(cov_type="HC3")
    # salva resumo em txt
    with open(os.path.join(OUT_DIR, f"modelo_{dep_var}.txt"), "w", encoding="utf-8") as f:
        f.write(model.summary().as_text())
    print(f"\n📈 Regressão {dep_var} concluída. Resumo salvo em modelo_{dep_var}.txt")
    return model

modelo_roe    = rodar_regressao("ROE")
modelo_margem = rodar_regressao("Margem_Liquida")
modelo_ev     = rodar_regressao("EV_EBITDA")

# ------------------ 6) Tabelas arrumadas (Excel) ------------------
# a) Correlações (arrumadinha)
corr_out = corr.round(3).reset_index().rename(columns={"index":"Variável"})
corr_out.to_excel(OUT_XLSX, sheet_name="Correlacoes", index=False)

# b) Tidy de cada modelo
def tidy(model, dep_var):
    out = pd.DataFrame({
        "Variavel": model.params.index,
        "Coeficiente": model.params.values,
        "Erro_Padrao": model.bse.values,
        "t": model.tvalues.values,
        "p_valor": model.pvalues.values
    })
    out["Estrelas"] = out["p_valor"].apply(estrelas)
    out["Variavel Dependente"] = dep_var
    out["Interpretacao"] = out.apply(lambda r: interpreta_linha(dep_var, r["Variavel"], r["Coeficiente"], r["p_valor"]), axis=1)
    # ordena: const, Dummy ISE, demais
    order = ["const","Dummy ISE"]
    out["ord"] = out["Variavel"].apply(lambda v: 0 if v=="const" else (1 if v=="Dummy ISE" else 2))
    out = out.sort_values(["ord","Variavel"]).drop(columns=["ord"])
    return out

tidy_roe    = tidy(modelo_roe, "ROE")
tidy_margem = tidy(modelo_margem, "Margem_Liquida")
tidy_ev     = tidy(modelo_ev, "EV_EBITDA")

with pd.ExcelWriter(OUT_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
    tidy_roe.to_excel(wr, sheet_name="Reg_ROE", index=False)
    tidy_margem.to_excel(wr, sheet_name="Reg_Margem", index=False)
    tidy_ev.to_excel(wr, sheet_name="Reg_EV_EBITDA", index=False)

# c) Tabela 1 (apenas efeito ESG/Dummy ISE, pronta pra colar no Word)
def extrair_dummy_esg(df_tidy):
    linha = df_tidy[df_tidy["Variavel"]=="Dummy ISE"].copy()
    return linha[["Variavel Dependente","Coeficiente","Erro_Padrao","t","p_valor","Estrelas","Interpretacao"]]

tab1 = pd.concat([
    extrair_dummy_esg(tidy_roe),
    extrair_dummy_esg(tidy_margem),
    extrair_dummy_esg(tidy_ev)
], axis=0).reset_index(drop=True)

# arredondamentos elegantes
tab1["Coeficiente"]  = tab1["Coeficiente"].round(3)
tab1["Erro_Padrao"]  = tab1["Erro_Padrao"].round(3)
tab1["t"]            = tab1["t"].round(2)
tab1["p_valor"]      = tab1["p_valor"].round(3)

with pd.ExcelWriter(OUT_XLSX, engine="openpyxl", mode="a", if_sheet_exists="overlay") as wr:
    tab1.to_excel(wr, sheet_name="Tabela_1_ESG", index=False, startrow=0)

print(f"\n📂 Tabelas exportadas para: {OUT_XLSX}")

# ------------- formatação básica no Excel (openpyxl) -------------
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

wb = load_workbook(OUT_XLSX)

def format_sheet(ws):
    # cabeçalho em negrito, largura colunas e alinhamento
    for c in range(1, ws.max_column+1):
        ws.cell(row=1, column=c).font = Font(bold=True)
        ws.column_dimensions[get_column_letter(c)].width = 22
    for r in range(2, ws.max_row+1):
        for c in range(1, ws.max_column+1):
            ws.cell(row=r, column=c).alignment = Alignment(vertical="center")

# aplica nas folhas criadas
for nm in ["Correlacoes","Reg_ROE","Reg_Margem","Reg_EV_EBITDA","Tabela_1_ESG"]:
    if nm in wb.sheetnames:
        format_sheet(wb[nm])

# legenda da Tabela 1 (linha abaixo da tabela)
ws_tab1 = wb["Tabela_1_ESG"]
legenda_row = ws_tab1.max_row + 2
ws_tab1.cell(row=legenda_row, column=1, value="Tabela 1 – Resultados do efeito ESG (Dummy ISE) por variável dependente.")
ws_tab1.cell(row=legenda_row+1, column=1, value="Fonte: elaboração própria (2025), com base em CVM e B3. *** p<0,01; ** p<0,05; * p<0,10.")

wb.save(OUT_XLSX)

print("\n✅ Finalizado com sucesso.")
print(f"- Excel de resultados: {OUT_XLSX}")
print(f"- Resumos dos modelos: {os.path.join(OUT_DIR, 'modelo_ROE.txt')}, modelo_Margem_Liquida.txt, modelo_EV_EBITDA.txt")
print(f"- Log completo: {LOG_PATH}")
print(f"- Figuras: {FIG_DIR}")

# restaura stdout
os.sys.stdout = orig_stdout
log_file.close()

