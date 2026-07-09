import pandas as pd

# Caminho da pasta de trabalho (raw string evita erros de escape no Windows)
pasta = r"C:\Users\Wever\Desktop\Conciliação"

# Arquivos de entrada
arquivo_planilha = rf"{pasta}\NOVA - ABRIL - Balanço Braslog.xlsx"
arquivo_extrato  = rf"{pasta}\Sicoob Abril.csv"
# Arquivo de saída
arquivo_saida    = rf"{pasta}\Resultado_Conciliacao.xlsx"


# =========================================================================
# 1. LEITURA DA PLANILHA DE LANÇAMENTOS
#    - Aba "3"
#    - Os lançamentos reais começam na linha 6 (cabeçalho na linha 6)
#      -> header=5 (0-indexado)
# =========================================================================
df_plan = pd.read_excel(arquivo_planilha, sheet_name="3", header=5)
# Limpa espaços dos nomes de coluna
df_plan.columns = [str(c).strip() for c in df_plan.columns]

# Filtra SOMENTE os lançamentos da conta Sicoob (CONTA = SIC)
df_plan = df_plan[df_plan["CONTA"] == "SIC"].copy()


# =========================================================================
# 2. PADRONIZAÇÃO DAS COLUNAS-CHAVE DA PLANILHA
#    Coluna C -> DATA      (data de referência)
#    Coluna R -> Recebido  (crédito / entrada  -> direção C)
#    Coluna T -> Pago      (débito  / saída    -> direção D)
# =========================================================================
df_plan["DATA"] = pd.to_datetime(df_plan["DATA"], errors="coerce").dt.normalize()
df_plan["Recebido"] = pd.to_numeric(df_plan["Recebido"], errors="coerce").fillna(0)
df_plan["Pago"]     = pd.to_numeric(df_plan["Pago"],     errors="coerce").fillna(0)

# Composição do Valor + Direção a partir das colunas Recebido/Pago
# (cada linha é OU recebimento OU pagamento)
def compoe_valor_direcao(row):
    if row["Recebido"] > 0:                                  # entrada
        return pd.Series({"Valor": row["Recebido"], "Direcao": "C"})
    if row["Pago"] > 0:                                      # saída
        return pd.Series({"Valor": row["Pago"], "Direcao": "D"})
    return pd.Series({"Valor": 0.0, "Direcao": None})

df_plan[["Valor", "Direcao"]] = df_plan.apply(compoe_valor_direcao, axis=1)

# Mantém apenas lançamentos que tenham valor (Recebido ou Pago > 0)
df_plan = df_plan[df_plan["Valor"] > 0].copy()
# Arredonda para 2 casas (evita erro de precisão de float no cruzamento)
df_plan["Valor"] = df_plan["Valor"].round(2)


# =========================================================================
# 3. LEITURA DOS EXTRATOS BANCÁRIOS (CSV) - ABRIL, MAIO E JUNHO
#    - separador: vírgula  |  decimal: vírgula  |  milhar: ponto
#    - data: DD/MM/AAAA
#    - tipo: D = saída  /  C = entrada
#    Os três arquivos são lidos e combinados num único DataFrame,
#    com uma coluna "Mes" indicando a origem de cada lançamento.
# =========================================================================
arquivos_extrato = [
    rf"{pasta}\Sicoob Abril.csv",
    rf"{pasta}\Sicoob Maio.csv",
    rf"{pasta}\Sicoob Junho.csv",
]
meses_extrato = ["Abril", "Maio", "Junho"]

lista_ext = []
for arquivo, mes in zip(arquivos_extrato, meses_extrato):
    df_tmp = pd.read_csv(
        arquivo,
        sep=",",
        decimal=",",
        thousands=".",
        encoding="utf-8",
    )
    df_tmp.columns = [str(c).strip() for c in df_tmp.columns]
    lista_ext.append(df_tmp)

# Concatena os três extratos em um só
df_ext = pd.concat(lista_ext, ignore_index=True)

df_ext["data"]   = pd.to_datetime(df_ext["data"], format="%d/%m/%Y", errors="coerce").dt.normalize()
df_ext["valor"]  = pd.to_numeric(df_ext["valor"], errors="coerce")
# Renomeia para unificar os nomes das chaves com a planilha
df_ext = df_ext.rename(columns={"data": "DATA", "valor": "Valor"})
df_ext["Direcao"] = df_ext["tipo"]          # C (entrada) ou D (saída)
df_ext["Valor"]   = df_ext["Valor"].round(2)


# =========================================================================
# 4. FILTRO DE PERÍODO
#    Restringe a planilha ao mesmo intervalo de datas do extrato
#    (extrato é de abril; evita que lançamentos de outros meses sujem o Ok)
# =========================================================================
data_min = df_ext["DATA"].min()
data_max = df_ext["DATA"].max()
df_plan  = df_plan[(df_plan["DATA"] >= data_min) & (df_plan["DATA"] <= data_max)].copy()


# =========================================================================
# 5. NUMERAR OCORRÊNCIAS REPETIDAS (mesma Data + Valor + Direção)
#    O cumcount gera 0, 1, 2... dentro de cada grupo repetido, permitindo
#    casar 1 para 1 mesmo quando existem lançamentos idênticos no mesmo dia.
# =========================================================================
df_plan["Seq"] = df_plan.groupby(["DATA", "Valor", "Direcao"]).cumcount()
df_ext["Seq"]  = df_ext.groupby(["DATA", "Valor", "Direcao"]).cumcount()


# =========================================================================
# 6. MERGE TOTAL (outer) ENTRE AS DUAS BASES
#    Chaves: Data + Valor + Direção (C/D) + Sequência
# =========================================================================
colunas_plan = ["FORNECEDOR", "DESTINO", "Categoria", "SUB Categoria", "OBSERVAÇÃO"]
colunas_ext  = ["historico", "detalhes"]

# Mantém só as colunas informativas relevantes + as chaves
df_plan_keep = df_plan[["DATA", "Valor", "Direcao", "Seq"] + [c for c in colunas_plan if c in df_plan.columns]].copy()
df_ext_keep  = df_ext[["DATA", "Valor", "Direcao", "Seq"]  + [c for c in colunas_ext  if c in df_ext.columns]].copy()

df = pd.merge(
    df_plan_keep,
    df_ext_keep,
    on=["DATA", "Valor", "Direcao", "Seq"],
    how="outer",
    suffixes=("_plan", "_ext"),
    indicator=True,                 # cria coluna '_merge' com a origem de cada linha
)


# =========================================================================
# 7. COLUNA FINAL 'Status'
# =========================================================================
df["Status"] = df["_merge"].map({
    "both":       "Ok",                 # existe nos dois arquivos
    "left_only":  "Falta no Extrato",   # só na planilha de lançamentos
    "right_only": "Falta na Planilha",  # só no extrato bancário
})
df = df.drop(columns=["_merge"])

# Rótulo legível da direção
df["Direcao"] = df["Direcao"].map({"C": "Entrada", "D": "Saída"})

# Deriva o mês diretamente da DATA para TODAS as linhas
# (garante que entradas vindas só da planilha também tenham o mês)
mapa_meses = {4: "Abril", 5: "Maio", 6: "Junho"}
df["Mes"] = df["DATA"].dt.month.map(mapa_meses)


# =========================================================================
# 8. ORGANIZAR E SALVAR O RESULTADO
# =========================================================================
# Ordem final das colunas
ordem = [
    "Status", "Mes", "DATA", "Valor", "Direcao",
    "FORNECEDOR", "DESTINO", "Categoria", "SUB Categoria", "OBSERVAÇÃO",
    "historico", "detalhes",
]
colunas_finais = [c for c in ordem if c in df.columns]
df = df[colunas_finais]

# Ordena por data e status (Ok primeiro, depois divergências)
df = df.sort_values(["DATA", "Status", "Valor"]).reset_index(drop=True)

df.to_excel(arquivo_saida, index=False)

print("Conciliação concluída! Arquivo salvo em:", arquivo_saida)
print(df["Status"].value_counts())
