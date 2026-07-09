import pandas as pd
import numpy as np

# Caminho da pasta de trabalho (raw string evita erros de escape no Windows)
pasta = r"C:\Users\Wever\Desktop\Conciliação"

# Arquivo de entrada (planilha de lançamentos)
arquivo_planilha = rf"{pasta}\NOVA - ABRIL - Balanço Braslog.xlsx"
# Arquivo de saída (novo - separado da conciliação)
arquivo_saida    = rf"{pasta}\Diferenca.xlsx"


# =========================================================================
# 1. LEITURA DA PLANILHA DE LANÇAMENTOS
#    - Aba "3", com os lançamentos reais a partir da linha 6 (header=5)
# =========================================================================
df = pd.read_excel(arquivo_planilha, sheet_name="3", header=5)
df.columns = [str(c).strip() for c in df.columns]

# Localiza a coluna de OBSERVAÇÃO (tem acentuação, evita erro de encode)
col_obs = [c for c in df.columns if "OBS" in c.upper()][0]

# Padroniza datas e valores
df["DATA"]     = pd.to_datetime(df["DATA"], errors="coerce").dt.normalize()
df["Recebido"] = pd.to_numeric(df["Recebido"], errors="coerce").fillna(0)
df["Pago"]     = pd.to_numeric(df["Pago"],     errors="coerce").fillna(0)


# =========================================================================
# 2. COLUNAS ONDE SERÃO BUSCADOS OS TERMOS (inclui OBSERVAÇÃO)
# =========================================================================
colunas_busca = [
    "DESTINO", "FORNECEDOR", "Categoria", "SUB Categoria",
    "Centro de custo", col_obs,
]
# Concatena o conteúdo de todas as colunas de busca numa só string (maiúscula)
texto = (
    df[colunas_busca]
    .fillna("")
    .astype(str)
    .agg(" ".join, axis=1)
    .str.upper()
)


# =========================================================================
# 3. FILTRO DAS TRANSFERÊNCIAS ENTRE BRASLOG E MH (só empréstimos/APORTES)
#    - envolve MH Gestão
#    - é operação de APORTES / EMPRÉSTIMO (transferência financeira entre as empresas)
#    - NÃO inclui adiantamento salarial / antecipação
# =========================================================================
masc_mh       = texto.str.contains("MH GEST", na=False) | texto.str.contains("MHGEST", na=False)
masc_aportes  = (df["Categoria"].fillna("").str.upper() == "APORTES") | \
                texto.str.contains("EMPREST", na=False)
masc_adiant   = texto.str.contains("ADIANT", na=False) | \
                texto.str.contains("ANTECIP", na=False) | \
                texto.str.contains("SALARI", na=False) | \
                texto.str.contains("SALÁR", na=False)

# Mantém transferências Braslog <-> MH, excluindo adiantamento salarial
df_filtrado = df[masc_mh & masc_aportes & ~masc_adiant].copy()

# =========================================================================
# 4. DIREÇÃO DA TRANSFERÊNCIA
#    - Pago > 0     -> Braslog pagou/emprestou a MH  -> "MH deve a Braslog"
#    - Recebido > 0 -> Braslog recebeu de MH         -> "Braslog deve a MH"
# =========================================================================
def direcao(row):
    if row["Pago"] > 0:
        return "Braslog -> MH (MH deve a Braslog)"
    if row["Recebido"] > 0:
        return "MH -> Braslog (Braslog deve a MH)"
    return ""

df_filtrado["Direção"] = df_filtrado.apply(direcao, axis=1)


# =========================================================================
# 5. CÁLCULO DO SALDO: QUEM DEVE A QUEM
#    Saldo = Pago - Recebido
#      > 0 -> MH deve a Braslog
#      < 0 -> Braslog deve a MH
# =========================================================================
total_pago     = df_filtrado["Pago"].sum()
total_recebido = df_filtrado["Recebido"].sum()
saldo          = total_pago - total_recebido

if saldo > 0.005:
    devedor, credor, valor = "MH", "Braslog", saldo
elif saldo < -0.005:
    devedor, credor, valor = "Braslog", "MH", abs(saldo)
else:
    devedor, credor, valor = None, None, 0.0


# =========================================================================
# 6. ORGANIZAR E SALVAR O RESULTADO
# =========================================================================
# Ordem final das colunas
ordem = [
    "Direção", "DATA", "CONTA", "BANCO",
    "DESTINO", "FORNECEDOR", "Centro de custo", "Categoria", "SUB Categoria",
    "Recebido", "Pago",
    col_obs,
]
colunas_finais = [c for c in ordem if c in df_filtrado.columns]
df_filtrado = df_filtrado[colunas_finais]

# Ordena por Data
df_filtrado = df_filtrado.sort_values(["DATA"]).reset_index(drop=True)

# Adiciona linhas de resumo no final da planilha
linhas_resumo = [
    {c: "" for c in colunas_finais},
    {**{c: "" for c in colunas_finais}, "Direção": "RESUMO DAS TRANSFERÊNCIAS BRASLOG <-> MH"},
    {**{c: "" for c in colunas_finais}, "Direção": "Total pago por Braslog a MH (Pago):",
     "Pago": total_pago},
    {**{c: "" for c in colunas_finais}, "Direção": "Total recebido por Braslog de MH (Recebido):",
     "Recebido": total_recebido},
    {**{c: "" for c in colunas_finais}, "Direção": "Saldo líquido (Pago - Recebido):",
     "Pago": saldo},
]
if devedor:
    linhas_resumo.append(
        {**{c: "" for c in colunas_finais},
         "Direção": f"=> {devedor} DEVE a {credor}:",
         "Pago": valor}
    )
df_saida = pd.concat([df_filtrado, pd.DataFrame(linhas_resumo)], ignore_index=True)

df_saida.to_excel(arquivo_saida, index=False)

print("Arquivo gerado em:", arquivo_saida)
print("Total de transferências:", len(df_filtrado))
print()
print(f"Total pago por Braslog a MH:      R$ {total_pago:,.2f}")
print(f"Total recebido por Braslog de MH: R$ {total_recebido:,.2f}")
print(f"Saldo líquido:                    R$ {saldo:,.2f}")
print()
if devedor:
    print(f">>> {devedor} DEVE a {credor}: R$ {valor:,.2f}")
else:
    print(">>> Contas zeradas.")
