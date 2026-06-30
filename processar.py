from pathlib import Path
import pandas as pd
import os

# ==================================================
# CONFIGURAÇÃO
# ==================================================

BASE = Path(__file__).resolve().parent
PASTA_UPLOAD = BASE / "upload"
SAIDA = PASTA_UPLOAD / "ordemcronologica_2026.xlsx"

# Ignora o próprio arquivo de saída de execuções anteriores ao buscar o
# arquivo bruto de entrada (evita reprocessar um arquivo já processado,
# que tem estrutura diferente e causaria KeyError).
arquivos = [
    a for a in PASTA_UPLOAD.glob("*.xlsx")
    if a.name != SAIDA.name and not a.name.startswith("~$")
]

if not arquivos:
    raise Exception(
        "Nenhum arquivo XLSX de entrada encontrado na pasta upload "
        "(o arquivo de saída anterior foi ignorado)."
    )

if len(arquivos) > 1:
    nomes = ", ".join(a.name for a in arquivos)
    raise Exception(
        f"Mais de um arquivo XLSX de entrada encontrado na pasta upload: {nomes}. "
        "Deixe apenas o arquivo bruto a ser processado."
    )

ARQUIVO = arquivos[0]

print(f"Arquivo encontrado: {ARQUIVO.name}")

# ==================================================
# FUNÇÕES
# ==================================================

def moeda_para_float(valor):
    """
    Converte valor monetário para float.
    Tenta float direto primeiro (formato "852.99" já com ponto decimal).
    Fallback para formato BR "1.234,56" → remove milhar, troca vírgula por ponto.
    IMPORTANTE: não remover o ponto antes de tentar float direto!
    """
    if pd.isna(valor):
        return 0.0
    s = str(valor).strip()
    if s in ["", "/", "nan", "None"]:
        return 0.0
    # Tentativa direta (já está em formato float padrão)
    try:
        return float(s)
    except:
        pass
    # Fallback: formato brasileiro "1.234,56"
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

# ==================================================
# LEITURA
# ==================================================

print("Lendo arquivo...")

df = pd.read_excel(
    ARQUIVO,
    header=None,
    engine="openpyxl",
    dtype=str
)

# remove primeira linha (em branco)
df = df.iloc[1:].reset_index(drop=True)

# remove primeira coluna (em branco)
df = df.iloc[:, 1:].reset_index(drop=True)

# primeira linha vira cabeçalho
df.columns = df.iloc[0]
df = df.iloc[1:].reset_index(drop=True)

# limpa nomes das colunas
df.columns = [str(c).strip() for c in df.columns]

print(f"Registros carregados: {len(df):,}")

# Validação: garante que as colunas esperadas existem após o parse do
# cabeçalho. Se faltar alguma, o arquivo de entrada provavelmente não tem
# a estrutura bruta esperada (linha/coluna em branco no início).
_colunas_esperadas = [
    "data_liquidacao", "data_pagamento", "valor_despesa_liquidada",
    "valor_pago_financeiro", "numero_empenho", "numero_op"
]
_faltando = [c for c in _colunas_esperadas if c not in df.columns]
if _faltando:
    raise Exception(
        f"Colunas esperadas não encontradas no arquivo: {_faltando}. "
        f"Colunas lidas: {list(df.columns)}. "
        "Verifique se o ARQUIVO de entrada é o arquivo bruto original "
        "(com uma linha e uma coluna em branco antes do cabeçalho) e não "
        "um arquivo já processado por este script."
    )

# ==================================================
# LIMPEZA
# ==================================================

for col in [
    "data_liquidacao",
    "data_pagamento",
    "valor_despesa_liquidada",
    "valor_pago_financeiro"
]:
    df[col] = df[col].fillna("").astype(str).str.strip()

# Remove apenas linhas completamente vazias (lixo).
# Linhas pendentes (sem datas mas com numero_op e valor) são MANTIDAS.
mask_lixo = (
    df["data_liquidacao"].eq("")
    & df["data_pagamento"].eq("")
    & (df["valor_despesa_liquidada"].apply(moeda_para_float).eq(0))
    & (df["valor_pago_financeiro"].apply(moeda_para_float).eq(0))
)

removidos = mask_lixo.sum()
df = df[~mask_lixo].copy()

print(f"Linhas removidas: {removidos:,}")
print(f"Registros após limpeza: {len(df):,}")

# ==================================================
# DATAS
# ==================================================

def limpar_data(valor):
    if pd.isna(valor):
        return None
    try:
        num = float(valor)
        if num < 0:
            return None
    except:
        pass
    return valor

for col in ["data_empenho", "data_liquidacao", "data_registro", "data_pagamento"]:
    df[col] = df[col].apply(limpar_data)
    df[col] = pd.to_datetime(df[col], errors="coerce")

# ==================================================
# VALORES
# ==================================================

df["valor_liquidado_float"] = df["valor_despesa_liquidada"].apply(moeda_para_float)
df["valor_pago_float"] = df["valor_pago_financeiro"].apply(moeda_para_float)

# ==================================================
# SEPARAÇÃO: LIQUIDAÇÕES / PAGAMENTOS
# ==================================================

print(
    "Datas de liquidação inválidas:",
    df["data_liquidacao"].isna().sum()
)
print(
    "Datas de pagamento inválidas:",
    df["data_pagamento"].isna().sum()
)

liquidacoes = df[df["data_liquidacao"].notna()].copy()

# Pagamentos = linhas com data_pagamento OU linhas pendentes
# (sem data_pagamento mas com numero_op válido e valor > 0).
_op_valido = (
    df["numero_op"].notna()
    & ~df["numero_op"].astype(str).str.strip().isin(["", "0", "nan"])
)
pagamentos = df[
    df["data_pagamento"].notna()
    | (df["data_pagamento"].isna() & _op_valido & (df["valor_pago_float"] > 0))
].copy()

print(f"Liquidações: {len(liquidacoes):,}")
print(f"Pagamentos (incl. pendentes): {len(pagamentos):,}")

# ==================================================
# CHAVE DE CORRESPONDÊNCIA
# ==================================================

CHAVE = [
    "ano",
    "codigo_unidade_orcamentaria",
    "nome_unidade_orcamentaria",
    "codigo_unidade_executora",
    "funcional_programatica",
    "codigo_elemento_despesa",
    "descricao_elemento_despesa",
    "codigo_item_despesa",
    "descricao_item_despesa",
    "codigo_fonte_recurso",
    "descricao_fonte_recurso",
    "numero_empenho",
    "razao_social_credor",
    "cnpj_cpf_credor"
]

resultado = []
pagamentos_utilizados = set()

# Tolerância para considerar pagamento "compartilhado" (cobre múltiplas
# liquidações): pagamento significativamente maior que a liquidação.
TOLERANCIA = 0.15

# ==================================================
# PROCESSAMENTO PRINCIPAL
# ==================================================
# Processa liquidações em ordem cronológica DENTRO de cada grupo de CHAVE
# (mesmo empenho/credor/etc), pois pagamentos tendem a seguir a mesma
# ordem cronológica das liquidações que os originaram.

liquidacoes_ordenadas = liquidacoes.sort_values(
    CHAVE + ["data_liquidacao"], kind="stable"
)

for i, (_, liq) in enumerate(liquidacoes_ordenadas.iterrows(), start=1):

    if i % 1000 == 0:
        print(f"Processadas {i:,} liquidações...")

    # Liquidação cancelada/anulada: não inclui no resultado
    if liq["valor_liquidado_float"] <= 0:
        continue

    # Filtra candidatos pela CHAVE completa
    candidatos = pagamentos.copy()
    for campo in CHAVE:
        candidatos = candidatos[
            candidatos[campo].astype(str) == str(liq[campo])
        ]

    # Para pagamentos confirmados: data_pagamento >= data_liquidacao.
    # Para pendentes: sem restrição de data.
    mask_data_ok = (
        candidatos["data_pagamento"].isna()
        | (candidatos["data_pagamento"] >= liq["data_liquidacao"])
    )
    candidatos = candidatos[mask_data_ok]

    # Ignora pagamentos já usados exclusivamente por outra liquidação
    candidatos = candidatos[~candidatos.index.isin(pagamentos_utilizados)]

    # Sem candidatos = liquidação sem correspondência → excluir do resultado
    if len(candidatos) == 0:
        continue

    candidatos = candidatos.copy()

    candidatos["dias"] = (
        candidatos["data_pagamento"] - liq["data_liquidacao"]
    ).dt.days.abs()
    candidatos["dif_valor"] = (
        candidatos["valor_pago_float"] - liq["valor_liquidado_float"]
    ).abs()
    candidatos["sem_data"] = candidatos["data_pagamento"].isna().astype(int)

    liq_val = liq["valor_liquidado_float"]
    candidatos["valor_compativel"] = (
        candidatos["valor_pago_float"] <= liq_val * (1 + TOLERANCIA)
    ).astype(int)
    candidatos["dif_valor_pct"] = candidatos["dif_valor"] / max(liq_val, 1.0)

    # Existem OUTRAS liquidações no mesmo grupo (mesma CHAVE) com a MESMA
    # data_liquidacao desta? Se sim, "dias" sozinho não consegue
    # distinguir entre elas (todas teriam o mesmo dias mínimo possível
    # para um dado pagamento) — nesse caso o VALOR é o sinal decisivo.
    # Caso contrário (liquidações espalhadas em datas diferentes), a
    # PROXIMIDADE DE DATA é mais confiável, pois descontos percentuais
    # fixos (retenção padronizada) fazem a diferença de valor ficar
    # ambígua entre liquidações de datas distintas.
    outras_mesma_data = liquidacoes_ordenadas[
        (liquidacoes_ordenadas["data_liquidacao"] == liq["data_liquidacao"])
        & (liquidacoes_ordenadas.index != liq.name)
    ]
    mesma_data_existe = False
    if len(outras_mesma_data) > 0:
        for campo in CHAVE:
            outras_mesma_data = outras_mesma_data[
                outras_mesma_data[campo].astype(str) == str(liq[campo])
            ]
        mesma_data_existe = len(outras_mesma_data) > 0

    if mesma_data_existe:
        # Valor decide primeiro (dentro da tolerância), dias desempata
        candidatos["pontuacao"] = (
            candidatos["dif_valor_pct"] * 1000 + candidatos["dias"]
        )
    else:
        # Data decide primeiro; diferença de valor além da tolerância
        # ainda penaliza (sinal de candidato provavelmente errado)
        candidatos["excesso_pct"] = (
            candidatos["dif_valor_pct"] - TOLERANCIA
        ).clip(lower=0)
        candidatos["pontuacao"] = (
            candidatos["dias"] + candidatos["excesso_pct"] * 1000
        )

    melhor_idx = (
        candidatos
        .sort_values(
            ["sem_data", "valor_compativel", "pontuacao"],
            ascending=[True, False, True],
        )
        .index[0]
    )

    pag = pagamentos.loc[melhor_idx]
    pag_val = pag["valor_pago_float"]

    # SANIDADE: se o melhor candidato disponível tem valor muito menor que
    # a liquidação (abaixo de 50%), provavelmente é um pagamento "sobra"
    # de outra liquidação do mesmo empenho que já foi atendida por outro
    # pagamento (ex: split de pagamento em duas parcelas, onde a segunda
    # parcela pequena já foi consumida e sobrou disponível, mas não tem
    # relação real com ESTA liquidação). Retenções/descontos legítimos
    # normalmente não reduzem o valor pago a menos da metade do liquidado;
    # quando isso ocorre, é mais provável que não exista correspondência
    # real disponível neste lote — melhor excluir a liquidação do que
    # gerar uma vinculação claramente incorreta.
    if pag_val < liq_val * 0.50:
        continue

    # TIPO DE CORRESPONDÊNCIA:
    # 1. Direta: pag_val <= liq_val * 1.15 → 1 pagamento para 1 liquidação.
    #    Inclui casos onde pag < liq (retenções tributárias, glosas,
    #    descontos, INSS/IR retido na fonte são legítimos e comuns — NÃO
    #    devem causar exclusão da liquidação).
    # 2. Compartilhada: pag_val > liq_val * 1.15 → pagamento
    #    significativamente MAIOR que a liquidação, indicando que cobre
    #    múltiplas liquidações (1 pagamento para N liquidações). Mostra o
    #    valor real do pagamento e marca observacao_script.
    #
    # IMPORTANTE: nunca se exclui uma liquidação apenas pela diferença de
    # valor — só é excluída quando não existe candidato algum (acima).

    eh_compartilhada = pag_val > liq_val * (1 + TOLERANCIA)

    if eh_compartilhada:
        # 1:N — pagamento cobre múltiplas liquidações.
        # Não remove do pool para que outras liq também possam usá-lo.
        nova_linha = liq.copy()
        nova_linha["numero_op"] = pag["numero_op"]
        nova_linha["data_registro"] = pag["data_registro"]
        nova_linha["data_pagamento"] = pag["data_pagamento"]
        nova_linha["valor_pago_financeiro"] = pag["valor_pago_financeiro"]
        nova_linha["observacao_script"] = "1 pagamento / 2 liquidações"
        resultado.append(nova_linha)
    else:
        # 1:1 — inclui casos de pag < liq (retenções/descontos legítimos).
        pagamentos_utilizados.add(melhor_idx)
        nova_linha = liq.copy()
        nova_linha["numero_op"] = pag["numero_op"]
        nova_linha["data_registro"] = pag["data_registro"]
        nova_linha["data_pagamento"] = pag["data_pagamento"]
        nova_linha["valor_pago_financeiro"] = pag["valor_pago_financeiro"]
        nova_linha["observacao_script"] = None
        resultado.append(nova_linha)

# ==================================================
# RESULTADO
# ==================================================

final = pd.DataFrame(resultado)

# Remove colunas auxiliares
for col in ["valor_liquidado_float", "valor_pago_float", "dias", "dif_valor", "sem_data"]:
    if col in final.columns:
        final = final.drop(columns=[col])

# Formata datas
for col in ["data_empenho", "data_liquidacao", "data_registro", "data_pagamento"]:
    if col in final.columns:
        final[col] = final[col].dt.strftime("%d/%m/%Y")

# Formata valores numéricos
for col in ["valor_despesa_liquidada", "valor_pago_financeiro"]:
    if col in final.columns:
        final[col] = (
            pd.to_numeric(final[col], errors="coerce")
            .round(2)
        )

print()
print("=" * 60)
print(f"Registros finais: {len(final):,}")
print("=" * 60)

# ==================================================
# EXPORTA
# ==================================================

if len(final) == 0:
    raise Exception("Arquivo final vazio. Exportação cancelada.")

final.to_excel(SAIDA, index=False)

print(f"Arquivo gerado: {SAIDA.name}")

if Path(ARQUIVO) != Path(SAIDA):
    os.remove(ARQUIVO)
    print(f"Arquivo removido: {Path(ARQUIVO).name}")
