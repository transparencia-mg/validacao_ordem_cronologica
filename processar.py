import pandas as pd
from pathlib import Path
import subprocess
import os

REPO = Path(__file__).parent
UPLOAD = REPO / "upload"

anos = [2022, 2023, 2024, 2025, 2026]

arquivos_gerados = 0

for ano in anos:

    arquivo = UPLOAD / f"ORDEMCRONOLOGICA{ano}.xlsx"

    if not arquivo.exists():
        print(f"Arquivo não encontrado: {arquivo.name}")
        continue

    print(f"\nProcessando {arquivo.name}...")

    # ==============================
    # CARREGAR ABAS
    # ==============================

    abas = pd.read_excel(
        arquivo,
        sheet_name=None,
        header=1
    )

    if "liquidacao" not in abas:
        print(f"Aba liquidacao não encontrada em {arquivo.name}")
        continue

    if "pagamento" not in abas:
        print(f"Aba pagamento não encontrada em {arquivo.name}")
        continue

    liq = abas["liquidacao"]
    pag = abas["pagamento"]

    # ==============================
    # REMOVE PRIMEIRA COLUNA
    # ==============================

    liq = liq.iloc[:, 1:]
    pag = pag.iloc[:, 1:]

    liq = liq.dropna(how="all")
    pag = pag.dropna(how="all")

    liq = liq.dropna(axis=1, how="all")
    pag = pag.dropna(axis=1, how="all")

    # ==============================
    # PADRONIZAÇÃO
    # ==============================

    liq.columns = liq.columns.str.strip().str.lower()
    pag.columns = pag.columns.str.strip().str.lower()

    liq["data_liquidacao"] = pd.to_datetime(
        liq["data_liquidacao"],
        errors="coerce"
    )

    liq["data_empenho"] = pd.to_datetime(
        liq["data_empenho"],
        errors="coerce"
    )

    pag["data_pagamento"] = pd.to_datetime(
        pag["data_registro"],
        errors="coerce"
    )

    pag["data_empenho"] = pd.to_datetime(
        pag["data_empenho"],
        errors="coerce"
    )

    liq["id_liq"] = liq.index

    # ==============================
    # CHAVES
    # ==============================

    chaves = [
        "codigo_uo",
        "numero_empenho",
        "data_empenho",
        "cnpj_cpf",
        "razao_social_credor"
    ]

    # ==============================
    # FIFO
    # ==============================

    matches = []

    print("Aplicando matching FIFO...")

    for chave, liq_grp in liq.groupby(chaves):

        pag_grp = pag.copy()

        for col, val in zip(chaves, chave):
            pag_grp = pag_grp[pag_grp[col] == val]

        if pag_grp.empty:
            continue

        liq_grp = liq_grp.sort_values("data_liquidacao")
        pag_grp = pag_grp.sort_values("data_registro")

        pag_list = pag_grp.to_dict("records")
        pag_idx = 0

        for _, liq_row in liq_grp.iterrows():

            while pag_idx < len(pag_list):

                pag_row = pag_list[pag_idx]

                if (
                    pd.notna(pag_row["data_registro"])
                    and pd.notna(liq_row["data_liquidacao"])
                    and pag_row["data_registro"] >= liq_row["data_liquidacao"]
                ):

                    matches.append({
                        "id_liq": liq_row["id_liq"],
                        "data_pagamento": pag_row["data_registro"]
                    })

                    pag_idx += 1
                    break

                pag_idx += 1

    # ==============================
    # RESULTADO
    # ==============================

    df_match = pd.DataFrame(matches)

    resultado = liq.merge(
        df_match,
        on="id_liq",
        how="left"
    )

    for col in [
        "data_liquidacao",
        "data_pagamento",
        "data_empenho"
    ]:
        if col in resultado.columns:
            resultado[col] = resultado[col].apply(
                lambda x: x.strftime("%d/%m/%Y")
                if pd.notnull(x)
                else ""
            )

    saida = UPLOAD / f"pagamentos{ano}.xlsx"

    resultado.to_excel(
        saida,
        index=False
    )

    print(f"Gerado: {saida.name}")

    arquivos_gerados += 1

    os.remove(arquivo)

    print(f"Removido: {arquivo.name}")

# ==============================
# GIT
# ==============================

if arquivos_gerados > 0:

    resultado_git = subprocess.run(
        ["git", "status", "--porcelain"],
        cwd=REPO,
        capture_output=True,
        text=True
    )

    if resultado_git.stdout.strip():

        subprocess.run(
            ["git", "add", "upload"],
            cwd=REPO,
            check=True
        )

        subprocess.run(
            [
                "git",
                "commit",
                "-m",
                "Atualização automática portal ordem cronologica"
            ],
            cwd=REPO,
            check=True
        )

        subprocess.run(
            ["git", "push"],
            cwd=REPO,
            check=True
        )

        print("\nGitHub atualizado com sucesso.")

    else:
        print("\nNenhuma alteração encontrada.")

else:
    print("\nNenhum arquivo foi processado.")