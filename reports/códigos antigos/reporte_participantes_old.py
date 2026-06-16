"""
Tratamento da planilha de participantes — Cambalhotas / THE.

O QUE ESTE SCRIPT FAZ:
  1. Seleciona apenas as colunas desejadas (na ordem definida).
  2. Substitui os códigos numéricos por nomes (Localidade, Projeto, FY, Trimestre, Turmas).
  3. Renomeia as colunas para rótulos mais legíveis.
  4. Salva o resultado em um novo arquivo .xlsx.

"""

import pandas as pd

# ============================ CONFIGURAÇÃO ============================

ARQUIVO_ENTRADA = "arquivos_baixados_sharepoint/[main]_participantesPUs_json.xlsx" # "/" está diferente do main
ARQUIVO_SAIDA = "Cambalhotas_THE_tratado.xlsx"

# 1) Colunas mantidas, na ordem desejada (nomes ORIGINAIS do arquivo)
COLUNAS_DESEJADAS = [
    "nomeCompleto",
    "genero",
    "racaEtnia",
    "dataNascimentoTxt",
    "escolaridade",
    "escolaParticipante",
    "periodoEstudo",
    "nomeResponsavel",
    "telefoneResponsavel",
    "endereço",
    "IDTurma",
    "comunidade",
    "IDFY",
    "IDProjeto",
    "IDPU",
    "IDTrimestre",
]

# 2) Substituição de valores: código -> nome
MAPA_IDPU = {3: "TERESINA"}
MAPA_IDPROJETO = {22: "Cambalhotas - THE"}
MAPA_IDFY = {
    7: "FY25",
    8: "FY26",
    9: "FY27",
    10: "FY28",
    11: "FY29",
    12: "FY30",
    }
MAPA_IDTRIMESTRE = {
    1: "T1",
    2: "T2",
    3: "T3",
    4: "T4"
    }
MAPA_IDTURMA = {
    15: "SANTA TERESA",
    246: "ALDA FREITAS",
    131: "GUARUPÁ",
    132: "SOCOPO",
    133: "BAIXÃO DO CARLOS",
    134: "SÃO VICENTE DE CIMA",
    135: "AMPARO",
}

# 3) Renomeação das colunas: nome original -> nome final amigável
RENOMEAR = {
    "nomeCompleto": "Nome completo",
    "genero": "Gênero",
    "racaEtnia": "Raça/Etnia",
    "dataNascimentoTxt": "Data de nascimento",
    "escolaridade": "Escolaridade",
    "escolaParticipante": "Escola",
    "periodoEstudo": "Período de estudo",
    "nomeResponsavel": "Nome do responsável",
    "telefoneResponsavel": "Telefone do responsável",
    "endereço": "Endereço",
    "comunidade": "Comunidade",
    "IDTurma": "Turmas",
    "IDFY": "FY",
    "IDProjeto": "Projeto",
    "IDPU": "Localidade",
    "IDTrimestre": "Trimestre",
}


# ============================ PROCESSAMENTO ============================

def aplicar_mapa(serie: pd.Series, mapa: dict) -> pd.Series:
    """Substitui valores casando tanto números quanto texto (ex.: 131 ou '131')."""
    mapa_texto = {str(chave): valor for chave, valor in mapa.items()}
    return serie.replace(mapa).replace(mapa_texto)


def main() -> None:
    df = pd.read_excel(ARQUIVO_ENTRADA)

    # Mantém apenas as colunas desejadas que existem no arquivo
    presentes = [c for c in COLUNAS_DESEJADAS if c in df.columns]
    faltando = [c for c in COLUNAS_DESEJADAS if c not in df.columns]
    if faltando:
        print(f"Aviso: colunas não encontradas (serão ignoradas): {faltando}")
    df = df[presentes].copy()

    # Substitui os códigos pelos nomes (antes de renomear as colunas)
    if "IDPU" in df.columns:
        df["IDPU"] = aplicar_mapa(df["IDPU"], MAPA_IDPU)
    if "IDProjeto" in df.columns:
        df["IDProjeto"] = aplicar_mapa(df["IDProjeto"], MAPA_IDPROJETO)
    if "IDFY" in df.columns:
        df["IDFY"] = aplicar_mapa(df["IDFY"], MAPA_IDFY)
    if "IDTrimestre" in df.columns:
        df["IDTrimestre"] = aplicar_mapa(df["IDTrimestre"], MAPA_IDTRIMESTRE)
    if "IDTurma" in df.columns:
        df["IDTurma"] = aplicar_mapa(df["IDTurma"], MAPA_IDTURMA)

    # Renomeia as colunas
    df = df.rename(columns=RENOMEAR)

    # Salva o resultado
    df.to_excel(ARQUIVO_SAIDA, index=False)
    print(f"Arquivo gerado: {ARQUIVO_SAIDA} ({len(df)} linhas)")


if __name__ == "__main__":
    main()