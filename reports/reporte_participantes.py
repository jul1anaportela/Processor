"""
Tratamento da tabela principal ([main]_participantesPUs) com PANDAS — 100% via TABELAS.

Toda substituição de código por nome é lida de uma tabela de referência na pasta
BD_sharepoint. Não há dicionário fixo no código: ao encontrar um ID, o script busca
o nome correspondente na tabela indicada.

IMPORTANTE SOBRE PASTAS:
  Por padrão, o script procura os arquivos na MESMA pasta onde este .py está
  (não importa de onde você executa o terminal). Então deixe, na mesma pasta:
         <pasta deste script>/
           projeto_manip_pandas.py
           Cambalhotas_THE.xlsx
           BD_sharepoint/
             PU.xlsx, trimestre.xlsx, projetos.xlsx, fy.xlsx, turmas.xlsx, ...
  Se preferir manter os dados em outro lugar, defina BASE_DIR com um caminho
  absoluto (veja o comentário em BASE_DIR, mais abaixo).

COMO USAR:
  1. pip install pandas openpyxl
  2. python projeto_manip_pandas.py
"""

import os
import pandas as pd

# ============================ CONFIGURAÇÃO ============================

# Pasta-base onde estão Cambalhotas_THE.xlsx e a pasta BD_sharepoint.
# Padrão: a mesma pasta deste arquivo .py (resolve o problema de "pasta atual").
# Para apontar para outro lugar, troque pela linha de exemplo abaixo:
#     BASE_DIR = r"C:\Users\JulianaPortela\Desktop\projeto_processor"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PASTA_BD = os.path.join(BASE_DIR, "BD_sharepoint")
ARQUIVO_ENTRADA = "arquivos_baixados_sharepoint/[main]_participantesPUs_json.xlsx"
ARQUIVO_SAIDA = os.path.join(BASE_DIR, "arquivos_reports/01[main]_participantes_tratado.xlsx")

# (coluna na tabela principal, arquivo de referência, coluna do CÓDIGO, coluna do NOME)
REGRAS = [
    ("IDPU",        "PU.xlsx",        "IDPU",      "Title"),     # ex.: 3  -> Teresina
    ("IDTrimestre", "trimestre.xlsx", "ID",        "ano"),       # código = coluna "ID"; ex.: 3 -> T3
    ("IDProjeto",   "projetos.xlsx",  "IDProjeto", "projetos"),  # ex.: 22 -> Cambalhotas - THE
    ("IDFY",        "fy.xlsx",        "IDFY",      "ano"),        # ex.: 7  -> FY25
    ("IDTurma",     "turmas.xlsx",    "ID",        "nomeTurma"), # código = coluna "ID"; ex.: 131 -> TURMA - GURUPÁ

    # opcionais — descomente se a tabela principal tiver essas colunas:
    ("IDCidade",      "cidades.xlsx",     "IDCidade",      "NomeCidade"),
    ("IDComunidade", "comunidades.xlsx", "IDComunidades", "NomeComunidade"),
    #("IDaogd",        "aogd.xlsx",        "IDaogd",        "area"),  # ou "pilar"
]

# Seleção de colunas (na ordem desejada). Deixe a lista vazia [] para manter TODAS.
COLUNAS_DESEJADAS = [
    "nomeCompleto", "genero", "racaEtnia", "dataNascimentoTxt", "escolaridadeTxt",
    "escolaParticipante", "periodoEstudo", "nomeResponsavel", "telefoneResponsavel",
    "endereco", "IDTurma", "IDComunidade","IDCidade", "IDFY", "IDProjeto", "IDPU", "IDTrimestre"
]

# Renomeação final das colunas (nome original -> nome amigável)
RENOMEAR = {
    "nomeCompleto": "Nome completo", "genero": "Gênero", "racaEtnia": "Raça/Etnia",
    "dataNascimentoTxt": "Data de nascimento", "escolaridadeTxt": "Escolaridade",
    "escolaParticipante": "Escola", "periodoEstudo": "Período de estudo",
    "nomeResponsavel": "Nome do responsável", "telefoneResponsavel": "Telefone do responsável",
    "endereco": "Endereço", "IDComunidade": "Comunidade", "IDTurma": "Turmas", "IDCidade": "Cidade",
    "IDFY": "FY", "IDProjeto": "Projeto", "IDPU": "PU", "IDTrimestre": "Trimestre"
}


# ============================ FUNÇÕES ============================

def _normaliza_colunas(df):
    """Remove espaços extras dos nomes de coluna."""
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    return df


def _chaves(valor):
    """Formas equivalentes de um código para casar 3, 3.0 e '3' indistintamente."""
    chaves = {valor, str(valor)}
    try:
        numero = float(valor)
        if numero.is_integer():
            chaves.add(int(numero))
            chaves.add(str(int(numero)))
    except (TypeError, ValueError):
        pass
    return chaves


def validar_regras(regras, pasta_bd):
    """Confere se a pasta e as tabelas existem e têm as colunas esperadas."""
    problemas = []

    if not os.path.isdir(pasta_bd):
        problemas.append(f"  - A pasta de referência não foi encontrada: {pasta_bd}")
        return problemas

    existentes = sorted(f for f in os.listdir(pasta_bd) if f.lower().endswith(".xlsx"))

    for coluna, arquivo, col_cod, col_nome in regras:
        caminho = os.path.join(pasta_bd, arquivo)
        if not os.path.exists(caminho):
            problemas.append(f"  - Falta o arquivo '{arquivo}' (necessário para a coluna '{coluna}').")
            continue
        cols = [c.strip() if isinstance(c, str) else c
                for c in pd.read_excel(caminho, nrows=0).columns]
        ausentes = [c for c in (col_cod, col_nome) if c not in cols]
        if ausentes:
            problemas.append(
                f"  - Em '{arquivo}', não encontrei a(s) coluna(s) {ausentes}. Cabeçalho: {cols}"
            )

    if problemas:
        problemas.append(f"  > Arquivos .xlsx encontrados na pasta: {existentes or '(nenhum)'}")
    return problemas


def carregar_mapa(caminho, col_codigo, col_nome):
    """Lê uma tabela de referência e devolve {codigo: nome} (com variações de tipo)."""
    ref = _normaliza_colunas(pd.read_excel(caminho))
    ref = ref[[col_codigo, col_nome]].dropna(subset=[col_codigo])
    mapa = {}
    for codigo, nome in zip(ref[col_codigo], ref[col_nome]):
        if isinstance(nome, str):
            nome = nome.strip()
        for chave in _chaves(codigo):
            mapa.setdefault(chave, nome)
    return mapa


def aplicar(serie, mapa):
    """Troca cada código pelo nome; valores sem correspondência ficam como estão."""
    def trocar(valor):
        if pd.isna(valor):
            return valor
        for chave in _chaves(valor):
            if chave in mapa:
                return mapa[chave]
        return valor
    return serie.map(trocar)


def main():
    # Mostra onde o script está procurando (ajuda a conferir caminhos)
    print(f"Pasta de referência: {PASTA_BD}")
    print(f"Tabela principal:    {ARQUIVO_ENTRADA}\n")

    # 1) Valida as regras antes de tudo
    problemas = validar_regras(REGRAS, PASTA_BD)
    if problemas:
        print("Não foi possível processar. Resolva os pontos abaixo:")
        print("\n".join(problemas))
        print("\nNenhum arquivo foi gerado.")
        return

    # 2) Confere a tabela principal
    if not os.path.exists(ARQUIVO_ENTRADA):
        print(f"Não foi possível processar: a tabela principal não foi encontrada em {ARQUIVO_ENTRADA}")
        print("Nenhum arquivo foi gerado.")
        return

    # 3) Lê a tabela principal
    df = _normaliza_colunas(pd.read_excel(ARQUIVO_ENTRADA))

    # 4) Seleciona as colunas desejadas (se a lista não estiver vazia)
    if COLUNAS_DESEJADAS:
        presentes = [c for c in COLUNAS_DESEJADAS if c in df.columns]
        faltando = [c for c in COLUNAS_DESEJADAS if c not in df.columns]
        if faltando:
            print(f"Aviso: colunas não encontradas na tabela principal (ignoradas): {faltando}")
        df = df[presentes].copy()

    # 5) Aplica as substituições lendo de cada tabela
    for coluna, arquivo, col_cod, col_nome in REGRAS:
        if coluna in df.columns:
            mapa = carregar_mapa(os.path.join(PASTA_BD, arquivo), col_cod, col_nome)
            df[coluna] = aplicar(df[coluna], mapa)
        else:
            print(f"Aviso: a coluna '{coluna}' tem regra de substituição, "
                  f"mas não existe na tabela principal.")

    # 6) Renomeia as colunas e salva
    df = df.rename(columns=RENOMEAR)
    df.to_excel(ARQUIVO_SAIDA, index=False)
    print(f"Arquivo gerado: {ARQUIVO_SAIDA} ({len(df)} linhas)")


if __name__ == "__main__":
    main()