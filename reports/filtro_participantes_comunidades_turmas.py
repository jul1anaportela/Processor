"""
Filtro interativo de participantes — DETALHADO (4 níveis).

Pergunta, em cascata:
   1) qual LOCALIDADE filtrar;
   2) qual PROJETO daquela localidade;
   3) quais COMUNIDADES (uma, várias ou todas);
   4) quais TURMAS, entre as que existem nas comunidades escolhidas (uma, várias ou todas).
No final, gera uma planilha só com as linhas que você escolheu.

COMO USAR:
  1. pip install pandas openpyxl
  2. Confira o caminho em ARQUIVO_ENTRADA (já preenchido com o seu).
  3. python filtro_comunidades_turmas.py
     e responda às perguntas digitando os números.
"""

import os
import re
import pandas as pd

# ============================ CONFIGURAÇÃO ============================

# Planilha tratada (saída do primeiro script).
ARQUIVO_ENTRADA = r"C:\Users\JulianaPortela\Desktop\projeto_processor\reports\arquivos_reports\01[main]_participantes_tratado.xlsx"

# A planilha filtrada é salva na subpasta "filtros", dentro de "arquivos_reports".
PASTA_SAIDA = os.path.join(os.path.dirname(ARQUIVO_ENTRADA), "filtros")

# Nomes das colunas usadas nos filtros (ajuste se você renomeou diferente).
COLUNA_LOCALIDADE = "PU"
COLUNA_PROJETO = "Projeto"
COLUNA_COMUNIDADE = "Comunidade"
COLUNA_TURMA = "Turmas"


# ============================ FUNÇÕES AUXILIARES ============================

def _norm(serie):
    """Versão em texto, sem espaços nas pontas, para exibir/comparar com segurança."""
    return serie.astype(str).str.strip()


def valores_unicos(df, coluna):
    """Lista ordenada dos valores distintos de uma coluna (ignora vazios)."""
    return sorted(v for v in _norm(df[coluna]).unique() if v not in ("", "nan", "None"))


def mostrar_menu(opcoes):
    for i, opcao in enumerate(opcoes, start=1):
        print(f"  [{i}] {opcao}")


def escolher_um(opcoes, pergunta):
    """Mostra o menu e devolve a opção escolhida (um único item)."""
    mostrar_menu(opcoes)
    while True:
        resposta = input(pergunta).strip()
        if resposta.isdigit() and 1 <= int(resposta) <= len(opcoes):
            return opcoes[int(resposta) - 1]
        print(f"  Entrada inválida. Digite um número de 1 a {len(opcoes)}.")


def escolher_varios(opcoes, pergunta):
    """Devolve a lista de opções escolhidas, ou None se o usuário pedir TODAS (0)."""
    mostrar_menu(opcoes)
    print("  [0] TODAS")
    while True:
        resposta = input(pergunta).strip()
        if resposta == "0":
            return None
        partes = [p.strip() for p in resposta.split(",") if p.strip()]
        if partes and all(p.isdigit() and 1 <= int(p) <= len(opcoes) for p in partes):
            indices = sorted({int(p) for p in partes})
            return [opcoes[i - 1] for i in indices]
        print(f"  Entrada inválida. Use números de 1 a {len(opcoes)} separados por "
              f"vírgula, ou 0 para todas.")


def sanitizar(texto, limite=40):
    """Remove caracteres inválidos para nome de arquivo no Windows e limita o tamanho."""
    limpo = re.sub(r'[\\/:*?"<>|]', "-", str(texto)).strip()
    return limpo[:limite].strip()


def filtrar_multi(df, coluna, escolhidas):
    """Aplica o filtro de múltipla escolha (None = todas) e devolve (df_filtrado, rótulo)."""
    if escolhidas is None:
        return df, "todas"
    df_filtrado = df[_norm(df[coluna]).isin(escolhidas)]
    rotulo = escolhidas[0] if len(escolhidas) == 1 else f"{len(escolhidas)}-itens"
    return df_filtrado, rotulo


# ============================ PROGRAMA ============================

def main():
    print("=== FILTRO DE PARTICIPANTES (DETALHADO) ===\n")

    if not os.path.exists(ARQUIVO_ENTRADA):
        print(f"Arquivo não encontrado:\n  {ARQUIVO_ENTRADA}")
        print("Confira o caminho em ARQUIVO_ENTRADA, no topo do script.")
        return

    df = pd.read_excel(ARQUIVO_ENTRADA)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    print(f"Lendo: {ARQUIVO_ENTRADA}\n({len(df)} participantes no total)\n")

    # Confere se as colunas de filtro existem
    necessarias = (COLUNA_LOCALIDADE, COLUNA_PROJETO, COLUNA_COMUNIDADE, COLUNA_TURMA)
    faltando = [c for c in necessarias if c not in df.columns]
    if faltando:
        print(f"Não encontrei a(s) coluna(s) {faltando} na planilha.")
        print(f"Colunas disponíveis: {list(df.columns)}")
        return

    # 1) LOCALIDADE
    localidades = valores_unicos(df, COLUNA_LOCALIDADE)
    if not localidades:
        print("Não há valores de Localidade para filtrar.")
        return
    print("Localidades disponíveis:")
    localidade = escolher_um(localidades, "Escolha a Localidade (número): ")
    df_loc = df[_norm(df[COLUNA_LOCALIDADE]) == localidade]
    print(f"\n>> Localidade: {localidade}  ({len(df_loc)} participantes)\n")

    # 2) PROJETO (dentro da localidade)
    projetos = valores_unicos(df_loc, COLUNA_PROJETO)
    if not projetos:
        print("Não há projetos para essa localidade.")
        return
    print(f"Projetos em {localidade}:")
    projeto = escolher_um(projetos, "Escolha o Projeto (número): ")
    df_proj = df_loc[_norm(df_loc[COLUNA_PROJETO]) == projeto]
    print(f"\n>> Projeto: {projeto}  ({len(df_proj)} participantes)\n")

    # 3) COMUNIDADES (dentro do projeto)
    comunidades = valores_unicos(df_proj, COLUNA_COMUNIDADE)
    if not comunidades:
        print("Não há comunidades para esse projeto.")
        return
    print(f"Comunidades em {projeto}:")
    escolha_com = escolher_varios(comunidades, "Escolha as comunidades (números separados por vírgula, ou 0 para todas): ")
    df_com, rotulo_com = filtrar_multi(df_proj, COLUNA_COMUNIDADE, escolha_com)
    if escolha_com is not None:
        print(f"\n>> Comunidades selecionadas: {', '.join(escolha_com)}  ({len(df_com)} participantes)\n")
    else:
        print(f"\n>> Comunidades: todas  ({len(df_com)} participantes)\n")

    # 4) TURMAS (entre as comunidades escolhidas)
    turmas = valores_unicos(df_com, COLUNA_TURMA)
    if not turmas:
        print("Não há turmas para essa combinação.")
        return
    print(f"Turmas disponíveis:")
    escolha_turma = escolher_varios(turmas, "Escolha as turmas (números separados por vírgula, ou 0 para todas): ")
    df_final, rotulo_turma = filtrar_multi(df_com, COLUNA_TURMA, escolha_turma)
    if escolha_turma is not None:
        print(f"\n>> Turmas selecionadas: {', '.join(escolha_turma)}")

    # 5) SALVA O RESULTADO
    nome_saida = (
        f"03[filtro]_{sanitizar(localidade)}_{sanitizar(projeto)}"
        f"_com-{sanitizar(rotulo_com)}_turma-{sanitizar(rotulo_turma)}.xlsx"
    )
    caminho_saida = os.path.join(PASTA_SAIDA, nome_saida)
    df_final.to_excel(caminho_saida, index=False)

    print(f"\nPronto! {len(df_final)} participante(s) no filtro.")
    print(f"Arquivo gerado:\n  {caminho_saida}")


if __name__ == "__main__":
    main()