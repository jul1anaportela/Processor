"""
Filtro interativo de participantes (segundo passo).

Lê a planilha já tratada (com nomes no lugar dos códigos) e pergunta, em cascata:
   1) qual LOCALIDADE filtrar;
   2) qual PROJETO daquela localidade;
   3) quais TURMAS (uma, várias ou todas).
No final, gera uma planilha só com as linhas que você escolheu.

COMO USAR:
  1. pip install pandas openpyxl
  2. Confira o caminho em ARQUIVO_ENTRADA (já preenchido com o seu).
  3. python filtro_participantes.py
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


def sanitizar(texto):
    """Remove caracteres inválidos para nome de arquivo no Windows."""
    return re.sub(r'[\\/:*?"<>|]', "-", str(texto)).strip()


# ============================ PROGRAMA ============================

def main():
    print("=== FILTRO DE PARTICIPANTES ===\n")

    if not os.path.exists(ARQUIVO_ENTRADA):
        print(f"Arquivo não encontrado:\n  {ARQUIVO_ENTRADA}")
        print("Confira o caminho em ARQUIVO_ENTRADA, no topo do script.")
        return

    df = pd.read_excel(ARQUIVO_ENTRADA)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    print(f"Lendo: {ARQUIVO_ENTRADA}\n({len(df)} participantes no total)\n")

    # Confere se as colunas de filtro existem
    faltando = [c for c in (COLUNA_LOCALIDADE, COLUNA_PROJETO, COLUNA_TURMA) if c not in df.columns]
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

    # 2) PROJETO (dentro da localidade escolhida)
    projetos = valores_unicos(df_loc, COLUNA_PROJETO)
    if not projetos:
        print("Não há projetos para essa localidade.")
        return
    print(f"Projetos em {localidade}:")
    projeto = escolher_um(projetos, "Escolha o Projeto (número): ")
    df_proj = df_loc[_norm(df_loc[COLUNA_PROJETO]) == projeto]
    print(f"\n>> Projeto: {projeto}  ({len(df_proj)} participantes)\n")

    # 3) TURMAS (dentro do projeto escolhido)
    turmas = valores_unicos(df_proj, COLUNA_TURMA)
    if not turmas:
        print("Não há turmas para esse projeto.")
        return
    print(f"Turmas em {projeto}:")
    escolhidas = escolher_varios(turmas, "Escolha as turmas (números separados por vírgula, ou 0 para todas): ")

    if escolhidas is None:
        df_final = df_proj
        rotulo_turma = "todas"
    else:
        df_final = df_proj[_norm(df_proj[COLUNA_TURMA]).isin(escolhidas)]
        rotulo_turma = escolhidas[0] if len(escolhidas) == 1 else f"{len(escolhidas)}-turmas"
        print(f"\n>> Turmas selecionadas: {', '.join(escolhidas)}")

    # 4) SALVA O RESULTADO
    nome_saida = f"02[filtro]_{sanitizar(localidade)}_{sanitizar(projeto)}_{sanitizar(rotulo_turma)}.xlsx"
    caminho_saida = os.path.join(PASTA_SAIDA, nome_saida)
    df_final.to_excel(caminho_saida, index=False)

    print(f"\nPronto! {len(df_final)} participante(s) no filtro.")
    print(f"Arquivo gerado:\n  {caminho_saida}")


if __name__ == "__main__":
    main()