import pandas as pd
import os
import json
import glob
import logging

logger = logging.getLogger(__name__)

class ParticipantesProcessor:
    def __init__(self, arquivo_excel, pasta_saida_json, pasta_saida_json_limpos):
        self.arquivo_excel = arquivo_excel
        self.pasta_saida_json = pasta_saida_json
        self.pasta_saida_json_limpos = pasta_saida_json_limpos
        self.df = None
        self.dataframes_por_projeto = {}

    def carregar_excel(self):
        if not os.path.exists(self.arquivo_excel):
            raise FileNotFoundError(f"Arquivo '{self.arquivo_excel}' não encontrado.")
        self.df = pd.read_excel(self.arquivo_excel)
        if 'IDProjeto' not in self.df.columns:
            raise ValueError("A coluna 'IDProjeto' não foi encontrada no arquivo.")
        logger.info(f"--- DataFrame carregado. {len(self.df)} linhas encontradas. ---")

    def separar_por_projeto(self):
        ids_projetos = self.df['IDProjeto'].unique()
        for id_projeto in ids_projetos:
            df_projeto = self.df[self.df['IDProjeto'] == id_projeto].copy()
            self.dataframes_por_projeto[id_projeto] = df_projeto
        logger.info(f"\n{len(self.dataframes_por_projeto)} DataFrames separados por projeto.")

    def salvar_jsons(self):
        if not os.path.exists(self.pasta_saida_json):
            os.makedirs(self.pasta_saida_json)
        for id_projeto, df_projeto in self.dataframes_por_projeto.items():
            df_para_json = df_projeto.fillna('')
            for col in df_para_json.select_dtypes(include=['datetime64']).columns:
                df_para_json[col] = df_para_json[col].apply(
                    lambda x: x.strftime('%Y-%m-%dT%H:%M:%S') if pd.notna(x) and x != '' else None
                )
            json_path = os.path.join(self.pasta_saida_json, f'projeto_{id_projeto}.json')
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(df_para_json.to_dict(orient='records'), f, indent=4, ensure_ascii=False)
        logger.info(f"\nJSONs salvos em: '{self.pasta_saida_json}'.\n")


    def exportar_jsons_por_projeto(self):
        try:
            self.carregar_excel()
            self.separar_por_projeto()
            self.salvar_jsons()
        except Exception as e:
            logger.error(f">>> Erro ao exportar JSONs por projeto: {e}")
    
    def limpar_jsons_exportados(self):
        if not os.path.exists(self.pasta_saida_json):
            raise FileNotFoundError(f"Pasta de entrada '{self.pasta_saida_json}' não encontrada.")

        if not os.path.exists(self.pasta_saida_json_limpos):
            os.makedirs(self.pasta_saida_json_limpos)

        arquivos_json = [f for f in os.listdir(self.pasta_saida_json) if f.endswith('.json')]

        if not arquivos_json:
            logger.warning("\nNenhum arquivo JSON encontrado.")
            return

        COLUNAS_DESEJADAS = [
            "ID", "nomeCompleto", "genero", "racaEtnia", "dataInscricaoTxt",
            "dataNascimentoTxt", "telefoneParticipante", "telParticipanteWhatsapp",
            "telParticipanteInternet", "statusDosEstudos", "escolaridadeTxt", "periodoEstudo",
            "escolaParticipante", "nomeResponsavel", "grauParentescoResponsavelPrincipal",
            "telefoneResponsavel", "endereco", "SCtxt", "nomeResponsavelSecundario",
            "grauPrarentescoResponsavelSecundario", "telefoneResponsavelSecundario",
            "reportadoPGDPP", "IDTurma", "IDComunidade", "comunidade", "IDFY", "IDProjeto",
            "IDCidade", "IDPU", "IDTrimestre", "IDaogd", "desistente", "motivoDesistencia",
            "dataDesistencia", "termoConsentimentoAssinado", "termoImagemAssinado"
        ]

        for nome_arquivo in arquivos_json:
            try:
                caminho_entrada = os.path.join(self.pasta_saida_json, nome_arquivo)
                caminho_saida = os.path.join(self.pasta_saida_json_limpos, nome_arquivo)

                df_temp = pd.read_json(caminho_entrada)
                df_limpo = df_temp.fillna('')

                # Filtrar colunas desejadas
                df_filtrado = df_limpo[COLUNAS_DESEJADAS]
                
                # Substituir valores vazios na coluna 'dataDesistencia' por '2000-01-01' de forma segura
                df_filtrado.loc[:, 'dataDesistencia'] = df_filtrado['dataDesistencia'].replace('', '2000-01-01')

                with open(caminho_saida, 'w', encoding='utf-8') as f:
                    json.dump(df_filtrado.to_dict(orient='records'), f, ensure_ascii=False)

                logger.info(f"Arquivo limpo salvo: {caminho_saida}")
            except Exception as e:
                logger.error(f"\nErro ao processar '{nome_arquivo}': {e}")

        logger.info(f"\nTodos os arquivos limpos salvos em '{self.pasta_saida_json_limpos}'.")


class FrequenciaProcessor:
    def __init__(self, arquivo_excel_frequencia, pasta_saida_json):
        self.arquivo_excel_frequencia = arquivo_excel_frequencia
        self.pasta_saida_json = pasta_saida_json
        self.df_frequencia = None
        self.dataframes_por_projeto = {}

    def carregar_excel(self):
        if not os.path.exists(self.arquivo_excel_frequencia):
            raise FileNotFoundError(f"Arquivo '{self.arquivo_excel_frequencia}' não encontrado.")
            
        self.df_frequencia = pd.read_excel(self.arquivo_excel_frequencia)
        
        # <<< CORREÇÃO DEFINITIVA ADICIONADA AQUI >>>
        # Limpa todos os nomes de colunas de espaços em branco no início e no fim.
        # Ex: " IDjson " se tornará "IDjson"
        self.df_frequencia.columns = self.df_frequencia.columns.str.strip()
        
        # Para diagnóstico, você pode descomentar a linha abaixo para ver os nomes das colunas
        # logger.info(f"Nomes das colunas após a limpeza: {list(self.df_frequencia.columns)}")

        if 'IDProjeto' not in self.df_frequencia.columns:
            raise ValueError("A coluna 'IDProjeto' não foi encontrada no arquivo.")
            
        logger.info(f"DataFrame de frequência carregado. {len(self.df_frequencia)} linhas encontradas.")

    def separar_por_projeto(self):
        self.df_frequencia['IDProjeto'] = self.df_frequencia['IDProjeto'].astype(str)
        ids_projetos = self.df_frequencia['IDProjeto'].unique()
        
        for id_projeto in ids_projetos:
            df_projeto = self.df_frequencia[self.df_frequencia['IDProjeto'] == id_projeto].copy()
            self.dataframes_por_projeto[id_projeto] = df_projeto
        logger.info(f"{len(self.dataframes_por_projeto)} DataFrames separados por IDProjeto.")

    def salvar_jsons(self):
        os.makedirs(self.pasta_saida_json, exist_ok=True)
        colunas_desejadas = [
            "ID", "Criado", "IDParticipante", "IDOficina", "IDFY", "IDProjeto",
            "IDEducadora", "IDTurma", "dataOficina", "dataReposicao",
            "presenca", "atividadeReposicao", "IDGrupoParticipacao", "IDjson"
        ]

        for id_projeto, df_projeto in self.dataframes_por_projeto.items():
            if 'IDjson' not in df_projeto.columns:
                df_projeto['IDjson'] = ""
            
            colunas_existentes = [col for col in colunas_desejadas if col in df_projeto.columns]
            df_para_json = df_projeto[colunas_existentes].copy()

            # 🔹 Preencher valores vazios em dataReposicao
            if 'dataReposicao' in df_para_json.columns:
                df_para_json['dataReposicao'] = df_para_json['dataReposicao'].fillna(pd.Timestamp('2000-01-01'))

            for col in df_para_json.select_dtypes(include=['datetime64[ns]']).columns:
                df_para_json[col] = df_para_json[col].apply(
                    lambda x: x.strftime('%Y-%m-%dT%H:%M:%S') if pd.notnull(x) else ""
                )
            
            df_para_json = df_para_json.where(pd.notnull(df_para_json), '')

            json_path = os.path.join(self.pasta_saida_json, f'projeto_{id_projeto}.json')
            records = df_para_json.to_dict(orient='records')
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(records, f, ensure_ascii=False)
                
            logger.info(f"Arquivo JSON salvo: {json_path}")

    def processar_frequencias(self):
        try:
            self.carregar_excel()
            self.separar_por_projeto()
            self.salvar_jsons()
        except Exception as e:
            logger.error(f"Erro ao processar frequências por projeto: {e}")