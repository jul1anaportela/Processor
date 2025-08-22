import logging
from processor import ParticipantesProcessor, FrequenciaProcessor

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    # # Processar Participantes
    # arquivo_participantes = "arquivos_baixados_sharepoint\\[main]_participantesPUs_json.xlsx"
    # pasta_json_participantes = "arquivos_tratados\\tb_projetos__tb_participantes\\projetos_exportados_json\\PU"
    # pasta_json_limpos = "arquivos_tratados\\tb_projetos__tb_participantes\\projetos_exportados_json_limpos\\PU"

    # processor_participantes = ParticipantesProcessor(
    #     arquivo_participantes, 
    #     pasta_json_participantes, 
    #     pasta_json_limpos
    # )

    # processor_participantes.exportar_jsons_por_projeto()
    # processor_participantes.limpar_jsons_exportados()

    # Processar Frequencias
    arquivo_frequencia = "arquivos_baixados_sharepoint\\[main]_frequenciaPUs_json.xlsx"
    pasta_saida= "arquivos_tratados\\tb_frequencia\\frequencias_exportadas_json\\PU\\IDjson"


    processor_frequencia = FrequenciaProcessor(
        arquivo_frequencia,
        pasta_saida
    )
    processor_frequencia.processar_frequencias()

    

if __name__ == "__main__":
    main()
    logging.info("Processamento concluído com sucesso.")
