import pytest
import os
import pandas as pd
from processor import ParticipantesProcessor

@pytest.fixture
def setup_excel(tmp_path):
    df = pd.DataFrame({
        'IDProjeto': [1, 1, 2],
        'Nome': ['Alice', 'Bob', 'Charlie']
    })
    path = tmp_path / "participantes.xlsx"
    df.to_excel(path, index=False)
    return path

def test_processamento(setup_excel, tmp_path):
    pasta_json = tmp_path / "projetos_exportados_json"
    pasta_json_limpos = tmp_path / "projetos_exportados_json_limpos"

    processor = ParticipantesProcessor(setup_excel, str(pasta_json), str(pasta_json_limpos))
    
    # Testar exportação de JSONs
    processor.exportar_jsons_por_projeto()
    assert os.path.exists(pasta_json)
    assert len(os.listdir(pasta_json)) == 2  # Dois projetos, dois arquivos JSON

    # Testar limpeza dos JSONs exportados
    processor.limpar_jsons_exportados()
    assert os.path.exists(pasta_json_limpos)
    assert len(os.listdir(pasta_json_limpos)) == 2  # Dois arquivos limpos devem ser criados