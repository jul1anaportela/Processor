# Projeto Processor

<b>Projeto processor </b>é uma ferramenta em Python para processar um aruqivo Excel com dados de participantes, separando os dados por projeto (com base na coluna `IDProjeto` ou `IDTurma`) e exportando arquivos `.json` individuas por projeto. Também realiza uma etapa de "limpeza" dos arquivos exportados para padronização e compatibilidade com integrações.

## Recursos
* Leitura de planilha Excel
* Separar dados por `IDProjeto` ou `IDTurma`
* Exporta para arquivos JSON (um por um grupo)
* Conversão de datas para formato ISO (`%Y-%m-%dT%H:%M:%S`)
* Limpeza de valores `NaN` ou `null`
* Logging e tratamento de erros
* Testes automatizados com Pytest

## Componentes

### 📁 ParticipantesProcessor
Processa uma planilha de participantes, separando os dados por `IDProjeto`

**Funcionalidades**
- Carrega um arquivo Excecl de participantes
- Agrupa os dados por `IDProjeto`
- Salva arquivos `.json` individuais por projeto
- Gera uma cópia dos JSONs com limpeza de valores nulos

**Arquivos esperados:**
- Entrada: `[main]_participantesPus_json.xlsx`
- Saídas:
    - JSONs por projeto: `projetos_exportados_json/`
    - JSONs limpos: `projetos_exportados_json_limpos/`

### 📁 FrequenciaProcessor
Processa uma planilha de frequencia, separando os dados por `IDTurma`

**Funcionalidades**
- Carrega um arquivo Excel ded frequencia (`[main]_frequenciaPUs_json.xlsx`)
- Agrupa os dados por `IDTurma`
- Salva arquivos `.json` individuais por turma

**Arquivos esperados:**
- Entrada: `[main]_frequenciaPUs_json.xlsx`
- Saída: `frequencias_exportadas_json/`

## Requisitos
* Python 3.8+
* [pandas](https://pandas.pydata.org)
* [openpyxl](https://openpyxl.readthedocs.io)
* [pytest](https://docs.pytest.org)

## Instalação
Clone o repositório e instale com:

```bash
pip install .
```

## Uso

Após a instalação, execute:
```bash
processar-participantes
```
Por padrão, o script espera os seguintes caminhos:
* Arquivo Excel: `[main]_participantesPus_json.xlsx`
* Saída JSON: `projetos_exportados_json/`
* Saída JSON limpos: `projetos_exportados_json_limpos/`
</br>

Esses caminhos podem ser modificados no `main.py`

## Uso adicional com FrequenciaProcessor
Para processar frequencias por turma:
```py
from processor import FrequenciaProcessor

processor = FrequenciaProcessor(
    arquivo_excel_frequencia='[main]_frequenciaPUs_json.xlsx',
    pasta_saida_json='frequencias_exportadas_json'
)
processor.processar_frequencias()

```

## Estrutura do Projeto
```
projeto_processor/
├── __init__.py
├── main.py
├── processor.py
├── tests/
│   └── test_processor.py
└── setup.py
```

## Executando os Testes
```bash
pytest
```

## Autora

Juliana Portela

## Licença

Este projeto está licenciado sob os termos da licença MIT