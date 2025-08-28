# Automação de Indicadores OnePage

Este projeto automatiza a geração e envio diário de indicadores de desempenho para uma rede fictícia de 25 lojas de roupas.

## Objetivo
Simular a rotina de análise de dados de uma rede de lojas, calculando indicadores como faturamento, diversidade de produtos e ticket médio, e enviando relatórios automáticos por e-mail para gerentes e diretoria.

## Funcionalidades
- Organiza e trata bases de vendas, lojas e contatos.
- Gera planilhas individuais para cada loja e salva em backup.
- Calcula indicadores diários e anuais, comparando com metas.
- Cria relatórios em HTML personalizados e envia automaticamente via Outlook.
- Consolida rankings de desempenho do dia e do ano para a diretoria.

## Tecnologias
- Python
- Pandas (manipulação de dados)
- Pathlib (gerenciamento de arquivos)
- Win32com.client (integração com Outlook)

## Como usar
1. Tenha Python instalado e as bibliotecas mencionadas.
2. Coloque os arquivos de base (`Emails.xlsx`, `Lojas.csv`, `Vendas.xlsx`) na pasta `Bases de Dados`.
3. Execute o script `Automacao de Processo.ipynb`.
4. Os e-mails serão enviados automaticamente e as planilhas de backup serão salvas em `Backup Arquivos Lojas`.

## Observações
Este projeto é um exemplo de como Python pode ser usado para automação de processos corporativos, aumentando eficiência, confiabilidade e liberando tempo para atividades estratégicas.
