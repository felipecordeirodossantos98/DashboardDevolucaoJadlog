# DashboardDevolucao

Aplicação Streamlit independente para o painel **Devolução - TCv2**, com navegação e componentes próprios do Streamlit. A base é carregada pelo campo de importação do app.

## Executar no Linux

```bash
cd /home/felipe/Projects/Jadlog/DashboardDevolucaoJadlog
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
streamlit run app.py
```

O Streamlit abre o painel no navegador. Importe um `.xlsx` pelo campo **Importar base** na barra lateral; o painel fica disponível após o upload e o arquivo enviado permanece ativo na sessão atual. Não há uma base local padrão no projeto. A navegação lateral contém BACKLOG, ENTRADA, CONTROLE CLIENTES, CONTROLE LOJAS, FAIXA DE VALOR e Lista CTE. Os filtros usam os valores presentes na base.

## Conteúdo

- Interface responsiva com navegação lateral, cartões, filtros e tabelas interativas do Streamlit.
- Exibe indicadores, matrizes de aging, gráficos e consulta de CTE calculados a partir da planilha.
- O gráfico DT_EVENTO X QTD tem um seletor de intervalo na linha do tempo para aproximar e afastar a janela visível.
- Identidade Jadlog em uma paleta mais leve, com realces vermelhos.
- CTEs, pedidos e manifestos são carregados como texto para preservar os identificadores.

O carregador também reconhece as colunas da nova exportação, como `DATA_SOLICITACAO`, `OCORRENCIA`, `CLIENTE` e `DATA_ULTIMA_DATA`, e remove linhas de total e observações do Power BI. Quando a fonte não traz `TIPO`, `STATUS_TRATATIVA` ou `CIDADE_ORIGEM`, o app mantém os campos sem inventar classificações; os recortes que dependem dessas dimensões aparecem como indisponíveis. O aging da aba Entrada usa `DATA_ULTIMA_DATA` (normalizada internamente como `DATA_EVENTO`) e calcula os dias até a data de referência do painel.

O filtro `VENCIMENTO`, disponível em todas as abas, classifica datas a partir do dia atual: `No prazo` inclui vencimentos de hoje em diante e `Vencido` inclui datas anteriores. Resumos, matrizes e arquivos XLSX são armazenados em cache para acelerar a navegação e os reruns com os mesmos dados.
