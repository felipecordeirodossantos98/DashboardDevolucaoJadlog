# DashboardDevolucao

Aplicação Streamlit independente para o painel **Devolução - TCv2**, com navegação e componentes próprios do Streamlit. A fonte local é `data.xlsx`, na aba `Export`.

## Executar no Linux

```bash
cd /home/felipe/Projects/Jadlog/DashboardDevolucaoJadlog
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
streamlit run app.py
```

O Streamlit abre o painel no navegador. Mantenha `data.xlsx` junto de `app.py`. A navegação lateral contém BACKLOG, ENTRADA, CONTROLE CLIENTES, CONTROLE LOJAS, FAIXA DE VALOR e Lista CTE. Os filtros usam os valores presentes na base. Ao substituir a planilha, atualize a página do app para recarregar os dados.

## Conteúdo

- Interface responsiva com navegação lateral, cartões, filtros e tabelas interativas do Streamlit.
- Exibe indicadores, matrizes de aging, gráficos e consulta de CTE calculados a partir da planilha.
- O gráfico DT_EVENTO X QTD tem um seletor de intervalo na linha do tempo para aproximar e afastar a janela visível.
- Identidade Jadlog em uma paleta mais leve, com realces vermelhos.
- CTEs, pedidos e manifestos são carregados como texto para preservar os identificadores.

O `data.xlsx` do app contém `DATA_ATUAL` e `CIDADE_ORIGEM` recuperadas do modelo PBIX e associadas à base completa pela CTE; esses campos não existem na exportação original. As matrizes de aging reproduzem o cálculo do PBIX (`DATEDIFF(DATA_ATUAL, TODAY())`) com referência em 08/09/2026. As tabelas `AGING - CUSTODIA` e `AGING ENTRADA - LOJAS E PA'S` usam os filtros de status, cidade, tipo e manifesto dos respectivos visuais. O total combinado de TRAVADO e ENCERRADO conta CTEs únicas, sem duplicar os registros que aparecem nos dois grupos. `DATA_EVENTO` continua vindo de `DT_ENTRADA_REVERSAS`. Se a planilha for substituída por uma exportação sem os campos recuperados, o app usa `DT_ENTRADA_REVERSAS` e `STATUS_TRATATIVA` como aproximações. Como a exportação não traz `FINALIZADO?`, o app considera `ENCERRADO` e `BAIXADO` como finalizados.
