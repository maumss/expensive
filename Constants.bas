' Constantes globais

Public Const BROKER As String = "Broker"
Public Const RESERVA_ESTRATEGICA As String = "Reserva estratégica"

' Planilha Intro
Public Const RANGE_DATA_ULTIMA_ATUALIZ As String = "E12"
Public Const RANGE_POSICAO = "intPosicao"

' Planilha Orcamento
Public Const RANGE_TOLERANCIA As String = "orcTolerancia"
Public Const TIPO_LANCAMENTO_INVESTIMENTOS As String = "investimentos"

' Planilha Alocacao
Public Const RANGE_CELULA_INICIO_ADHOC As String = "C95"
Public Const RANGE_CELULA_FIM_ADHOC As String = "C99"
Public Const RANGE_CELULA_INICIO_PORTFOLIO As String = "C36"
Public Const RANGE_CELULA_FIM_PORTFOLIO As String = "C75"

' Planilha Retorno
Public Const RANGE_PLAN_FECHADA = "retPlanFechada"

' Planilha Retrato
Public Const RANGE_RELAT_RETRAT = "$A$1:$Q$197"

' Planilhas Jan a Dez - geral
Public Const RANGE_SITUAC_PLANILHA As String = "E4"
Public Const RANGE_DATA_POSICAO As String = "N4"
Public Const SITUAC_ABERTO As String = "Aberto"
Public Const SITUAC_FECHADO As String = "Fechado"
Public Const RANGE_SALDO_MES As String = "B2"
Public Const NOME_PLAN_DEZ As String = "Dez"

' Planilhas Jan a Dez - Conta corretora
Public Const RANGE_SALDO_CONTA_XP As String = "B22"
Public Const RANGE_SALDO_CONTA_AVENUE_USD_TOTAL As String = "B26"
Public Const RANGE_SALDO_CONTA_AVENUE_USD_DO_BRASIL As String = "B28"
Public Const RANGE_SALDO_CONTA_AVENUE_BR As String = "B30"

' Planilhas Jan a Dez - movimentações
Public Const RANGE_HEADER_MOVIMENTACAO As String = "D14"
Public Const RANGE_HEADER_DATA_MOVIMENTACAO As String = "D15"
Public Const RANGE_PRIMEIRA_DATA_MOVIMENTACAO As String = "D16"
Public Const RANGE_HEADER_DESC_MOVIMENTACAO As String = "E15"
Public Const RANGE_HEADER_TIPO_MOVIMENTACAO As String = "F15"
Public Const RANGE_HEADER_VALOR_MOVIMENTACAO As String = "G15"
Public Const RANGE_COLUNA_DATA_MOVIMENTACAO As String = "D16:D96"
Public Const RANGE_COLUNA_VALOR_MOVIMENTACAO As String = "G16:G96"
Public Const RANGE_TAB_MOVIMENTACAO As String = "D16:G96"

' Planilhas Jan a Dez - cartões
Public Const RANGE_HEADER_CARTOES As String = "J14"
Public Const RANGE_PRIMEIRA_DATA_CARTOES As String = "J16"
Public Const RANGE_COLUNA_DATA_CARTOES As String = "J16:J96"
Public Const RANGE_COLUNA_VALOR_CARTOES As String = "N16:N96"
Public Const RANGE_ULTIMO_VALOR_CARTAO As String = "N96"
Public Const RANGE_TAB_CARTOES As String = "J16:N96"

' Planilhas Jan a Dez - Portfólio
Public Const RANGE_COLUNA_ATIVO_PORTFOLIO As String = "D104:D142"
Public Const RANGE_COLUNA_SALDO_INICIAL_PORTFOLIO As String = "F104:F142"
Public Const RANGE_COLUNA_AUXILIAR_REND_PORTFOLIO As String = "I104:I142"
Public Const RANGE_COLUNA_SALDO_FINAL_PORTFOLIO As String = "N104:N142"

Public Const RANGE_AREA_RELATORIO As String = "C102:N144"

' Planilhas Jan a Dez - Carteira Ações
Public Const RANGE_COLUNA_DATA_ACOES As String = "D147:D176"
Public Const RANGE_COLUNA_ATIVO_ACOES As String = "E147:E176"
Public Const RANGE_COLUNA_QTDE_ACOES As String = "G147:G176"
Public Const RANGE_COLUNA_SALDO_INICIAL_ACOES As String = "F147:F176"
Public Const RANGE_COLUNA_SALDO_FINAL_ACOES As String = "N147:N176"
Public Const RANGE_CELULA_TRIBUTA_ACOES As String = "Q177"
Public Const RANGE_COLUNA_RESULTADO_COMUM_ACOES As String = "X147:X176"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_ACOES As String = "AB147:AB176"
Public Const RANGE_TAB_ACOES As String = "D147:N176"

' Planilhas Jan a Dez - Carteira Fundos Imobiliários
Public Const RANGE_COLUNA_DATA_FII As String = "D183:D212"
Public Const RANGE_COLUNA_ATIVO_FII As String = "E183:E212"
Public Const RANGE_COLUNA_QTDE_FII As String = "G183:G212"
Public Const RANGE_COLUNA_SALDO_INICIAL_FII As String = "F183:F212"
Public Const RANGE_COLUNA_SALDO_FINAL_FII As String = "N183:N212"
Public Const RANGE_COLUNA_RESULTADO_COMUM_FII As String = "X183:X212"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_FII As String = "AB183:AB212"
Public Const RANGE_TAB_FII As String = "D183:N212"

' Planilhas Jan a Dez - Carteira Tesouro Direto RF
Public Const RANGE_COLUNA_DATA_TESOURO_DIRETO As String = "D219:D233"
Public Const RANGE_COLUNA_ATIVO_TESOURO_DIRETO As String = "E219:E233"
Public Const RANGE_COLUNA_QTDE_TESOURO_DIRETO As String = "G219:G233"
Public Const RANGE_COLUNA_SALDO_INICIAL_TESOURO_DIRETO As String = "F219:F233"
Public Const RANGE_COLUNA_SALDO_FINAL_TESOURO_DIRETO As String = "N219:N233"

' Planilhas Jan a Dez - Carteira Tesouro Direto Selic
Public Const RANGE_COLUNA_DATA_TESOURO_SELIC As String = "D240:D245"
Public Const RANGE_COLUNA_ATIVO_TESOURO_SELIC As String = "E240:E245"
Public Const RANGE_COLUNA_QTDE_TESOURO_SELIC As String = "G240:G245"
Public Const RANGE_COLUNA_SALDO_INICIAL_TESOURO_SELIC As String = "F240:F245"
Public Const RANGE_COLUNA_SALDO_FINAL_TESOURO_SELIC As String = "N240:N245"

' Planilhas Jan a Dez - Carteira ETF
Public Const RANGE_COLUNA_DATA_ETF As String = "D252:D258"
Public Const RANGE_COLUNA_ATIVO_ETF As String = "E252:E258"
Public Const RANGE_COLUNA_QTDE_ETF As String = "G252:G258"
Public Const RANGE_COLUNA_SALDO_INICIAL_ETF As String = "F252:F258"
Public Const RANGE_COLUNA_SALDO_FINAL_ETF As String = "N252:N258"
Public Const RANGE_COLUNA_RESULTADO_COMUM_ETF As String = "X252:X258"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_ETF As String = "AB252:AB258"

' Planilhas Jan a Dez - Carteira Ações USD
Public Const RANGE_COLUNA_DATA_STOCK As String = "D265:D294"
Public Const RANGE_COLUNA_ATIVO_STOCK As String = "E265:E294"
Public Const RANGE_COLUNA_QTDE_STOCK As String = "G265:G294"
Public Const RANGE_COLUNA_SALDO_INICIAL_STOCK As String = "F265:F294"
Public Const RANGE_COLUNA_SALDO_FINAL_STOCK As String = "N265:N294"
Public Const RANGE_CELULA_TRIBUTA_STOCK As String = "Q295"
Public Const RANGE_COLUNA_RESULTADO_COMUM_STOCK As String = "X265:X294"
Public Const RANGE_TAB_STOCK As String = "D265:N294"

' Planilhas Jan a Dez - Carteira REIT
Public Const RANGE_COLUNA_DATA_REIT As String = "D301:D330"
Public Const RANGE_COLUNA_ATIVO_REIT As String = "E301:E330"
Public Const RANGE_COLUNA_QTDE_REIT As String = "G301:G330"
Public Const RANGE_COLUNA_SALDO_INICIAL_REIT As String = "F301:F330"
Public Const RANGE_COLUNA_SALDO_FINAL_REIT As String = "N301:N330"
Public Const RANGE_CELULA_TRIBUTA_REIT As String = "Q331"
Public Const RANGE_COLUNA_RESULTADO_COMUM_REIT As String = "X301:X330"
Public Const RANGE_TAB_REIT As String = "D301:N330"

' Planilhas Jan a Dez - Carteira Treasuries
Public Const RANGE_COLUNA_DATA_TREASURY As String = "D337:D346"
Public Const RANGE_COLUNA_ATIVO_TREASURY As String = "E337:E346"
Public Const RANGE_COLUNA_QTDE_TREASURY As String = "G337:G346"
Public Const RANGE_COLUNA_SALDO_INICIAL_TREASURY As String = "F337:F346"
Public Const RANGE_COLUNA_SALDO_FINAL_TREASURY As String = "N337:N346"
Public Const RANGE_COLUNA_RESULTADO_COMUM_TREASURY As String = "X337:X346"

' Planilhas Jan a Dez - Carteira Ouro
Public Const RANGE_COLUNA_DATA_OURO As String = "D353:D362"
Public Const RANGE_COLUNA_ATIVO_OURO As String = "E353:E362"
Public Const RANGE_COLUNA_QTDE_OURO As String = "G353:G362"
Public Const RANGE_COLUNA_SALDO_INICIAL_OURO As String = "F353:F362"
Public Const RANGE_COLUNA_SALDO_FINAL_OURO As String = "N353:N362"
Public Const RANGE_COLUNA_RESULTADO_COMUM_OURO As String = "X353:X362"

' Planilhas Jan a Dez - Carteira Cripto
Public Const RANGE_COLUNA_DATA_CRIPTO As String = "D369:D378"
Public Const RANGE_COLUNA_ATIVO_CRIPTO As String = "E369:E378"
Public Const RANGE_COLUNA_QTDE_CRIPTO As String = "G369:G378"
Public Const RANGE_COLUNA_SALDO_INICIAL_CRIPTO As String = "F369:F378"
Public Const RANGE_COLUNA_SALDO_FINAL_CRIPTO As String = "N369:N378"
Public Const RANGE_COLUNA_RESULTADO_COMUM_CRIPTO As String = "X369:X378"
Public Const RANGE_CELULA_IGNORA_AGENDA_CRIPTO As String = "R381"

' Planilhas Jan. a Dez - indicadores
Public Const RANGE_COLUNA_DESCR_INDICADORES As String = "D388:D396"
Public Const RANGE_COLUNA_MES_INDICADORES As String = "F388:F396"
Public Const RANGE_COLUNA_ANO_INDICADORES As String = "H388:H396"
Public Const RANGE_COLUNA_DOZE_MESES_INDICADORES As String = "I388:I396"
Public Const RANGE_CELULA_DOLAR_FINAL_MES As String = "G393"
Public Const SP500 As String = "S&P 500"
Public Const RANGE_CELULA_DOLAR_BACEN_COMPRA As String = "F400"
Public Const RANGE_CELULA_DOLAR_BACEN_VENDA As String = "G400"

' Planilha Acoes
Public Const RANGE_TAB_ACOES_365 As String = "B7:J146"
Public Const RANGE_COLUNA_TICKETS As String = "E7:E146"
Public Const RANGE_COLUNA_PRECO As String = "G7:G146"
Public Const RANGE_TAB_MOEDAS_365 As String = "C151:D153"
Public Const RANGE_CELULA_DOLAR As String = "D151"
Public Const RANGE_CELULA_BTC As String = "D152"
Public Const RANGE_CELULA_ETH As String = "D153"

' Planilha TesouroDireto
Public Const RANGE_DATA_ATUALIZ_TD = "H2"
Public Const RANGE_COLUNA_TICKETS_TD As String = "F2:F36"
Public Const RANGE_COLUNA_PRECO_TD As String = "D2:D36"
