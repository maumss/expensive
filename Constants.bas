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
Public Const RANGE_CELULA_INICIO_ADHOC As String = "C87"
Public Const RANGE_CELULA_FIM_ADHOC As String = "C91"
Public Const RANGE_CELULA_INICIO_PORTFOLIO As String = "C38"
Public Const RANGE_CELULA_FIM_PORTFOLIO As String = "C77"

' Planilha Retrato
Public Const RANGE_RELAT_RETRAT = "$A$1:$Q$175"

' Planilhas Jan. a Dez - geral
Public Const RANGE_SITUAC_PLANILHA As String = "E4"
Public Const RANGE_DATA_POSICAO As String = "N4"
Public Const SITUAC_ABERTO As String = "Aberto"
Public Const SITUAC_FECHADO As String = "Fechado"
Public Const RANGE_SALDO_MES As String = "B2"
Public Const NOME_PLAN_DEZ As String = "Dez"

' Planilhas Jan. a Dez - Conta corretora
Public Const RANGE_SALDO_CONTA_XP As String = "B22"
Public Const RANGE_SALDO_CONTA_AVENUE_USD_TOTAL As String = "B26"
Public Const RANGE_SALDO_CONTA_AVENUE_USD_DO_BRASIL As String = "B28"
Public Const RANGE_SALDO_CONTA_AVENUE_BR As String = "B30"

' Planilhas Jan. a Dez - movimentações
Public Const RANGE_HEADER_MOVIMENTACAO As String = "D14"
Public Const RANGE_HEADER_DATA_MOVIMENTACAO As String = "D15"
Public Const RANGE_PRIMEIRA_DATA_MOVIMENTACAO As String = "D16"
Public Const RANGE_HEADER_DESC_MOVIMENTACAO As String = "E15"
Public Const RANGE_HEADER_TIPO_MOVIMENTACAO As String = "F15"
Public Const RANGE_HEADER_VALOR_MOVIMENTACAO As String = "G15"
Public Const RANGE_COLUNA_DATA_MOVIMENTACAO As String = "D16:D66"
Public Const RANGE_COLUNA_VALOR_MOVIMENTACAO As String = "G16:G66"
Public Const RANGE_TAB_MOVIMENTACAO As String = "D16:G66"

' Planilhas Jan. a Dez - cartões
Public Const RANGE_COLUNA_VALOR_CARTOES As String = "N16:N66"
Public Const RANGE_ULTIMO_VALOR_CARTAO As String = "N66"
Public Const RANGE_HEADER_CARTOES As String = "J14"
Public Const RANGE_PRIMEIRA_DATA_CARTOES As String = "J16"
Public Const RANGE_COLUNA_DATA_CARTOES As String = "J16:J66"
Public Const RANGE_TAB_CARTOES As String = "J16:N66"

' Portfólio
Public Const RANGE_COLUNA_ATIVO_PORTFOLIO As String = "D74:D112"
Public Const RANGE_COLUNA_SALDO_INICIAL_PORTFOLIO As String = "F74:F112"
Public Const RANGE_COLUNA_AUXILIAR_REND_PORTFOLIO As String = "I74:I112"
Public Const RANGE_COLUNA_SALDO_FINAL_PORTFOLIO As String = "N74:N112"

Public Const RANGE_AREA_RELATORIO As String = "C72:N115"

' Carteira Ações
Public Const RANGE_COLUNA_DATA_ACOES As String = "D117:D146"
Public Const RANGE_COLUNA_ATIVO_ACOES As String = "E117:E146"
Public Const RANGE_COLUNA_QTDE_ACOES As String = "G117:G146"
Public Const RANGE_COLUNA_SALDO_INICIAL_ACOES As String = "F117:F146"
Public Const RANGE_COLUNA_SALDO_FINAL_ACOES As String = "N117:N146"
Public Const RANGE_CELULA_TRIBUTA_ACOES As String = "Q147"
Public Const RANGE_COLUNA_RESULTADO_COMUM_ACOES As String = "X117:X146"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_ACOES As String = "AB117:AB146"
Public Const RANGE_TAB_ACOES As String = "D117:N146"

' Carteira Fundos Imobiliários
Public Const RANGE_COLUNA_DATA_FII As String = "D153:D182"
Public Const RANGE_COLUNA_ATIVO_FII As String = "E153:E182"
Public Const RANGE_COLUNA_QTDE_FII As String = "G153:G182"
Public Const RANGE_COLUNA_SALDO_INICIAL_FII As String = "F153:F182"
Public Const RANGE_COLUNA_SALDO_FINAL_FII As String = "N153:N182"
Public Const RANGE_COLUNA_RESULTADO_COMUM_FII As String = "X153:X182"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_FII As String = "AB153:AB182"
Public Const RANGE_TAB_FII As String = "D153:N182"

' Carteira Tesouro Direto RF
Public Const RANGE_COLUNA_DATA_TESOURO_DIRETO As String = "D189:D203"
Public Const RANGE_COLUNA_ATIVO_TESOURO_DIRETO As String = "E189:E203"
Public Const RANGE_COLUNA_QTDE_TESOURO_DIRETO As String = "G189:G203"
Public Const RANGE_COLUNA_SALDO_INICIAL_TESOURO_DIRETO As String = "F189:F203"
Public Const RANGE_COLUNA_SALDO_FINAL_TESOURO_DIRETO As String = "N189:N203"

' Carteira Tesouro Direto Selic
Public Const RANGE_COLUNA_DATA_TESOURO_SELIC As String = "D210:D215"
Public Const RANGE_COLUNA_ATIVO_TESOURO_SELIC As String = "E210:E215"
Public Const RANGE_COLUNA_QTDE_TESOURO_SELIC As String = "G210:G215"
Public Const RANGE_COLUNA_SALDO_INICIAL_TESOURO_SELIC As String = "F210:F215"
Public Const RANGE_COLUNA_SALDO_FINAL_TESOURO_SELIC As String = "N210:N215"

' Carteira ETF
Public Const RANGE_COLUNA_DATA_ETF As String = "D222:D228"
Public Const RANGE_COLUNA_ATIVO_ETF As String = "E222:E228"
Public Const RANGE_COLUNA_QTDE_ETF As String = "G222:G228"
Public Const RANGE_COLUNA_SALDO_INICIAL_ETF As String = "F222:F228"
Public Const RANGE_COLUNA_SALDO_FINAL_ETF As String = "N222:N228"
Public Const RANGE_COLUNA_RESULTADO_COMUM_ETF As String = "X222:X228"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_ETF As String = "AB222:AB228"

' Carteira Ações USD
Public Const RANGE_COLUNA_DATA_STOCK As String = "D235:D264"
Public Const RANGE_COLUNA_ATIVO_STOCK As String = "E235:E264"
Public Const RANGE_COLUNA_QTDE_STOCK As String = "G235:G264"
Public Const RANGE_COLUNA_SALDO_INICIAL_STOCK As String = "F235:F264"
Public Const RANGE_COLUNA_SALDO_FINAL_STOCK As String = "N235:N264"
Public Const RANGE_CELULA_TRIBUTA_STOCK As String = "Q265"
Public Const RANGE_COLUNA_RESULTADO_COMUM_STOCK As String = "X235:X264"
Public Const RANGE_TAB_STOCK As String = "D235:N264"

' Carteira REIT
Public Const RANGE_COLUNA_DATA_REIT As String = "D271:D300"
Public Const RANGE_COLUNA_ATIVO_REIT As String = "E271:E300"
Public Const RANGE_COLUNA_QTDE_REIT As String = "G271:G300"
Public Const RANGE_COLUNA_SALDO_INICIAL_REIT As String = "F271:F300"
Public Const RANGE_COLUNA_SALDO_FINAL_REIT As String = "N271:N300"
Public Const RANGE_CELULA_TRIBUTA_REIT As String = "Q301"
Public Const RANGE_COLUNA_RESULTADO_COMUM_REIT As String = "X271:X300"
Public Const RANGE_TAB_REIT As String = "D271:N300"

' Carteira Treasuries
Public Const RANGE_COLUNA_DATA_CART_TREASURY As String = "D307:D316"
Public Const RANGE_COLUNA_ATIVO_CART_TREASURY As String = "E307:E316"
Public Const RANGE_COLUNA_QTDE_CART_TREASURY As String = "G307:G316"
Public Const RANGE_COLUNA_SALDO_INICIAL_CART_TREASURY As String = "F307:F316"
Public Const RANGE_COLUNA_SALDO_FINAL_CART_TREASURY As String = "N307:N316"
Public Const RANGE_COLUNA_RESULTADO_COMUM_TREASURY As String = "X307:X316"

' Carteira Ouro
Public Const RANGE_COLUNA_DATA_OURO As String = "D323:D332"
Public Const RANGE_COLUNA_ATIVO_OURO As String = "E323:E332"
Public Const RANGE_COLUNA_QTDE_OURO As String = "G323:G332"
Public Const RANGE_COLUNA_SALDO_INICIAL_OURO As String = "F323:F332"
Public Const RANGE_COLUNA_SALDO_FINAL_OURO As String = "N323:N332"
Public Const RANGE_COLUNA_RESULTADO_COMUM_OURO As String = "X323:X332"

' Carteira Alternativo
Public Const RANGE_COLUNA_DATA_ALTERNATIVO As String = "D339:D348"
Public Const RANGE_COLUNA_ATIVO_ALTERNATIVO As String = "E339:E348"
Public Const RANGE_COLUNA_QTDE_ALTERNATIVO As String = "G339:G348"
Public Const RANGE_COLUNA_SALDO_INICIAL_ALTERNATIVO As String = "F339:F348"
Public Const RANGE_COLUNA_SALDO_FINAL_ALTERNATIVO As String = "N339:N348"
Public Const RANGE_COLUNA_RESULTADO_COMUM_ALTERNATIVO As String = "X339:X348"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_ALTERNATIVO As String = "AB339:AB348"
Public Const RANGE_CELULA_IGNORA_AGENDA_ALTERNATIVO As String = "R351"

' Planilhas Jan. a Dez - indicadores
Public Const RANGE_COLUNA_DESCR_INDICADORES As String = "D358:D367"
Public Const RANGE_COLUNA_MES_INDICADORES As String = "F358:F367"
Public Const RANGE_COLUNA_ANO_INDICADORES As String = "H358:H367"
Public Const RANGE_COLUNA_DOZE_MESES_INDICADORES As String = "I358:I367"
Public Const RANGE_CELULA_DOLAR_FINAL_MES As String = "G364"
Public Const SP500 As String = "S&P 500"
Public Const RANGE_CELULA_DOLAR_BACEN_COMPRA As String = "F371"
Public Const RANGE_CELULA_DOLAR_BACEN_VENDA As String = "G371"
