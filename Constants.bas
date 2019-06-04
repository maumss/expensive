' Constantes globais

Public Const BROKER As String = "Broker"
Public Const RESERVA_ESTRATEGICA As String = "Reserva estratégica"

' Planilha Intro
Public Const RANGE_DATA_ULTIMA_ATUALIZ As String = "E12"
Public Const RANGE_POSICAO = "posicao"

' Planilha Orcamento
Public Const RANGE_TOLERANCIA As String = "tolerancia"
Public Const TIPO_LANCAMENTO_INVESTIMENTOS As String = "investimentos"

' Planilha Alocacao
Public Const RANGE_CELULA_INICIO_ADHOC As String = "C81"
Public Const RANGE_CELULA_FIM_ADHOC As String = "C90"
Public Const RANGE_CELULA_INICIO_CONSOLIDADA As String = "C38"
Public Const RANGE_CELULA_FIM_CONSOLIDADA As String = "C72"

' Planilha Retrato
Public Const RANGE_RELAT_RETRAT = "$A$1:$Q$84"

' Planilhas Jan. a Dez - geral
Public Const RANGE_SITUAC_PLANILHA As String = "E4"
Public Const RANGE_DATA_POSICAO As String = "N4"
Public Const SITUAC_ABERTO As String = "Aberto"
Public Const SITUAC_FECHADO As String = "Fechado"
Public Const RANGE_SALDO_MES As String = "B2"
Public Const NOME_PLAN_DEZ As String = "Dez."

' Planilhas Jan. a Dez - movimentações
Public Const RANGE_HEADER_MOVIMENTACOES As String = "D14"
Public Const RANGE_HEADER_DATA_MOVIMENTACOES As String = "D15"
Public Const RANGE_PRIMEIRA_DATA_MOVIMENTACOES As String = "D16"
Public Const RANGE_HEADER_DESC_MOVIMENTACAO As String = "E15"
Public Const RANGE_HEADER_TIPO_MOVIMENTACAO As String = "F15"
Public Const RANGE_HEADER_VALOR_MOVIMENTACAO As String = "G15"
Public Const RANGE_COL_DATA_MOVIMENTACOES As String = "D16:D66"
Public Const RANGE_COL_VALOR_MOVIMENTACOES As String = "G16:G66"
Public Const RANGE_TAB_MOVIMENTACOES As String = "D16:G66"

' Planilhas Jan. a Dez - cartões
Public Const RANGE_COL_VALOR_CARTOES As String = "N16:N66"
Public Const RANGE_ULTIMO_VALOR_CARTAO As String = "N66"
Public Const RANGE_HEADER_CARTOES As String = "J14"
Public Const RANGE_PRIMEIRA_DATA_CARTOES As String = "J16"
Public Const RANGE_COL_DATA_CARTOES As String = "J16:J66"
Public Const RANGE_TAB_CARTOES As String = "J16:N66"

' Planilhas Jan. a Dez - movimentação dos ativos - carteiras
' Carteira Ad Hoc
Public Const RANGE_COLUNA_ATIVO_ADHOC As String = "D74:D83"
Public Const RANGE_COLUNA_SALDO_INICIAL_ADHOC As String = "F74:F83"
Public Const RANGE_COLUNA_SALDO_FINAL_ADHOC As String = "N74:N83"

' Carteira Consolidada
Public Const RANGE_COLUNA_ATIVO_CONSOLIDADA As String = "D88:D112"
Public Const RANGE_COLUNA_SALDO_INICIAL_CONSOLIDADA As String = "F88:F112"
Public Const RANGE_COLUNA_SALDO_FINAL_CONSOLIDADA As String = "N88:N112"

Public Const RANGE_AREA_RELATORIO As String = "C72:N115"

' Carteira Ações
Public Const RANGE_COL_DATA_ACOES As String = "D117:D146"
Public Const RANGE_COLUNA_ATIVO_ACOES As String = "E117:E146"
Public Const RANGE_COLUNA_QTDE_ACOES As String = "G117:G146"
Public Const RANGE_COLUNA_SALDO_INICIAL_ACOES As String = "F117:F146"
Public Const RANGE_COLUNA_SALDO_FINAL_ACOES As String = "N117:N146"
Public Const RANGE_COLUNA_CUSTO_MEDIO_ACOES As String = "S117:S146"
Public Const RANGE_CELULA_TRIBUTAL_ACOES As String = "Q147"
Public Const RANGE_COLUNA_RESULTADO_COMUM_ACOES As String = "W117:W146"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_ACOES As String = "AA117:AA146"

' Carteira Fundos Imobiliários
Public Const RANGE_COL_DATA_FII As String = "D153:D182"
Public Const RANGE_COLUNA_ATIVO_FII As String = "E153:E182"
Public Const RANGE_COLUNA_QTDE_FII As String = "G153:G182"
Public Const RANGE_COLUNA_SALDO_INICIAL_FII As String = "F153:F182"
Public Const RANGE_COLUNA_SALDO_FINAL_FII As String = "N153:N182"
Public Const RANGE_COLUNA_CUSTO_MEDIO_FII As String = "S153:S182"
Public Const RANGE_COLUNA_RESULTADO_COMUM_FII As String = "W153:W182"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_FII As String = "AA153:AA182"

' Carteira Tesouro Direto RF
Public Const RANGE_COL_DATA_RF As String = "D189:D203"
Public Const RANGE_COLUNA_ATIVO_RF As String = "E189:E203"
Public Const RANGE_COLUNA_QTDE_RF As String = "G189:G203"
Public Const RANGE_COLUNA_SALDO_INICIAL_RF As String = "F189:F203"
Public Const RANGE_COLUNA_SALDO_FINAL_RF As String = "N189:N203"

' Carteira Tesouro Direto Selic
Public Const RANGE_COL_DATA_SELIC As String = "D210:D215"
Public Const RANGE_COLUNA_ATIVO_SELIC As String = "E210:E215"
Public Const RANGE_COLUNA_QTDE_SELIC As String = "G210:G215"
Public Const RANGE_COLUNA_SALDO_INICIAL_SELIC As String = "F210:F215"
Public Const RANGE_COLUNA_SALDO_FINAL_SELIC As String = "N210:N215"

' Carteira ETF
Public Const RANGE_COL_DATA_ETF As String = "D222:D228"
Public Const RANGE_COLUNA_ATIVO_ETF As String = "E222:E228"
Public Const RANGE_COLUNA_QTDE_ETF As String = "G222:G228"
Public Const RANGE_COLUNA_SALDO_INICIAL_ETF As String = "F222:F228"
Public Const RANGE_COLUNA_SALDO_FINAL_ETF As String = "N222:N228"
Public Const RANGE_COLUNA_CUSTO_MEDIO_ETF As String = "S222:S228"
Public Const RANGE_COLUNA_RESULTADO_COMUM_ETF As String = "W222:W228"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_ETF As String = "AA222:AA228"

' Carteira Exterior
Public Const RANGE_COL_DATA_EXTERIOR As String = "D235:D264"
Public Const RANGE_COLUNA_ATIVO_EXTERIOR As String = "E235:E264"
Public Const RANGE_COLUNA_QTDE_EXTERIOR As String = "G235:G264"
Public Const RANGE_COLUNA_SALDO_INICIAL_EXTERIOR As String = "F235:F264"
Public Const RANGE_COLUNA_SALDO_FINAL_EXTERIOR As String = "N235:N264"
Public Const RANGE_COLUNA_CUSTO_MEDIO_EXTERIOR As String = "S235:S264"
Public Const RANGE_CELULA_TRIBUTAL_EXTERIOR As String = "Q265"
Public Const RANGE_COLUNA_RESULTADO_COMUM_EXTERIOR As String = "W235:W264"

' Carteira Opções
Public Const RANGE_COL_DATA_CART_OPCOES As String = "D271:D280"
Public Const RANGE_COLUNA_ATIVO_CART_OPCOES As String = "E271:E280"
Public Const RANGE_COLUNA_QTDE_CART_OPCOES As String = "G271:G280"
Public Const RANGE_COLUNA_SALDO_INICIAL_CART_OPCOES As String = "F271:F280"
Public Const RANGE_COLUNA_SALDO_FINAL_CART_OPCOES As String = "N271:N280"
Public Const RANGE_COLUNA_CUSTO_MEDIO_CART_OPCOES As String = "S271:S280"
Public Const RANGE_COLUNA_RESULTADO_COMUM_OPCOES As String = "W271:W280"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_OPCOES As String = "AA271:AA280"

' Planilhas Jan. a Dez - indicadores
Public Const RANGE_COLUNA_DESCR_INDICADORES As String = "D290:D299"
Public Const RANGE_COLUNA_MES_INDICADORES As String = "F290:F299"
Public Const RANGE_COLUNA_ANO_INDICADORES As String = "H290:H299"
Public Const RANGE_COLUNA_DOZE_MESES_INDICADORES As String = "I290:I299"
Public Const RANGE_CELULA_DOLAR_FINAL_MES As String = "G296"
Public Const SP500 As String = "S&P 500"

' Planilhas Jan. a Dez - Conta corretora
Public Const RANGE_COL_DESC_CONTA_CORRETORA As String = "D307:D309"
Public Const RANGE_COL_SALDO_CONTA_CORRETORA As String = "F307:F309"
Public Const RANGE_COL_BLOQUEADO_CONTA_CORRETORA As String = "G307:G309"

' Planilha Mercado
Public Const RANGE_CELULA_INICIO_QUOTACAO_SIMBOLO As String = "B40"
Public Const RANGE_CELULA_INICIO_QUOTACAO_SIMBOLOATIVO As String = "D40"
Public Const YAHOO_FINANCE_URL As String = "http://download.finance.yahoo.com/d/quotes.csv?s="
Public Const YAHOO_TAG_DADOS As String = "snd1t1c1ol1ghv"
Public Const YAHOO_TAG_FORMATO As String = ".csv"

