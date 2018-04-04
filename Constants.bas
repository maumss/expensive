Attribute VB_Name = "Constants"
' Constantes globais

' Planilha Intro
Public Const RANGE_DATA_ULTIMA_ATUALIZ As String = "E12"
Public Const RANGE_POSICAO = "posicao"

' Planilha Orcamento
Public Const RANGE_TOLERANCIA As String = "tolerancia"
Public Const TIPO_LANCAMENTO_INVESTIMENTOS As String = "investimentos"

' Planilha Alocacao
Public Const RANGE_CELULA_INICIO_ADHOC As String = "C67"
Public Const RANGE_CELULA_FIM_ADHOC As String = "C72"
Public Const RANGE_CELULA_INICIO_CONSOLIDADA As String = "C39"
Public Const RANGE_CELULA_FIM_CONSOLIDADA As String = "C61"

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
Public Const RANGE_COLUNA_ATIVO_ADHOC As String = "D75:D79"
Public Const RANGE_COLUNA_SALDO_INICIAL_ADHOC As String = "F75:F79"
Public Const RANGE_COLUNA_SALDO_FINAL_ADHOC As String = "N75:N79"

' Carteira Consolidada
Public Const RANGE_COLUNA_ATIVO_CONSOLIDADA As String = "D84:D104"
Public Const RANGE_COLUNA_SALDO_INICIAL_CONSOLIDADA As String = "F84:F104"
Public Const RANGE_COLUNA_SALDO_FINAL_CONSOLIDADA As String = "N84:N104"

' Carteira Ações
Public Const RANGE_COL_DATA_ACOES As String = "D109:D128"
Public Const RANGE_COLUNA_ATIVO_ACOES As String = "E109:E128"
Public Const RANGE_COLUNA_QTDE_ACOES As String = "G109:G128"
Public Const RANGE_COLUNA_SALDO_INICIAL_ACOES As String = "F109:F128"
Public Const RANGE_COLUNA_SALDO_FINAL_ACOES As String = "N109:N128"
Public Const RANGE_COLUNA_CUSTO_MEDIO_ACOES As String = "S109:S128"
Public Const RANGE_CELULA_TRIBUTAL_ACOES As String = "Q129"
Public Const RANGE_COLUNA_RESULTADO_COMUM_ACOES As String = "W109:W128"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_ACOES As String = "AA109:AA128"

' Carteira Opções
Public Const RANGE_COL_DATA_CART_OPCOES As String = "D135:D144"
Public Const RANGE_COLUNA_ATIVO_CART_OPCOES As String = "E135:E144"
Public Const RANGE_COLUNA_QTDE_CART_OPCOES As String = "G135:G144"
Public Const RANGE_COLUNA_SALDO_INICIAL_CART_OPCOES As String = "F135:F144"
Public Const RANGE_COLUNA_SALDO_FINAL_CART_OPCOES As String = "N135:N144"
Public Const RANGE_COLUNA_CUSTO_MEDIO_CART_OPCOES As String = "S135:S144"
Public Const RANGE_COLUNA_RESULTADO_COMUM_OPCOES As String = "W135:W144"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_OPCOES As String = "AA135:AA144"

' Carteira Fundos Imobiliários
Public Const RANGE_COL_DATA_FII As String = "D151:D165"
Public Const RANGE_COLUNA_ATIVO_FII As String = "E151:E165"
Public Const RANGE_COLUNA_QTDE_FII As String = "G151:G165"
Public Const RANGE_COLUNA_SALDO_INICIAL_FII As String = "F151:F165"
Public Const RANGE_COLUNA_SALDO_FINAL_FII As String = "N151:N165"
Public Const RANGE_COLUNA_CUSTO_MEDIO_FII As String = "S151:S165"
Public Const RANGE_COLUNA_RESULTADO_COMUM_FII As String = "W151:W165"
Public Const RANGE_COLUNA_RESULTADO_DAYTRADE_FII As String = "AA151:AA165"

' Carteira Tesouro Direto RF
Public Const RANGE_COL_DATA_RF As String = "D172:D186"
Public Const RANGE_COLUNA_ATIVO_RF As String = "E172:E186"
Public Const RANGE_COLUNA_QTDE_RF As String = "G172:G186"
Public Const RANGE_COLUNA_SALDO_INICIAL_RF As String = "F172:F186"
Public Const RANGE_COLUNA_SALDO_FINAL_RF As String = "N172:N186"

' Carteira Tesouro Direto Selic
Public Const RANGE_COL_DATA_SELIC As String = "D193:D198"
Public Const RANGE_COLUNA_ATIVO_SELIC As String = "E193:E198"
Public Const RANGE_COLUNA_QTDE_SELIC As String = "G193:G198"
Public Const RANGE_COLUNA_SALDO_INICIAL_SELIC As String = "F193:F198"
Public Const RANGE_COLUNA_SALDO_FINAL_SELIC As String = "N193:N198"

' Planilhas Jan. a Dez - indicadores
Public Const RANGE_AREA_RELATORIO As String = "C73:N107"
Public Const RANGE_COLUNA_DESCR_INDICADORES As String = "D208:D217"
Public Const RANGE_COLUNA_MES_INDICADORES As String = "F208:F217"
Public Const RANGE_COLUNA_ANO_INDICADORES As String = "H208:H217"
Public Const RANGE_COLUNA_DOZE_MESES_INDICADORES As String = "I208:I217"
Public Const RANGE_CELULA_DOLAR_FINAL_MES As String = "G214"
Public Const SP500 As String = "S&P 500"
' Conta corretora
Public Const RANGE_COL_DESC_CONTA_CORRETORA As String = "D221:D222"
Public Const RANGE_COL_SALDO_CONTA_CORRETORA As String = "F221:F222"
Public Const RANGE_COL_BLOQUEADO_CONTA_CORRETORA As String = "G221:G222"

' Planilha Mercado
Public Const RANGE_CELULA_INICIO_QUOTACAO_SIMBOLO As String = "B40"
Public Const RANGE_CELULA_INICIO_QUOTACAO_SIMBOLOATIVO As String = "D40"
Public Const YAHOO_FINANCE_URL As String = "http://download.finance.yahoo.com/d/quotes.csv?s="
Public Const YAHOO_TAG_DADOS As String = "snd1t1c1ol1ghv"
Public Const YAHOO_TAG_FORMATO As String = ".csv"

