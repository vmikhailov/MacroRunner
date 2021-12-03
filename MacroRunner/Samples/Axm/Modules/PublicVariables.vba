Option Explicit
' (General Declarations)
'ëÚÓÍË Ë ÍÓÎÓÌÍË ‚ ÒÔ‡‚Ó˜ÌËÍÂ ÚÓ‚‡Ó‚ ËÁ Ä‚‡‰˚
Public FirstRow As Long
Public LastRow As Long
Public LastColumn As Long
Public C_CodeS As Long
Public C_ArticleS As Long
Public C_NameS As Long
Public C_UnitS As Long
Public C_PriceS As Long
Public C_PriceBaseS As Long
Public C_RRCS As Long
Public C_NDSS As Long
Public C_FactorUnitS As Long
Public C_PriceS_RUB_FU As Long
Public C_PriceS_EUR As Long
Public C_PriceS_RUB As Long
Public C_PriceBaseS_FU As Long
Public C_ProfitS As Long
Public C_SkladS As Long
Public C_StatS As Long

'ÓÚ˜ÂÚ ÒÍÎ‡‰ Á‡ Î˛·ÓÈ ‰ÂÌ¸
Public C_GroupS As Long
Public C_Group2S  As Long
Public C_Group3S  As Long
Public C_BrandS As Long

Public FormatFile As Integer 'ÙÎ‡„ ÔÓ‚ÂÍË ÙÓÏ‡Ú‡ Ô‡ÈÒ‡ ÔÓÒÚ‡‚˘ËÍ‡

'‚ Ù‡ÈÎÂ AvardaDir
Public LastRow_W As Long 'èÂ‚‡ﬂ ÒÚÓÍ‡ Ì‡ ‡·Ó˜ÂÈ ÒÚ‡ÌËˆÂ ‚ Ù‡ÈÎÂ ËÁÏÂÌÂÌËÈ
Public FirstRow_W As Long 'èÓÒÎÂ‰Ìﬂﬂ ÒÚÓÍ‡ Ì‡ ‡·Ó˜ÂÈ ÒÚ‡ÌËˆÂ ‚ Ù‡ÈÎÂ ËÁÏÂÌÂÌËÈ
Public ColumnCodeKanz As Long 'äÓÎÓÌÍ‡: ‡ÚËÍÛÎ ÔÓÒÚ‡‚˘ËÍ‡ ‚ ÔÂ‚ÓÈ ÍÓÎÓÌÍÂ
Public ColumnArtKanz As Long 'äÓÎÓÌÍ‡: ‡ÚËÍÛÎ ä‡ÌˆÚÂÌ‰‡
Public ColumnFlagPr As Long 'äÓÎÓÌÍ‡: îÎ‡„ ‚ÍÎ˛˜ÂÌËﬂ ‚ Ô‡ÈÒ ‚ Ä‚‡‰Â
Public ColumnName As Long 'äÓÎÓÌÍ‡: Ì‡ËÏÂÌÓ‚‡ÌËÂ ÚÓ‚‡‡
Public ColumnPrfName As Long 'äÓÎÓÌÍ‡: ÔÂÙËÍÒ‡ ‚ Ì‡ËÏÂÌÓ‚‡ÌËË  ÚÓ‚‡‡
Public ColumnUnit As Long 'äÓÎÓÌÍ‡: Â‰ËÌËˆ˚ ËÁÏÂÂÌËﬂ
Public ColumnRRC As Long 'äÓÎÓÌÍ‡: Å‡ÁÓ‚‡ﬂ ˆÂÌ‡ ËÎË êêñ
Public ColumnNDS As Long '
Public ColumnRC As Long 'äÓÎÓÌÍ‡: ˆÂÌ‡ êñ (ÍÓÔÓ‡ÚË‚Ì‡ﬂ ÓÁÌËˆ‡)
Public ColumnPrice As Long 'äÓÎÓÌÍ‡: Ï‡ÍÒËÏ‡Î¸Ì‡ﬂ ÒÂ·ÂÒÚÓËÏÓÒÚ¸ ËÁ ÚÂı ÓÚ˜ÂÚÓ‚ ËÁ Ä‚‡‰˚

Public ColumnArtKanzNew As Long 'äÓÎÓÌÍ‡: ‡ÚËÍÛÎ ä‡ÌˆÚÂÌ‰‡ New
Public ColumnNameNew As Long 'äÓÎÓÌÍ‡: ÌÓ‚ÓÂ Ì‡ËÏÂÌÓ‚‡ÌËÂ ÚÓ‚‡‡
Public ColumnUnitNew As Long 'äÓÎÓÌÍ‡: ÌÓ‚˚Â Â‰ËÌËˆ˚ ËÁÏÂÂÌËﬂ
Public ColumnRRCNew As Long 'äÓÎÓÌÍ‡: ÌÓ‚‡ﬂ ·‡ÁÓ‚‡ﬂ ˆÂÌ‡
Public ColumnRRCCh As Long 'äÓÎÓÌÍ‡: ÌÓ‚‡ﬂ ·‡ÁÓ‚‡ﬂ ˆÂÌ‡

Public ColumnGroup2 As Long 'äÓÎÓÌÍ‡: ä‡ÚÂ„ÓËﬂ (‡Á‰ÂÎ)
Public ColumnGroup3 As Long 'äÓÎÓÌÍ‡: ä‡ÚÂ„ÓËﬂ (‡Á‰ÂÎ)
Public ColumnGroup As Long 'äÓÎÓÌÍ‡: ÉÛÔÔ‡
Public ColumnBrand As Long 'äÓÎÓÌÍ‡: ÅÂÌ‰/èÓËÁ‚Ó‰ËÚÂÎ¸
Public ColumnDATA_CH As Long 'äÓÎÓÌÍ‡: ‰‡Ú‡ ÔÓÒÎÂ‰ÌÂ„Ó ËÁÏÂÌÂÌËﬂ
Public ColumnDATA_New As Long 'äÓÎÓÌÍ‡: ‰‡Ú‡ ÔÓﬂ‚ÎÂÌËﬂ ‚ Ô‡ÈÒÂ
Public ColumnDATA_Del As Long 'äÓÎÓÌÍ‡: ‰‡Ú‡ Û‰‡ÎÂÌËﬂ
'ÍÓÎÓÌÍË ‰Îﬂ Á‡„ÛÁÍË ËÁ ÓÚ˜ÂÚ‡ ëÎ‡‰ Á‡ Î˛·ÓÈ ‰ÂÌ¸
Public ColumnPriceNew As Long 'äÓÎÓÌÍ‡: ÌÓ‚‡ﬂ ˆÂÌ‡ ‰Îﬂ Ì‡Ò
Public ColumnPriceCh As Long 'äÓÎÓÌÍ‡: ËÁÏÂÌÂÌËÂ Ì‡¯ÂÈ ˆÂÌ˚
Public ColumnSklad As Long 'äÓÎÓÌÍ‡: Ì‡ÎË˜ËÂ ÚÓ‚‡‡ Ì‡ ÒÍÎ‡‰Â
Public ColumnDATA_AnyDay As Long 'äÓÎÓÌÍ‡: ‰‡Ú‡ Á‡„ÛÁÍË ÒÂ·ÂÒÚÓËÏÓÒÚË ËÁ ëÍÎ. Á‡ Î˛·ÓÈ ‰ÂÌ¸
Public ColumnDATA_Last As Long 'äÓÎÓÌÍ‡: ‰‡Ú‡ ÔÓÒÎÂ‰ÌÂÈ ˆÂÌ˚ ëë

Public ColumnFlagCh As Long 'äÓÎÓÌÍ‡: îÎ‡„ Á‡„ÛÁÍË ‚ ÒÔ‡‚Ó˜ÌËÍ AXM
Public ColumnProfit As Long 'äÓÎÓÌÍ‡: ‚ÂÓﬂÚÌ‡ﬂ Ì‡ˆÂÌÍ‡ ÔË ËÒÔÓÎ¸ÁÓ‚‡ÌËË ·‡ÁÓ‚ÓÈ ˆÂÌ˚ ‚ Í‡˜ÂÒÚ‚Â êêñ
Public ColumnStat As Long 'äÓÎÓÌÍ‡: ëÚ‡ÚÛÒ ÚÓ‚‡‡


Public i As Long
Public j As Long
Public FlagAv As Long
Public f1 As String
Public sh1 As String
Public f2 As String
Public sh2 As String
Public PathFile As String
Public NameArticle As String
Public NameCells As String

Public Summ As Double
Public SummPr As Double
Public SummNew As Double
Public datat1 As Date
Public datat2 As Date

'fixes
Public ColCode As Integer
