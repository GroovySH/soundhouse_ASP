<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
'サウンドハウス
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->

<%
'========================================================================
'
'	商品比較ページ
'
'	更新履歴
'2005/07/01 納期表示をおよその日数に変更
'2007/05/30 色規格あり商品は1つにまとめ、カートボタンの代わりに、詳細表示ボタンを表示。商品個別画面を表示
'2008/12/24 在庫状況セット関数化
'2010/02/18 an ASK商品パラメータにServer.URLEncodeを行なう
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/10/19 hn 1063 ASK表示方法変更
'2012/01/20 GV データ取得 SELECT文へ LACクエリー案を適用
'2012/07/18 nt リニューアル用にasp画面出力を修正
'2014/03/19 GV 消費税増税に伴う2重表示対応
'
'========================================================================

On Error Resume Next

Dim wTitleWithLink
Dim wNaveWithLink

Dim wHikaku
Dim CategoryCd(5)
Dim MakerCd(5)
Dim ProductCd(5)
Dim Iro(5)
Dim Kikaku(5)
Dim MakerName(5)
Dim ProductName(5)
Dim Price(5)
Dim ImageFile(5)
Dim Chokusou(5)
Dim ASKfl(5)
Dim HikiateKanouSuu(5)
Dim HikiateKanouNyuukaYoteibi(5)
Dim KisyouSuu(5)
Dim Setfl(5)
Dim IroKikakuCnt(5)
Dim	WebNoukiHihyoujiFl(5)
Dim	NyukayoteiMiteiFl(5)
Dim	Haibanbi(5)
Dim	BhinFl(5)
Dim	BhinHikiateKanouQt(5)
Dim	KosuuGenteiQt(5)
Dim	KosuuGenteiJyuchuuQt(5)

Dim wRecCnt

Dim SpecNo(100)			'表示順に商品スペック項目番号
Dim SpecName(100)		'表示順に商品スペック名
Dim SpecComment(5,100)	'商品,表示順が添え字 スペック内容

Dim wSalesTaxRate
Dim wPrice

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS
Dim RS_Template

Dim wHTML

Dim wSQL
Dim wMsg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ProductCompare.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Main
'
'	更新履歴
'2008/05/07 区切り文字変更
'
'========================================================================
'
Function main()

Dim i
Dim vTemp

'---- 送信データーの取り出し
wHikaku = Split(ReplaceInput(Request("item")), "$")
wRecCnt = Ubound(wHikaku)

For i=1 to wRecCnt
	vTemp = Split(wHikaku(i), "^")
	CategoryCd(i) = Trim(vTemp(0))
	MakerCd(i) = Trim(vTemp(1))
	ProductCd(i) = Trim(vTemp(2))
	Iro(i) = Trim(vTemp(3))
	Kikaku(i) = Trim(vTemp(4))
Next

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- ナビゲーションセット
call SetNavi()

'---- タイトルセット
call SetTitle()

'---- 商品スペックテンプレート取り出し
call GetTemplate()

'---- 比較商品データ取り出し
call getCompareProduct()

'---- 比較商品一覧作成
call CreateCompareList()

End Function

'========================================================================
'
'	Function	ナビゲーションセット
'
'========================================================================
'
Function SetNavi()

Dim RSv

'---- ナビゲーションセット
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT a.大カテゴリーコード"
'wSQL = wSQL & "     , a.大カテゴリー名"
'wSQL = wSQL & "     , b.中カテゴリーコード"
'wSQL = wSQL & "     , b.中カテゴリー名日本語"
'wSQL = wSQL & "     , c.カテゴリーコード"
'wSQL = wSQL & "     , c.カテゴリー名"
'wSQL = wSQL & "     , c.お勧めカテゴリーフラグ"
'wSQL = wSQL & "  FROM 大カテゴリー a"
'wSQL = wSQL & "     , 中カテゴリー b"
'wSQL = wSQL & "     , カテゴリー c"
'wSQL = wSQL & " WHERE b.大カテゴリーコード = a.大カテゴリーコード"
'wSQL = wSQL & "   AND c.中カテゴリーコード = b.中カテゴリーコード"
'wSQL = wSQL & "   AND c.カテゴリーコード = '" & CategoryCd(i) & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "       a.大カテゴリーコード "
wSQL = wSQL & "     , a.大カテゴリー名 "
wSQL = wSQL & "     , b.中カテゴリーコード "
wSQL = wSQL & "     , b.中カテゴリー名日本語 "
wSQL = wSQL & "     , c.カテゴリーコード "
wSQL = wSQL & "     , c.カテゴリー名 "
wSQL = wSQL & "     , c.お勧めカテゴリーフラグ "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    大カテゴリー              a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN 中カテゴリー b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.大カテゴリーコード = a.大カテゴリーコード "
wSQL = wSQL & "      INNER JOIN カテゴリー   c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.中カテゴリーコード = b.中カテゴリーコード "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "         c.カテゴリーコード = '" & CategoryCd(1) & "' "
' 2012/01/20 GV Mod End

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wNaveWithLink = ""
wNaveWithLink = wNaveWithLink & "<div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>" & vbNewLine
wNaveWithLink = wNaveWithLink & " <p class='home'><a href='../'><img src='../images/icon_home.gif' alt='HOME'></a></p>" & vbNewLine
wNaveWithLink = wNaveWithLink & " <ul id='path'>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li><a href='LargeCategoryList.asp?LargeCategoryCd=" & RSv("大カテゴリーコード") & "'>" & RSv("大カテゴリー名") & "</a></li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("中カテゴリーコード") & "'>" & RSv("中カテゴリー名日本語") & "</a></li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li><a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("カテゴリーコード") & "'>" &  RSv("カテゴリー名") & "</a></li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  <li class='now'>商品比較</li>" & vbNewLine
wNaveWithLink = wNaveWithLink & "  </ul>" & vbNewLine
wNaveWithLink = wNaveWithLink & "</div></div></div>" & vbNewLine
RSv.close

End Function

'========================================================================
'
'	Function	タイトルセット
'
'========================================================================
'
Function SetTitle()

Dim RSv

'---- タイトルセット
wSQL = ""
' 2012/01/20 GV Mod Start
'wSQL = wSQL & "SELECT a.大カテゴリーコード"
'wSQL = wSQL & "     , a.大カテゴリー名"
'wSQL = wSQL & "     , b.中カテゴリーコード"
'wSQL = wSQL & "     , b.中カテゴリー名日本語"
'wSQL = wSQL & "     , c.カテゴリーコード"
'wSQL = wSQL & "     , c.カテゴリー名"
'wSQL = wSQL & "     , c.お勧めカテゴリーフラグ"
'wSQL = wSQL & "  FROM 大カテゴリー a"
'wSQL = wSQL & "     , 中カテゴリー b"
'wSQL = wSQL & "     , カテゴリー c"
'wSQL = wSQL & " WHERE b.大カテゴリーコード = a.大カテゴリーコード"
'wSQL = wSQL & "   AND c.中カテゴリーコード = b.中カテゴリーコード"
'wSQL = wSQL & "   AND c.カテゴリーコード = '" & CategoryCd(1) & "'"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "       a.大カテゴリーコード "
wSQL = wSQL & "     , a.大カテゴリー名 "
wSQL = wSQL & "     , b.中カテゴリーコード "
wSQL = wSQL & "     , b.中カテゴリー名日本語 "
wSQL = wSQL & "     , c.カテゴリーコード "
wSQL = wSQL & "     , c.カテゴリー名 "
wSQL = wSQL & "     , c.お勧めカテゴリーフラグ "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    大カテゴリー              a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN 中カテゴリー b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.大カテゴリーコード = a.大カテゴリーコード "
wSQL = wSQL & "      INNER JOIN カテゴリー   c WITH (NOLOCK) "
wSQL = wSQL & "        On     c.中カテゴリーコード = b.中カテゴリーコード "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "         c.カテゴリーコード = '" & CategoryCd(1) & "' "
' 2012/01/20 GV Mod Start

'@@@@@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

'2012/7/19 nt add
wTitleWithLink = ""
wTitleWithLink = wTitleWithLink & "<h1 class='title'>" & RSv("カテゴリー名") & " 商品比較</h1>" & vbNewLine

if RSv("お勧めカテゴリーフラグ") = "Y" then
	wTitleWithLink = wTitleWithLink & "<div class='btn_box'>" & vbNewLine
	wTitleWithLink = wTitleWithLink & "<a href='ProductGuide.asp?CategoryCd=" & RSv("カテゴリーコード") & "'><img src='images/btn_recommend.png' alt='おすすめ商品' class='opover'></a>" & vbNewLine
	wTitleWithLink = wTitleWithLink & "</div>" & vbNewLine
end if

'2012/7/19 nt del
'wTitleWithLink = "<b><font color='#696684'><a href='LargeCategoryList.asp?LargeCategoryCd=" & RSv("大カテゴリーコード") & "' class='link'>" & RSv("大カテゴリー名") & "</a>/<a href='MidCategoryList.asp?MidCategoryCd=" & RSv("中カテゴリーコード") & "' class='link'>" & RSv("中カテゴリー名日本語") & "</a>/<a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("カテゴリーコード") & "' class='link'>" & RSv("カテゴリー名") & "</a></font></b>"

'if RSv("お勧めカテゴリーフラグ") = "Y" then
'	wTitleWithLink = wTitleWithLink & "　　<a href='ProductGuide.asp?CategoryCd=" & RSv("カテゴリーコード") & "' class='link'>>>おすすめ商品はこちら</a>" & vbNewLine
'end if

RSv.close

End Function

'========================================================================
'
'	Function	商品スペックテンプレート取り出し
'
'========================================================================
'
Function GetTemplate()

Dim i

'---- 商品スペックテンプレート取り出し
wSQL = ""
wSQL = wSQL & "SELECT 商品スペック項目番号"
wSQL = wSQL & "     , 商品スペック項目名"
wSQL = wSQL & "  FROM 商品スペックテンプレート WITH (NOLOCK)"
wSQL = wSQL & " WHERE カテゴリーコード = '" & CategoryCd(1) & "'"
wSQL = wSQL & " ORDER BY 表示順"

Set RS_Template = Server.CreateObject("ADODB.Recordset")
RS_Template.Open wSQL, Connection, adOpenStatic

i = 1
Do until RS_Template.EOF = true
	SpecNo(i) = RS_Template("商品スペック項目番号")
	SpecName(i) = RS_Template("商品スペック項目名")
	RS_Template.Movenext
	i = i + 1
Loop

RS_Template.close

End Function

'========================================================================
'
'	Function	比較商品データ取り出し
'
'========================================================================
'
Function getCompareProduct()

Dim i
Dim j

For i=1 to wRecCnt
	'---- 商品Recordset作成
	wSQL = ""
' 2012/01/20 GV Mod Start
'	wSQL = wSQL & "SELECT b.メーカーコード"
'	wSQL = wSQL & "     , b.商品コード"
'	wSQL = wSQL & "     , b.色"
'	wSQL = wSQL & "     , b.規格"
'	wSQL = wSQL & "     , a.商品名"
'	wSQL = wSQL & "     , a.商品画像ファイル名_小"
'	wSQL = wSQL & "     , a.メーカー直送取寄区分"
'	wSQL = wSQL & "     , a.ASK商品フラグ"
'	wSQL = wSQL & "     , a.希少数量"
'	wSQL = wSQL & "     , a.セット商品フラグ"
'	wSQL = wSQL & "     , a.個数限定数量"
'	wSQL = wSQL & "     , a.個数限定受注済数量"
'	wSQL = wSQL & "     , a.Web納期非表示フラグ"
'	wSQL = wSQL & "     , a.入荷予定未定フラグ"
'	wSQL = wSQL & "     , a.廃番日"
'	wSQL = wSQL & "     , a.B品フラグ"
'	wSQL = wSQL & "     , CASE"
'	wSQL = wSQL & "         WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価"
'	wSQL = wSQL & "         ELSE a.販売単価"
'	wSQL = wSQL & "       END AS 販売単価"
'	wSQL = wSQL & "     , b.引当可能数量"
'	wSQL = wSQL & "     , b.引当可能入荷予定日"
'	wSQL = wSQL & "     , b.B品引当可能数量"
'	wSQL = wSQL & "     , c.メーカー名"
'	wSQL = wSQL & "     , d.商品スペック項目番号"
'	wSQL = wSQL & "     , d.商品スペック内容"
'
'		'色規格があるかどうか 2007/05/30
'	wSQL = wSQL & "     , (SELECT COUNT(*)"
'	wSQL = wSQL & "          FROM Web色規格別在庫 t"
'	wSQL = wSQL & "         WHERE t.メーカーコード = a.メーカーコード"
'	wSQL = wSQL & "           AND t.商品コード = a.商品コード"
'	wSQL = wSQL & "           AND (t.色 != '' OR t.規格 != '')"
'	wSQL = wSQL & "           AND t.終了日 IS NULL"
'	wSQL = wSQL & "       ) AS 色規格CNT"
'
'	wSQL = wSQL & "  FROM Web商品 a WITH (NOLOCK)"
'	wSQL = wSQL & "     , Web色規格別在庫 b WITH (NOLOCK)"
'	wSQL = wSQL & "     , メーカー c WITH (NOLOCK)"
'	wSQL = wSQL & "     , 商品スペック d WITH (NOLOCK)"
'	wSQL = wSQL & " WHERE b.メーカーコード = a.メーカーコード"
'	wSQL = wSQL & "   AND b.商品コード = a.商品コード"
'	wSQL = wSQL & "   AND c.メーカーコード = a.メーカーコード"
'	wSQL = wSQL & "   AND d.メーカーコード = a.メーカーコード"
'	wSQL = wSQL & "   AND d.商品コード = a.商品コード"
'	wSQL = wSQL & "   AND b.メーカーコード = '" & MakerCd(i) & "'"
'	wSQL = wSQL & "   AND b.商品コード = '" & ProductCd(i) & "'"
'	wSQL = wSQL & "   AND b.色 = '" & Iro(i) & "'"
'	wSQL = wSQL & "   AND b.規格 = '" & Kikaku(i) & "'"
'	wSQL = wSQL & " ORDER BY"
'	wSQL = wSQL & "       c.メーカー名"
'	wSQL = wSQL & "     , a.商品名"
'	wSQL = wSQL & "     , d.商品スペック項目番号"

	wSQL = wSQL & "SELECT "
	wSQL = wSQL & "      b.メーカーコード "
	wSQL = wSQL & "    , b.商品コード "
	wSQL = wSQL & "    , b.色 "
	wSQL = wSQL & "    , b.規格 "
	wSQL = wSQL & "    , a.商品名 "
	wSQL = wSQL & "    , a.商品画像ファイル名_小 "
	wSQL = wSQL & "    , a.メーカー直送取寄区分 "
	wSQL = wSQL & "    , a.ASK商品フラグ "
	wSQL = wSQL & "    , a.希少数量 "
	wSQL = wSQL & "    , a.セット商品フラグ "
	wSQL = wSQL & "    , a.個数限定数量 "
	wSQL = wSQL & "    , a.個数限定受注済数量 "
	wSQL = wSQL & "    , a.Web納期非表示フラグ "
	wSQL = wSQL & "    , a.入荷予定未定フラグ "
	wSQL = wSQL & "    , a.廃番日 "
	wSQL = wSQL & "    , a.B品フラグ "
	wSQL = wSQL & "    , CASE "
	wSQL = wSQL & "        WHEN (a.個数限定数量 > a.個数限定受注済数量 AND a.個数限定数量 > 0) THEN a.個数限定単価 "
	wSQL = wSQL & "        ELSE a.販売単価 "
	wSQL = wSQL & "      END AS 販売単価 "
	wSQL = wSQL & "    , b.引当可能数量 "
	wSQL = wSQL & "    , b.引当可能入荷予定日 "
	wSQL = wSQL & "    , b.B品引当可能数量 "
	wSQL = wSQL & "    , c.メーカー名 "
	wSQL = wSQL & "    , d.商品スペック項目番号 "
	wSQL = wSQL & "    , d.商品スペック内容 "
	wSQL = wSQL & "    , (SELECT COUNT(t.商品コード) "
	wSQL = wSQL & "         FROM Web色規格別在庫 t "
	wSQL = wSQL & "        WHERE     t.メーカーコード = a.メーカーコード "
	wSQL = wSQL & "              AND t.商品コード = a.商品コード "
	wSQL = wSQL & "              AND (t.色 != '' OR t.規格 != '') "
	wSQL = wSQL & "              AND t.終了日 IS NULL "
	wSQL = wSQL & "      ) AS 色規格CNT "
	wSQL = wSQL & "FROM "
	wSQL = wSQL & "    Web商品                      a WITH (NOLOCK) "
	wSQL = wSQL & "      INNER JOIN Web色規格別在庫 b WITH (NOLOCK) "
	wSQL = wSQL & "        ON     b.メーカーコード = a.メーカーコード "
	wSQL = wSQL & "           AND b.商品コード     = a.商品コード "
	wSQL = wSQL & "      INNER JOIN メーカー        c WITH (NOLOCK) "
	wSQL = wSQL & "        ON     c.メーカーコード = a.メーカーコード "
	wSQL = wSQL & "      INNER JOIN 商品スペック    d WITH (NOLOCK) "
	wSQL = wSQL & "        ON     d.メーカーコード = a.メーカーコード "
	wSQL = wSQL & "           AND d.商品コード     = a.商品コード "
	wSQL = wSQL & "WHERE "
	wSQL = wSQL & "        b.メーカーコード = '" & MakerCd(i) & "' "
	wSQL = wSQL & "    AND b.商品コード     = '" & Replace(ProductCd(i), "'", "''") & "' "
	wSQL = wSQL & "    AND b.色             = '" & Iro(i) & "' "
	wSQL = wSQL & "    AND b.規格           = '" & Kikaku(i) & "' "
	wSQL = wSQL & "ORDER BY "
	wSQL = wSQL & "      c.メーカー名 "
	wSQL = wSQL & "    , a.商品名 "
	wSQL = wSQL & "    , d.商品スペック項目番号 "
' 2012/01/20 GV Mod Start

	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open wSQL, Connection, adOpenStatic

	Do Until RS.EOF = true
		MakerName(i) = RS("メーカー名")
		ProductName(i) = RS("商品名")
		Price(i) = RS("販売単価")
		ImageFile(i) = RS("商品画像ファイル名_小")
		Chokusou(i) = RS("メーカー直送取寄区分")
		ASKfl(i) = RS("ASK商品フラグ")
		KisyouSuu(i) = RS("希少数量")
		Setfl(i) = RS("セット商品フラグ")
		HikiateKanouSuu(i) = RS("引当可能数量")
		HikiateKanouNyuukaYoteibi(i) = RS("引当可能入荷予定日")
		IroKikakuCnt(i) = RS("色規格CNT")

		WebNoukiHihyoujiFl(i) = RS("Web納期非表示フラグ")
		NyukayoteiMiteiFl(i) = RS("入荷予定未定フラグ")
		Haibanbi(i) = RS("廃番日")
		BhinFl(i) = RS("B品フラグ")
		BhinHikiateKanouQt(i) = RS("B品引当可能数量")
		KosuuGenteiQt(i) = RS("個数限定数量")
		KosuuGenteiJyuchuuQt(i) = RS("個数限定受注済数量")

		For j=1 to 100
			if SpecNo(j) = RS("商品スペック項目番号") then
				SpecComment(i, j) = RS("商品スペック内容")
				Exit for
			end if
		Next

		RS.MoveNext
	Loop
Next

RS.Close

End function

'========================================================================
'
'	Function	比較商品一覧作成
'
'========================================================================
'
Function createCompareList()

Dim i
Dim j
Dim vLine
Dim vPrice
Dim vProductName
Dim vInventoryCd
Dim vInventoryImage
Dim vWidth
Dim vBgColor

'2012/7/19 nt add
'---- 幅指定
if wRecCnt = 1 then
	vWidth = "200"
elseif wRecCnt = 2 then
	vWidth = "150"
elseif wRecCnt = 3 then
	vWidth = "120"
elseif wRecCnt = 4 then
	vWidth = "100"
elseif wRecCnt = 5 then
	vWidth = "80"
else
	vWidth = "200"
end if

'2012/7/19 nt add
wHTML = ""
wHTML = wHTML & "<table class='productcompare'>" & vbNewLine
wHTML = wHTML & " <tbody>" & vbNewLine

'2012/7/19 nt del
'---- 区切り線
'vLine = ""
'vLine = vLine & "  <tr>" & vbNewLine
'For i=0 to wRecCnt
'	vLine = vLine & "    <td width='100' height='1' bgcolor='#6699cc'><img src='images/blank.gif' width=1 height=1></td>" & vbNewLine
'Next
'vLine = vLine & "  </tr>"

'vWidth = (795 - 110) / wRecCnt

'----
'wHTML = ""
'wHTML = wHTML & "<table border='0' cellspacing='1' cellpadding='0'>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- 製品写真
wHTML = wHTML & "<tr id='prod_img'>" & vbNewLine
wHTML = wHTML & " <th width='" & vWidth & "'>製品写真</th>" & vbNewLine
For i=1 to wRecCnt
	wHTML = wHTML & " <td><img src='prod_img/" & ImageFile(i) & "' alt='" & MakerName(i) & " / " & ProductName(i) & "'></td>" & vbNewLine
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- 製品写真
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>製品写真</td>" & vbNewLine

'For i=1 to wRecCnt
'	wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='#ffffff'><img src='prod_img/" & ImageFile(i) & "' width='124' height='62'></td>" & vbNewLine
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- メーカー
wHTML = wHTML & "<tr>" & vbNewLine
wHTML = wHTML & " <th>メーカー</th>" & vbNewLine
For i=1 to wRecCnt
	wHTML = wHTML & " <td>" & MakerName(i) & "</td>" & vbNewLine
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- メーカー
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>メーカー</td>" & vbNewLine

'For i=1 to wRecCnt
'	if i Mod 2 = 0 then
'		vBgColor = "#eeeeee"
'	else
'		vBgColor = "#ffffff"
'	end if
'	wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "'class='honbun'>" & MakerName(i) & "</td>" & vbNewLine & vbNewLine
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- 商品名/色/規格
wHTML = wHTML & "<tr>" & vbNewLine
wHTML = wHTML & " <th>商品名</th>" & vbNewLine
For i=1 to wRecCnt
	vProductName = ProductName(i)
	if Trim(Iro(i)) <> "" then
		vProductName = vProductName & "/" & Trim(Iro(i))
	end if
	if Trim(Kikaku(i)) <> "" then
		vProductName = vProductName & "/" & Trim(Kikaku(i))
	end if

	wHTML = wHTML & " <td><a href='ProductDetail.asp?item=" & MakerCd(i) & "^" & ProductCd(i) & "^" & Iro(i) & "^" & Kikaku(i) & "'>" & vProductName & "</a></td>" & vbNewLine
Next
wHTML = wHTML & "</tr>"

'2012/7/19 nt del
'---- 商品名/色/規格
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>商品名</td>" & vbNewLine

'For i=1 to wRecCnt
'	if i Mod 2 = 0 then
'		vBgColor = "#eeeeee"
'	else
'		vBgColor = "#ffffff"
'	end if

'	vProductName = ProductName(i)
'	if Trim(Iro(i)) <> "" then
'		vProductName = vProductName & "/" & Trim(Iro(i))
'	end if
'	if Trim(Kikaku(i)) <> "" then
'		vProductName = vProductName & "/" & Trim(Kikaku(i))
'	end if

'	wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "'><a href='ProductDetail.asp?item=" & MakerCd(i) & "^" & ProductCd(i) & "^" & Iro(i) & "^" & Kikaku(i) & "' class='link'>" & vProductName & "</a></td>" & vbNewLine
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- 衝撃特価
wHTML = wHTML & "<tr>" & vbNewLine
wHTML = wHTML & " <th>衝撃特価</th>" & vbNewLine
For i=1 to wRecCnt
	vPrice = calcPrice(Price(i), wSalesTaxRate)
	wHTML = wHTML & " <td>"
	if ASKfl(i) = "Y" then
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "<span class='honbun'><a class='tip'>ASK<span>" & FormatNumber(vPrice,0) & "円(税込)</span></a></span>" & vbNewLine
		wHTML = wHTML & "<span class='honbun'><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(Price(i),0) & "円(税抜)</span><br>"
		wHTML = wHTML & "<span class='inc-tax'>(税込&nbsp;" & FormatNumber(vPrice,0) & "円)</span></a>" & vbNewLine
'2014/03/19 GV mod end   <----
	else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & FormatNumber(vPrice,0) & "円(税込)" & vbNewLine
		wHTML = wHTML & FormatNumber(Price(i),0) & "円(税抜)<br>" & vbNewLine
		wHTML = wHTML & "(税込&nbsp;" & FormatNumber(vPrice,0) & "円)" & vbNewLine
'2014/03/19 GV mod end   <----
	end if
	wHTML = wHTML & " </td>" & vbNewLine
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- 衝撃特価
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>衝撃特価</td>" & vbNewLine

'For i=1 to wRecCnt
'	if i Mod 2 = 0 then
'		vBgColor = "#eeeeee"
'	else
'		vBgColor = "#ffffff"
'	end if
'	vPrice = calcPrice(Price(i), wSalesTaxRate)
'	wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "'>"
'	if ASKfl(i) = "Y" then

'2011/10/19 hn mod s
'		wHTML = wHTML & "<a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(MakerName(i)) & "&ProductName=" & Server.URLEncode(ProductName(i)) & "&Price=" & vPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a>"

'		wHTML = wHTML & "<span class='honbun'><a class='tip'>ASK<span>" & FormatNumber(vPrice,0) & "円(税込)</span></a></span>"

'2011/10/19 hn mod e

'	else
'		wHTML = wHTML & "<span class='honbun'>" & FormatNumber(vPrice,0) & "円(税込)</span>"
'	end if
'	wHTML = wHTML & "</td>" & vbNewLine
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- 在庫状況
wHTML = wHTML & "<tr>" & vbNewLine
wHTML = wHTML & " <th>在庫状況</th>" & vbNewLine
For i=1 to wRecCnt
	if IroKikakuCnt(i) = 0 then	'2007/05/30
		vInventoryCd = GetInventoryStatus(Makercd(i),ProductCd(i),Iro(i),Kikaku(i),HikiateKanouSuu(i),KisyouSuu(i),Setfl(i),Chokusou(i),HikiateKanouNyuukaYoteibi(i),"N")

		'---- 在庫状況、色を最終セット
		call GetInventoryStatus2(HikiateKanouSuu(i), WebNoukiHihyoujiFl(i), NyukayoteiMiteiFl(i), Haibanbi(i), BhinFl(i), BhinHikiateKanouQt(i), KosuuGenteiQt(i), KosuuGenteiJyuchuuQt(i), "N", vInventoryCd, vInventoryImage)
		wHTML = wHTML & " <td class='stock'><img src='images/" & vInventoryImage & "'>" & vInventoryCd & "</td>" & vbNewLine

	else
		wHTML = wHTML & " <td class='stock'></td>" & vbNewLine
	end if
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- 在庫状況
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap class='honbun'>在庫状況</td>" & vbNewLine

'For i=1 to wRecCnt
'	if i Mod 2 = 0 then
'		vBgColor = "#eeeeee"
'	else
'		vBgColor = "#ffffff"
'	end if

'	if IroKikakuCnt(i) = 0 then	'2007/05/30
'		vInventoryCd = GetInventoryStatus(Makercd(i),ProductCd(i),Iro(i),Kikaku(i),HikiateKanouSuu(i),KisyouSuu(i),Setfl(i),Chokusou(i),HikiateKanouNyuukaYoteibi(i),"N")

		'---- 在庫状況、色を最終セット
'		call GetInventoryStatus2(HikiateKanouSuu(i), WebNoukiHihyoujiFl(i), NyukayoteiMiteiFl(i), Haibanbi(i), BhinFl(i), BhinHikiateKanouQt(i), KosuuGenteiQt(i), KosuuGenteiJyuchuuQt(i), "N", vInventoryCd, vInventoryImage)

'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "' class='honbun'><img src='images/" & vInventoryImage & "' width=10 height=10> " & vInventoryCd & "</td>" & vbNewLine

'	else
'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='" & vBgColor & "' class='honbun'></td>" & vbNewLine
'	end if
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine

'wHTML = wHTML & vLine

'2012/7/19 nt add
'---- スペック
wHTML = wHTML & "<tr id='spec'>" & vbNewLine
wHTML = wHTML & " <th colspan='6'>スペック</th>" & vbNewLine
wHTML = wHTML & "</tr>" & vbNewLine
For j=1 to 100
	if SpecNo(j) = "" then
		exit for
	end if

	wHTML = wHTML & "<tr>" & vbNewLine
	wHTML = wHTML & " <th>" & SpecName(j) & "</th>" & vbNewLine

	For i=1 to wRecCnt
		wHTML = wHTML & " <td>" & vbNewLine
		if Trim(SpecComment(i, j)) = "" OR IsNull(SpecComment(i, j)) = true then
			wHTML = wHTML & "-" & vbNewLine
		else
			wHTML = wHTML & SpecComment(i, j) & vbNewLine
		end if
		wHTML = wHTML & " </td>" & vbNewLine
	Next

	wHTML = wHTML & "  </tr>" & vbNewLine
Next

'2012/7/19 nt del
'---- スペック
'wHTML = wHTML & "  <tr align='left' valign='bottom'>" & vbNewLine
'wHTML = wHTML & "    <td align='center' height='30' class='honbun'><b>スペック</b></td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & vLine

'For j=1 to 100
'	if SpecNo(j) = "" then
'		exit for
'	end if

'	wHTML = wHTML & "  <tr class='honbun'>"
'	wHTML = wHTML & "    <td width='100' align='center' bgcolor='#eeeeee' nowrap>" & SpecName(j) & "</td>" & vbNewLine

'	For i=1 to wRecCnt
'		if i Mod 2 = 0 then
'			vBgColor = "#eeeeee"
'		else
'			vBgColor = "#ffffff"
'		end if
'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' valign='top' bgcolor='" & vBgColor & "'>"
'		if Trim(SpecComment(i, j)) = "" OR IsNull(SpecComment(i, j)) = true then
'			wHTML = wHTML & "-"
'		else
'			wHTML = wHTML & SpecComment(i, j)
'		end if
'		wHTML = wHTML & "</td>" & vbNewLine
'	Next

'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & vLine
'Next

'2012/7/19 nt add
'---- カート
wHTML = wHTML & "<tr id='cart'>" & vbNewLine
wHTML = wHTML & " <th>カートへ</th>" & vbNewLine

For i=1 to wRecCnt
	if IroKikakuCnt(i) = 0 then
		wHTML = wHTML & " <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine
		wHTML = wHTML & "  <td nowrap>" & vbNewLine
		wHTML = wHTML & "   <input type='text' name='qt' value='1'>" & vbNewLine
		wHTML = wHTML & "   <input type='image' src='images/btn_cart.png' alt='カートに入れる' class='opover'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='maker_cd' value='" & MakerCd(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='product_cd' value='" & ProductCd(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='iro' value='" & Iro(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='kikaku' value='" & Kikaku(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <input type='hidden' name='category_cd' value='" & CategoryCd(i) & "'>" & vbNewLine
		wHTML = wHTML & "  </td>" & vbNewLine
		wHTML = wHTML & " </form>" & vbNewLine
	else
		wHTML = wHTML & " <td>" & vbNewLine
		wHTML = wHTML & "  <a href='ProductDetail.asp?Item=" & MakerCd(i) & "^" & ProductCd(i) & "'>" & vbNewLine
		wHTML = wHTML & "   <img src='images/Shousai.gif'>" & vbNewLine
		wHTML = wHTML & "  </a>" & vbNewLine
		wHTML = wHTML & " </td>" & vbNewLine
	end if
Next
wHTML = wHTML & "</tr>" & vbNewLine

'2012/7/19 nt del
'---- カート
'wHTML = wHTML & "  <tr>"
'wHTML = wHTML & "    <td width='100' align='center' bgcolor='#ffffff' nowrap class='honbun'>カートへ</td>" & vbNewLine

'For i=1 to wRecCnt
'	if IroKikakuCnt(i) = 0 then	'2007/05/30
'		wHTML = wHTML & "    <form name='f_item' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine
'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='#ffffff'class='honbun'>" & vbNewLine
'		wHTML = wHTML & "      <input type='text' name='qt' size='3' value='1'>" & vbNewLine
'		wHTML = wHTML & "      <input type='image' src='images/CartSmall.jpg' width='22' height='18'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='maker_cd' value='" & MakerCd(i) & "'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='product_cd' value='" & ProductCd(i) & "'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='iro' value='" & Iro(i) & "'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='kikaku' value='" & Kikaku(i) & "'>" & vbNewLine
'		wHTML = wHTML & "      <input type='hidden' name='category_cd' value='" & CategoryCd(i) & "'>" & vbNewLine
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "    </form>" & vbNewLine

'	else
'		wHTML = wHTML & "    <td width='" & vWidth & "' align='center' bgcolor='#ffffff'class='honbun'>" & vbNewLine
'		wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & MakerCd(i) & "^" & ProductCd(i) & "'>"
'		wHTML = wHTML & "      <img src='images/Shousai.gif' border='0'></a>" & vbNewLine
'		wHTML = wHTML & "    </td>" & vbNewLine
'	end if
'Next

'wHTML = wHTML & "  </tr>" & vbNewLine

'wHTML = wHTML & vLine

wHTML = wHTML & "</tbody>" '2012/7/19 nt add
wHTML = wHTML & "</table>"

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing    '2011/08/01 an add

End function

'========================================================================
%>

<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>商品比較｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="Style/shop.css" type="text/css">
<link rel="stylesheet" href="Style/productcompare.css" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">

<script type="text/javascript">
//
// ====== 	Function:	order_onClick
//
function order_onClick(pForm){
	if (pForm.qt.value == ""){
		pForm.qt.value = 0;
	}else{
		if (numericCheck(pForm.qt.value) == false){
			pForm.qt.value = 0;
		}
	}
	if (pForm.qt.value > 0){
		return true;
	}else{
		alert("数量を入力してからカートボタンを押してください。");
		return false;
	}
}
</script>

</head>

<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
	<span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
	<!-- コンテンツstart -->
	<div id="globalContents">
		<!-- ナビゲーション -->
		<%=wNaveWithLink%>
		<!-- タイトル -->	
		<%=wTitleWithLink%>
		<!-- 比較リスト -->
		<%=wHTML%>
	</div>
	<div id="globalSide">
		<!--#include file="../Navi/NaviSide.inc"-->
	</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<div class="tooltip"><p>ASK</p></div>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
</body>
</html>
