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
'	保存カートの一覧
'
'更新履歴
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/07/17 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim userID

Dim wSalesTaxRate
Dim wPrice
Dim wCartHTML

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- UserID 取り出し
userID = Session("userID")

wMSG = ReplaceInput(Request("msg"))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SaveCartList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'========================================================================
'
Function main()

Dim vDateStored
Dim vCartName
Dim vTotalAm
Dim vBreakKey
Dim vBreakNextKey

wHTML = "" & vbNewLine

if userID = "" then
	wHTML = wHTML & "<p class='error'>ログインをしてください。</p>" & vbNewLine
	wCartHTML = wHTML
	exit function
end if

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

wHTML = ""

'----保存カートデータ取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.カート名"
wSQL = wSQL & "     , a.登録日"
wSQL = wSQL & "     , b.メーカーコード"
wSQL = wSQL & "     , b.商品コード"
wSQL = wSQL & "     , b.色"
wSQL = wSQL & "     , b.規格"
wSQL = wSQL & "     , b.受注数量"
wSQL = wSQL & "     , CASE"
wSQL = wSQL & "         WHEN (c.個数限定数量 > c.個数限定受注済数量 AND c.個数限定数量 > 0) THEN c.個数限定単価"
wSQL = wSQL & "         ELSE c.販売単価"
wSQL = wSQL & "       END AS 販売単価"
wSQL = wSQL & "     , c.終了日"
wSQL = wSQL & "     , c.取扱中止日"
wSQL = wSQL & "     , c.廃番日"
wSQL = wSQL & "     , c.完売日"
wSQL = wSQL & "     , c.B品単価"
wSQL = wSQL & "     , c.B品フラグ"
wSQL = wSQL & "  FROM 保存カート a WITH (NOLOCK)"
wSQL = wSQL & "     , 保存カート明細 b WITH (NOLOCK)"
wSQL = wSQL & "     , Web商品 c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.顧客番号 = a.顧客番号"
wSQL = wSQL & "   AND b.カート名 = a.カート名"
wSQL = wSQL & "   AND c.メーカーコード = b.メーカーコード"
wSQL = wSQL & "   AND c.商品コード = b.商品コード"
wSQL = wSQL & "   AND a.顧客番号 = " & userID
wSQL = wSQL & " ORDER BY a.登録日 DESC"

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

'----- 見出し
wHTML = wHTML & "<table id='saveCart'>" & vbNewLine
wHTML = wHTML & "  <tr>" & vbNewLine
wHTML = wHTML & "    <th class='date'>登録日</th>" & vbNewLine
wHTML = wHTML & "    <th class='name'>カート名</th>" & vbNewLine
wHTML = wHTML & "    <th class='total'>商品合計(税込)</th>" & vbNewLine
wHTML = wHTML & "    <th class='cart'>&nbsp;</th>" & vbNewLine
wHTML = wHTML & "    <th class='delete'>&nbsp;</th>" & vbNewLine
wHTML = wHTML & "  </tr>" & vbNewLine

if RS.EOF = true then
	wHTML = wHTML & "  <tr><td colspan='5'><p class='error'>保存されたカートがありません。</p></td></tr>" 
	wHTML = wHTML & "</table>" & vbNewLine
	wCartHTML = wHTML
	exit function
end if

vBreakNextKey = RS("カート名")
vBreakKey = vBreakNextKey
vTotalAm = 0

Do Until RS.EOF = true
	if RS("B品フラグ") = "Y" then
		wPrice = calcPrice(RS("B品単価"), wSalesTaxRate)
	else
		wPrice = calcPrice(RS("販売単価"), wSalesTaxRate)
	end if

	vTotalAm = vTotalAm + (wPrice * RS("受注数量"))
	vDateStored = fFormatDate(RS("登録日"))
	vCartName = RS("カート名")

	RS.MoveNext

	if RS.EOF = false then
		vBreakNextKey = RS("カート名")
	else
		vBreakNextKey = "@EOF"
	end if

	if vBreakKey <> vBreakNextKey then
		'------------- 登録日
		wHTML = wHTML & "  <tr>" & vbNewLine
		wHTML = wHTML & "    <td class='date'>" & vDateStored & "</td>" & vbNewLine

		'------------- カート名
		wHTML = wHTML & "    <td class='name'>" & vCartName & "</td>" & vbNewLine

			'------------- 商品合計
		wHTML = wHTML & "    <td class='total'>" & FormatNumber(vTotalAm,0) & "円</td>" & vbNewLine

			'------------- カートへボタン
		wHTML = wHTML & "    <td class='cart'><a href='SaveCartMoveToOrder.asp?CartName=" & Server.URLencode(vCartName) & "'><img src='images/btn_cart.png' alt='カートへ' class='opover'></a></td>" & vbNewLine

			'------------- 削除ボタン
		wHTML = wHTML & "    <td class='delete'><a href='SaveCartDelete.asp?CartName=" & Server.URLencode(vCartName) & "' class='tipBtn'>削除</a></td>" & vbNewLine

		wHTML = wHTML & "  </tr>" & vbNewLine

		vBreakKey = vBreakNextKey
		vTotalAm = 0
	end if

Loop

wHTML = wHTML & "</table>" & vbNewLine

RS.close
wCartHTML = wHTML

End Function

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
<title>保存カート一覧｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120717" type="text/css">
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">

  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now">保存カート一覧</li>
    </ul>
  </div></div></div>

  <h1 class="title">保存カート一覧</h1>

<% if wMSG <> "" then %>
	<p class="error"><%=wMSG%></p>
<% end if %>

<%=wCartHTML%>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>