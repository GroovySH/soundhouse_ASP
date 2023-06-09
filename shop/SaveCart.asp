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
'	カート内容の保存
'
'更新履歴
'2009/04/10 fCalcShippingのパラメータ追加（個口数）
'2009/04/30 エラー時にerror.aspへ移動
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/07/17 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim userID

Dim wSalesTaxRate
Dim wPrice
Dim wNoData
Dim wOrderProductHTML

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
	wErrDesc = "SaveCart.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Dim vProductNm
Dim vTotalAm
Dim vFreightAm
Dim vFreightForwarder
Dim vSoukoCnt
Dim vKoguchi

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

vTotalAm = 0
wHTML = ""

'----仮受注データ取り出し
wSQL = ""
wSQL = wSQL & "SELECT a.受注明細番号"
wSQL = wSQL & "     , a.メーカーコード"
wSQL = wSQL & "     , a.商品コード"
wSQL = wSQL & "     , a.色"
wSQL = wSQL & "     , a.規格"
wSQL = wSQL & "     , a.メーカー名"
wSQL = wSQL & "     , a.商品名"
wSQL = wSQL & "     , a.受注数量"
wSQL = wSQL & "     , a.受注単価" 
wSQL = wSQL & "     , a.受注金額" 
wSQL = wSQL & "     , b.ASK商品フラグ" 
wSQL = wSQL & "  FROM 仮受注明細 a WITH (NOLOCK)"
wSQL = wSQL & "     ,Web商品 b WITH (NOLOCK)"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
wSQL = wSQL & "   AND b.メーカーコード = a.メーカーコード"
wSQL = wSQL & "   AND b.商品コード = a.商品コード"
wSQL = wSQL & " ORDER BY 受注明細番号"

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

wNoData = false

'---- 明細HTML作成
if RS.EOF = true then
	wNoData = true
	exit function
end if

'----- 見出し
wHTML = wHTML & "<table id='cart'>" & vbNewLine
wHTML = wHTML & "  <tr>" & vbNewLine
wHTML = wHTML & "    <th class='maker'>メーカー</th>" & vbNewLine
wHTML = wHTML & "    <th class='name'>商品名</th>" & vbNewLine
wHTML = wHTML & "    <th class='price'>単価(税込)</th>" & vbNewLine
wHTML = wHTML & "    <th class='number'>数量</th>" & vbNewLine
wHTML = wHTML & "    <th class='amount'>金額(税込)</th>" & vbNewLine
wHTML = wHTML & "  </tr>" & vbNewLine

Do Until RS.EOF = true
	'------------- メーカー、商品名
	vProductNm = RS("商品名")
	if Trim(RS("色")) <> "" then
		vProductNm = vProductNm & "/" & RS("色")
	end if
	if Trim(RS("規格")) <> "" then
		vProductNm = vProductNm & "/" & RS("規格")
	end if
	wHTML = wHTML & "  <tr>" & vbNewLine
	wHTML = wHTML & "    <td>" & RS("メーカー名") & "</td>" & vbNewLine
	wHTML = wHTML & "    <td><a href='ProductDetail.asp?Item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "^" & RS("色") & "^" & RS("規格") & "'>" & vProductNm & "</a></td>" & vbNewLine

		'------------- 単価
	wPrice = calcPrice(RS("受注単価"), wSalesTaxRate)
	vTotalAm = vTotalAm + (wPrice * RS("受注数量"))
	wHTML = wHTML & "    <td class='num'>" & FormatNumber(wPrice,0) & "円</td>" & vbNewLine

		'------------- 数量
	wHTML = wHTML & "    <td class='num'>" & RS("受注数量") & "</td>" & vbNewLine

		'------------- 金額
	wHTML = wHTML & "    <td class='num'>" & FormatNumber(wPrice*RS("受注数量"),0) & "円</td>" & vbNewLine

	RS.MoveNext
Loop

wHTML = wHTML & "  <tr>" & vbNewLine
wHTML = wHTML & "    <td colspan='5'>" & vbNewLine
wHTML = wHTML & "      <dl class='total'>" & vbNewLine
'----商品合計金額
wHTML = wHTML & "        <dt>商品合計(税込)</dt><dd>" & FormatNumber(vTotalAm,0) & "円</dd>" & vbNewLine
'---- 送料
Call fCalcShipping(gSessionID, "一括", vFreightAm, vFreightForwarder, vSoukoCnt, vKoguchi)		'2011/04/14 hn mod
wPrice = Fix(vFreightAm * (100 + wSalesTaxRate) / 100)
wHTML = wHTML & "        <dt>送料見積(税込)</dt><dd>" & FormatNumber(wPrice,0) & "円</dd>" & vbNewLine
wHTML = wHTML & "      </dl>" & vbNewLine
wHTML = wHTML & "    </td>" & vbNewLine
wHTML = wHTML & "  </tr>" & vbNewLine

wHTML = wHTML & "</table>" & vbNewLine

RS.close
wOrderProductHTML = wHTML

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
<title>カート保存｜サウンドハウス</title>
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
      <li class="now">カート保存</li>
    </ul>
  </div></div></div>

  <h1 class="title">カート保存</h1>

<% if wMSG <> "" then %>
  <p class="error"><%=wMSG%></p>
<% end if %>

  <h2 class="cart_title">カート内容</h2>
<%=wOrderProductHTML%>

  <form name="fData" method="post" action="SaveCart2.asp">

    <p>カート名を入力し、[保存する]ボタンを押してください。<br>同じ名前のカートがある場合は上書きします。</p>
    <table class="form">
      <tr>
        <th>カート名</th>
        <td><input name="CartName" type="text" size="20" maxlength="10"><span>(10文字以内）</span></td>
      </tr>
    </table>

    <p>&laquo; <a href="Order.asp">戻る</a></p>
    <p class="btnBox"><input type="submit" value="保存する" class="opover"></p>

  </form>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>