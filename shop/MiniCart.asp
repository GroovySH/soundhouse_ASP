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
'	カートポップアップページ
'更新履歴
'2005/02/21 hn ASK表示を無しにする
'2009/04/30 エラー時にerror.aspへ移動
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/20 an SELECT文へLACクエリー案を適用
'
'========================================================================

On Error Resume Next

Dim msg

Dim Connection
Dim RS

Dim wTotalCnt
Dim wTotalAm
Dim wSalesTaxRate
Dim wPrice

Dim wListHTML

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim w_sql
Dim w_html
Dim w_error_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.Expires = -1			' Do not cache

'---- 呼び出し元プログラムからのメッセージ取り出し

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "MiniCart.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	商品明細作成
'
'========================================================================
Function main()
Dim i

'---- 消費税率取出し
call getCntlMst("共通","消費税率","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'消費税率
wSalesTaxRate = Clng(wItemNum1)

'---- 仮受注データSELECT

w_sql = ""
w_sql = w_sql & "SELECT a.メーカーコード"
w_sql = w_sql & "     , a.商品コード"
w_sql = w_sql & "     , a.メーカー名"
w_sql = w_sql & "     , a.商品名"
w_sql = w_sql & "     , a.色"
w_sql = w_sql & "     , a.規格"
w_sql = w_sql & "     , a.受注数量"
w_sql = w_sql & "     , a.受注単価" 
w_sql = w_sql & "     , a.受注金額" 
'w_sql = w_sql & "     , b.ASK商品フラグ"            '2012/01/20 an del
w_sql = w_sql & "  FROM 仮受注明細 a WITH (NOLOCK)"  '2012/01/20 an mod
'w_sql = w_sql & "     , Web商品 b"
w_sql = w_sql & " WHERE a.SessionID = '" & gSessionID & "'"       '2011/04/14 hn mod
'w_sql = w_sql & "   AND b.メーカーコード = a.メーカーコード"     '2012/01/20 an del
'w_sql = w_sql & "   AND b.商品コード = a.商品コード"             '2012/01/20 an del
w_sql = w_sql & " ORDER BY a.受注明細番号"

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

'---- 明細HTML作成
wTotalCnt = 0
wTotalAm = 0

w_html = ""

if RS.EOF = true then
	w_html = w_html & "<br><center><span class='honbun'>カートに商品がありません。</span></center>"
else
	w_html = w_html & "<table bgcolor='#000000' border='0' width='100%' cellpadding='0' cellspacing='1'>" & vbNewLine
	w_html = w_html & "<tr>" & vbNewLine
	w_html = w_html & "<td>" & vbNewLine
	w_html = w_html & "<table bgcolor='#ffffff' border='0' class='small' width='100%' cellpadding='0' cellspacing='2'>" & vbNewLine

	Do While RS.EOF = false
	'----- メーカー、商品名
		w_html = w_html & "  <tr>" & vbNewLine
		w_html = w_html & "    <td align='left' valign='middle' colspan='3'><font size='-1'>" & RS("メーカー名") & " <a href='JavaScript:Product_onClick(""Item=" & RS("メーカーコード") & "^" & Server.URLEncode(RS("商品コード")) & "^" & Trim(RS("色")) & "^" & Trim(RS("規格")) & """)'>" & RS("商品名")
		if Trim(RS("色")) <> "" then
			w_html = w_html & "/" & Trim(RS("色"))
		end if
		if Trim(RS("規格")) <> "" then
			w_html = w_html & "/" & Trim(RS("規格"))
		end if
		w_html = w_html & "</a></font></td>" & vbNewLine
		w_html = w_html & "  </tr>" & vbNewLine

	'----- 数量、単価、金額
		w_html = w_html & "  <tr>" & vbNewLine
		w_html = w_html & "    <td align='left' valign='middle' nowrap>数量: " & RS("受注数量") & "</td>" & vbNewLine

'@@@@2005/02/21 change start
		wPrice = calcPrice(RS("受注単価"), wSalesTaxRate)
		w_html = w_html & "    <td align='left' valign='middle' nowrap>単価: " & FormatNumber(wPrice,0) & "円</td>" & vbNewLine
		w_html = w_html & "    <td align='right' valign='middle' nowrap>金額: " & FormatNumber((wPrice * RS("受注数量")),0) & "円</td>" & vbNewLine

'		if RS("ASK商品フラグ") = "Y" then
'			wPrice = 0
'			w_html = w_html & "    <td align='left' valign='middle' nowrap>単価: ASK</td>" & vbNewLine
'			w_html = w_html & "    <td align='right' valign='middle' nowrap>金額: ASK</td>" & vbNewLine
'		else
'			wPrice = calcPrice(RS("受注単価"), wSalesTaxRate)
'			w_html = w_html & "    <td align='left' valign='middle' nowrap>単価: " & FormatNumber(wPrice,0) & "円</td>" & vbNewLine
'			w_html = w_html & "    <td align='right' valign='middle' nowrap>金額: " & FormatNumber((wPrice * RS("受注数量")),0) & "円</td>" & vbNewLine
'		end if
'@@@@ 2005/02/21 change end

		w_html = w_html & "  </tr>" & vbNewLine

		wTotalCnt = wTotalCnt + RS("受注数量")
		wTotalAm = wTotalAm + wPrice * RS("受注数量")

		RS.MoveNext

		'----- ライン
		if RS.EOF = false then
			w_html = w_html & "  <tr bgcolor='#ffc600'>" & vbNewLine
			w_html = w_html & "    <td height='1' colspan='3'><img src='images/blank.gif' width='1' height='1'></td>"
			w_html = w_html & "  </tr>" & vbNewLine
		end if

	Loop
	w_html = w_html & "</table>" & vbNewLine
	w_html = w_html & "</td>" & vbNewLine
	w_html = w_html & "</tr>" & vbNewLine
	w_html = w_html & "</table>" & vbNewLine

end if

wListHTML = w_html

RS.Close

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

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>カートの内容</title>

<!--#include file="../Navi/NaviStyle.inc"-->

<script language="JavaScript">

//
// ====== 呼び出し元Windowへカート確認ページを表示
//
function GoToCart_onClick(){
//	parent.opener.location = "Order.asp";
	parent.location = "<%=g_HTTP%>shop/Order.asp";
}

//
// ====== 注文ページを表示
//
function GoToOrder_onClick(){

	parent.location = "<%=g_HTTP%>shop/LoginCheck.asp?called_from=order";
}

//
// ====== 個別商品ページを表示
//
function Product_onClick(pItem){
//		parent.opener.location = "ProductDetail.asp?" + pItem;
		parent.location = "<%=g_HTTP%>shop/ProductDetail.asp?" + pItem;
}

</script>

</head>

<body bgcolor="#eeeeee" leftmargin="3" topmargin="3" marginwidth="0" marginheight="0">

<table class="honbun" border="0" width="100%">
  <tr bgcolor="#ffc600">
    <td align="left" valign="middle"><b>カートの内容</b></td>
  </tr>
</table>

<!-- 商品合計 -->
<table class="small" border="0" width="100%">
  <tr>
    <td width="100" align="left" nowrap>合計数量</td>
    <td align="right" nowrap><%=wTotalCnt%>個</td>
  </tr>
  <tr>
    <td width="100" align="left" nowrap>合計金額(税込)</td>
    <td align="right" nowrap><%=FormatNumber(wTotalAm,0)%>円</td>
  </tr>
</table>

<!-- 商品一覧 -->
<%=wListHTML%>

<!-- 注文へボタン、カートへボタン -->
<% if wTotalCnt > 0 then%>
<table border='0' width='100%'>
  <tr>
    <td align='center'><a href='JavaScript:GoToOrder_onClick();'><img src='images/GoToOrder.gif' border='0'></a></td>
  </tr>
  <tr>
    <td align='center'><a href='JavaScript:GoToCart_onClick();'><img src='images/GoToCart.gif' border='0'></a></td>
  </tr>
</table>
<% end if%>

</body>
</html>
