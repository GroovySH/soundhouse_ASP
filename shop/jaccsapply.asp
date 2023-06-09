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
<!--#include file="../common/HttpsSecurity.inc"-->
<%
'========================================================================
'
'	JACCS インターフェース
'
'更新履歴
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/07/13 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Dim userID
Dim userName
Dim msg

Dim order_no
Dim customer_nm
Dim customer_email

Dim ShippingAm
Dim SalesTaxAm
Dim OrderTotalAm
Dim ProductName(30)
Dim ProductPrice(30)
Dim ProductQt(30)
Dim ProductCnt

Dim wProductHTML

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.Expires = -1			' Do not cache

'---- UserID 取り出し
userID = Session("userID")
userName = Session("userName")

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "jaccsapply.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

'---- 受信データーの取り出し
order_no = ReplaceInput(Request("order_no"))

if IsNumeric(order_no) = false then
	Response.redirect g_HTTP
end if

'---- 受注情報取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.顧客名"
w_sql = w_sql & "     , a.顧客E_mail1"
w_sql = w_sql & "     , d.送料"
w_sql = w_sql & "     , d.外税合計金額"
w_sql = w_sql & "     , d.受注合計金額"
w_sql = w_sql & "     , e.メーカー名"
w_sql = w_sql & "     , e.商品名"
w_sql = w_sql & "     , e.色"
w_sql = w_sql & "     , e.規格"
w_sql = w_sql & "     , e.受注数量"
w_sql = w_sql & "     , e.受注単価"
w_sql = w_sql & "  FROM Web顧客 a WITH (NOLOCK)"
w_sql = w_sql & "     , Web受注 d WITH (NOLOCK)"
w_sql = w_sql & "     , Web受注明細 e WITH (NOLOCK)"
w_sql = w_sql & " WHERE d.受注番号 = " & order_no
w_sql = w_sql & "   AND a.顧客番号 = d.顧客番号"
w_sql = w_sql & "   AND e.受注番号 = d.受注番号"
w_sql = w_sql & " ORDER BY e.受注明細番号"
	  
'@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

customer_nm = RS("顧客名")
customer_email = RS("顧客E_mail1")
ShippingAm = RS("送料")
SalesTaxAm = RS("外税合計金額")
OrderTotalAm = RS("受注合計金額")

wProductHTML = ""
ProductCnt = 0

Do until RS.EOF = true or ProductCnt = 30
	ProductCnt = ProductCnt + 1
	wProductHTML = wProductHTML & "<input type='hidden' name='SINA_" & Right("0" & CStr(ProductCnt), 2) & "_SINAMEI' value='" & RS("商品名") & "'>" & vbNewLine
	wProductHTML = wProductHTML & "<input type='hidden' name='SINA_" & Right("0" & CStr(ProductCnt), 2) & "_SURYOU' value='" & RS("受注数量") & "'>" & vbNewLine
	wProductHTML = wProductHTML & "<input type='hidden' name='SINA_" & Right("0" & CStr(ProductCnt), 2) & "_TANKA' value='" & FIX(RS("受注単価")) & "'>" & vbNewLine & vbNewLine
	RS.MoveNext
Loop

RS.close

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
<title>JACCS呼出中｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
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
      <li class="now">JACCS呼出し中</li>
    </ul>
  </div></div></div>

  <h1 class="title">JACCS呼出し中</h1>

  <p>しばらくお待ちください。</p>

  <form name="f_JACCS" method="post" action="<%=g_JACCS_URL%>">
    <input type="hidden" name="KAMEITEN_BANGO" value="<%=g_JACCS_ID%>"> <!-- 加盟店識別番号 -->
    <input type="hidden" name="NAME" value="<%=customer_nm%>"> <!-- 氏名 -->
    <input type="hidden" name="EMAIL" value="<%=customer_email%>"> <!-- メールアドレス -->
    <input type="hidden" name="DENPYO_BANGO" value="<%=order_no%>"> <!-- 伝票番号 -->

<!-- 商品情報　品名、数量、単価 -->
<%=wProductHTML%>

    <input type="hidden" name="SYOHIZEI" value="<%=SalesTaxAm%>"> <!-- 消費税 -->
    <input type="hidden" name="SOURYOU" value="<%=ShippingAm%>"> <!-- 送料 -->
    <input type="hidden" name="GOUKEIGAKU" value="<%=OrderTotalAm%>"> <!-- 合計金額（税込） -->
  </form>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	document.f_JACCS.submit();		//JACCSページへジャンプ
</script>
</body>
</html>