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
'	ショップ CFどっとクレジット インターフェース
'
'更新履歴
'2008/05/14 HTTPSチェック対応
'2008/05/23 入力データチェック強化（LEFT他)
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/07/13 if-web リニューアルレイアウト調整
'2013/10/01 GV # セディナシステム移行対応
'
'========================================================================

On Error Resume Next

Const store	= "160437002000000"	'CF store code 111111111111111 (for test)

Dim userID
Dim userName
Dim msg

Dim order_no
Dim customer_nm
Dim furigana
Dim customer_email
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim loan_am
Dim continue
Dim order_estimate

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add
Dim sedyna_url '2013/10/01 GV # add

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
	wErrDesc = "CfApply.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
order_estimate = ReplaceInput(Request("order_estimate"))

if IsNumeric(order_no) = false then
	Response.redirect g_HTTP
end if

'---- 受注情報取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.顧客名"
w_sql = w_sql & "       , a.顧客フリガナ"
w_sql = w_sql & "       , a.顧客E_mail1"
w_sql = w_sql & "       , b.顧客郵便番号"
w_sql = w_sql & "       , b.顧客都道府県"
w_sql = w_sql & "       , b.顧客住所"
w_sql = w_sql & "       , c.顧客電話番号"
w_sql = w_sql & "       , d.受注合計金額"
w_sql = w_sql & "       , d.ローン頭金"
w_sql = w_sql & "  FROM Web顧客 a WITH (NOLOCK)"
w_sql = w_sql & "     , Web顧客住所 b WITH (NOLOCK)"
w_sql = w_sql & "     , Web顧客住所電話番号 c WITH (NOLOCK)"
w_sql = w_sql & "     , Web受注 d WITH (NOLOCK)"
w_sql = w_sql & " WHERE d.受注番号 = " & order_no
w_sql = w_sql & "   AND a.顧客番号 = d.顧客番号"
w_sql = w_sql & "   AND b.顧客番号 = a.顧客番号"
w_sql = w_sql & "   AND b.住所連番 = 1"
w_sql = w_sql & "   AND c.顧客番号 = a.顧客番号"
w_sql = w_sql & "   AND c.住所連番 = 1"
w_sql = w_sql & "   AND c.電話連番 = 1"
	  
'@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

customer_nm = RS("顧客名")
furigana = RS("顧客フリガナ")
customer_email = RS("顧客E_mail1")
zip = RS("顧客郵便番号")
prefecture = RS("顧客都道府県")
address = RS("顧客住所")
telephone = RS("顧客電話番号")
loan_am = RS("受注合計金額") - RS("ローン頭金") 

RS.close

if order_estimate = "ご注文" then
	continue = "1"
	'2013/10/01 GV # add start
	'注文受付
	sedyna_url = "https://c-web.cedyna.co.jp/customer/action/ssAA01/WAA0101Action/RWAA010101"
	'2013/10/01 GV # add end
else
	continue = "0"
	'2013/10/01 GV # add start
	'シミュレート単体
	sedyna_url = "https://c-web.cedyna.co.jp/customer/action/ssAA01/WAA0106Action/RWAA010601"
	'2013/10/01 GV # add end
end if

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
<title>セディナ呼出中｜サウンドハウス</title>
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
      <li class="now">セディナ呼出し中</li>
    </ul>
  </div></div></div>

  <h1 class="title">セディナ呼出し中</h1>

  <p>しばらくお待ちください。</p>
<%
' 2013/10/01 GV # mod 
'  <form name="f_cf" method="post" action="https://cf.ufit.ne.jp/dotcredit/simulate/simulate.asp">
%>
  <form name="f_cf" method="post" action="<%=sedyna_url%>">
    <input type="hidden" name="store" value="<%=store%>">
    <input type="hidden" name="amount" value="<%=loan_am%>">
    <input type="hidden" name="continue" value="<%=continue%>">
    <input type="hidden" name="labor" value="0">
    <input type="hidden" name="item1" value="音響機器">
    <input type="hidden" name="item1count" value="1">
    <input type="hidden" name="item1amount" value="<%=loan_am%>">
    <input type="hidden" name="tranno" value="<%=order_no%>">
    <input type="hidden" name="namekn" value="<%=furigana%>">
    <input type="hidden" name="namekj" value="<%=customer_nm%>">
    <input type="hidden" name="zip" value="<%=zip%>">
    <input type="hidden" name="address" value="<%=prefecture%><%=address%>">
    <input type="hidden" name="tel" value="<%=telephone%>">
    <input type="hidden" name="mail" value="<%=customer_email%>">
    <input type="hidden" name="bonusdeal" value="100">
    <input type="hidden" name="twobonusdeal" value="200">
    <input type="hidden" name="result" value="1">
  </form>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	document.f_cf.submit();		//CFページへジャンプ
</script>
</body>
</html>