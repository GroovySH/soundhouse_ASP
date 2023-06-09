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
'	ショップ オリコ クリックオン インターフェース
'
'更新履歴
'2008/05/14 HTTPSチェック対応
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/07/13 if-web リニューアルレイアウト調整
'
'========================================================================

On Error Resume Next

Const kamei_no	= "06218226"	'加盟店番号
'Const kyaku_syu	= "004"	'取扱契約番号
Const kyaku_syu	= "005"	'取扱契約番号	'2005/07/05 na mod
Const buten_cd	= "519"	'部店コード
Const OricoURL	= "https://www2.orico.co.jp/webcredit/sp/top.asp"	'本番
'@@@Const OricoURL	= "https://www2.orico.co.jp/webcredit/sp/simulation.asp"	'テスト

Dim userID
Dim userName
Dim msg

Dim order_no
Dim order_estimate

Dim CustomerName
Dim CustomerEmail
Dim Zip
Dim Prefecture
Dim Address
Dim Telephone
Dim ProductAm
Dim ShippingAm
Dim DownPaymentAm
Dim OrderTotalAm
Dim Continue
Dim SalesTaxRate

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_error_msg
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
	wErrDesc = "OricoApply.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
order_no = Clng(ReplaceInput(Request("order_no")))
order_estimate = ReplaceInput(Request("order_estimate"))

if IsNumeric(order_no) = false then
	Response.redirect g_HTTP
end if

'---- 受注情報取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.顧客名"
w_sql = w_sql & "     , a.顧客E_mail1"
w_sql = w_sql & "     , b.顧客郵便番号"
w_sql = w_sql & "     , b.顧客都道府県"
w_sql = w_sql & "     , b.顧客住所"
w_sql = w_sql & "     , c.顧客電話番号"
w_sql = w_sql & "     , d.商品合計金額"
w_sql = w_sql & "     , d.送料"
w_sql = w_sql & "     , d.受注合計金額"
w_sql = w_sql & "     , d.ローン頭金"
w_sql = w_sql & "     , d.消費税率"
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

CustomerName = RS("顧客名")
CustomerEmail = RS("顧客E_mail1")
Zip = RS("顧客郵便番号")
Prefecture = RS("顧客都道府県")
Address = RS("顧客住所")
Telephone = RS("顧客電話番号")
SalesTaxRate = Ccur(RS("消費税率"))
ShippingAm = Fix(RS("送料") * (100 + SalesTaxRate) / 100)
DownPaymentAm = RS("ローン頭金")
OrderTotalAm = RS("受注合計金額")
ProductAm = OrderTotalAm - ShippingAm		'消費税調整のため商品合計金額は使用しない

RS.close

if order_estimate = "ご注文" then
	Continue = "1"
else
	Continue = "0"
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
<title>オリコ呼出中｜サウンドハウス</title>
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
      <li class="now">オリコ呼出中</li>
    </ul>
  </div></div></div>

  <h1 class="title">オリコ呼出中</h1>

  <p>しばらくお待ちください。</p>

  <form name="f_cf" method="post" action="<%=OricoURL%>">
    <input type="hidden" name="kamei_no" value="<%=kamei_no%>">
    <input type="hidden" name="kyaku_syu" value="<%=kyaku_syu%>">
    <input type="hidden" name="buten_cd" value="<%=buten_cd%>">
    <input type="hidden" name="back_url" value="http://www.soundhouse.co.jp">
    <input type="hidden" name="pr_num" value="<%=order_no%>">
    <input type="hidden" name="brand_mei1" value="音響機器">
    <input type="hidden" name="brand_suu1" value="1">
    <input type="hidden" name="brand_kin1" value="<%=ProductAm%>">
    <input type="hidden" name="brand_gokei" value="<%=ProductAm%>">
    <input type="hidden" name="soryo_gokei" value="<%=ShippingAm%>">
    <input type="hidden" name="loan_kin" value="<%=OrderTotalAm%>">
    <input type="hidden" name="head_kin" value="<%=DownPaymentAm%>">
    <input type="hidden" name="h_name" value="<%=CustomerName%>">
    <input type="hidden" name="h_yubin" value="<%=Zip%>">
    <input type="hidden" name="h_addr1" value="<%=Prefecture & Address%>">
    <input type="hidden" name="h_addr2" value="">
    <input type="hidden" name="h_telno" value="<%=Telephone%>">
  </form>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	document.f_cf.submit();		//Oricoページへジャンプ
</script>
</body>
</html>