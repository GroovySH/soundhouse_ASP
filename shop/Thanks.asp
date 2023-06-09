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
<!--#include file="../3rdParty/EAgency.inc"-->

<%
'========================================================================
'
'	オーダーありがとうございましたページ
'
'2012/06/18 ok デザイン変更のため旧版を元に新規作成
'2012/08/24 ok インタースペース アフィリエイトプログラム用タグを新版に変更
'2013/05/20 GV #1505 さぶみっと！レコメンド対応
'2013/07/30 GV #1618 アフィリエイト重複送信対応
'
'========================================================================

On Error Resume Next

Dim userID
Dim userName
Dim msg

Dim w_order_estimate
Dim payment_method
Dim loan_company
Dim order_no
Dim product_am
Dim w_thanks_msg
Dim w_shiharai_about

'---- UserID 取り出し
userID = Session("userID")
userName = Session("userName")

'---- Get input data
msg = Session.contents("msg")
Session("msg") = ""

'---- Execute main
call main()

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

'2013/07/30 GV #1618 start
'1つ前のOrderSubmit.asp で仕込んだ値が存在しない場合、トップページへ遷移
Dim OrderAtOnce
OrderAtOnce = Session("OrderAtOnce")
If ((OrderAtOnce = "") Or (OrderAtOnce <> "1")) Then
	Response.Redirect g_HTTP
Else
	Session.Contents.Remove("OrderAtOnce")
End If
'2013/07/30 GV #1618 end

w_order_estimate = ReplaceInput(Request("order_estimate"))
payment_method = ReplaceInput(Request("payment_method"))
loan_company = ReplaceInput(Trim(Request("loan_company")))
order_no = ReplaceInput(Request("order_no"))
product_am = ReplaceInput(Request("product_am"))

w_thanks_msg = ""
w_thanks_msg = w_thanks_msg & "ご利用ありがとうございます。<br>"
if payment_method = "銀行振込" or payment_method = "コンビニ支払" then
	w_thanks_msg = w_thanks_msg & "ご依頼いただきました内容の受付確認メールを自動送信いたしました。<br>"
	w_thanks_msg = w_thanks_msg & "その後、別途お見積りをご案内いたしますので、内容をご確認ください。<br>"
else
	w_thanks_msg = w_thanks_msg & "ご依頼いただきました内容の受付確認メールを即時に送信いたします。<br>"
	w_thanks_msg = w_thanks_msg & "その後、別途ご注文確認書を送信いたしますのでご確認ください。<br><br>"
end if
w_thanks_msg = w_thanks_msg & "次回もぜひサウンドハウスをご利用くださいませ。"

w_shiharai_about = ""
if payment_method = "銀行振込" or payment_method = "コンビニ支払" then
	w_shiharai_about = w_shiharai_about & "  <dl class='about'>"
	w_shiharai_about = w_shiharai_about & "    <dt>お支払いについて</dt>"
	w_shiharai_about = w_shiharai_about & "    <dd>"

	if payment_method = "銀行振込" then
		w_shiharai_about = w_shiharai_about & "銀行振込をご利用の場合は、別途ご案内いたします 「サウンドハウス お見積書」メールをご確認の上、電信扱いにてお振込みください。<br>"
		w_shiharai_about = w_shiharai_about & "（文書扱いの場合は、ご入金確認までお時間がかかります。）<br>"
		w_shiharai_about = w_shiharai_about & "ご入金を確認後、商品を発送いたします。"
	else
		w_shiharai_about = w_shiharai_about & "ネットバンキング・ゆうちょ・コンビニ払いをご利用の場合は、別途ご案内いたします 「サウンドハウス お見積書」メールをご確認の上、メール内に記載されている専用URLにアクセスしてください。<br>"
		w_shiharai_about = w_shiharai_about & "ご希望のお支払方法を選択し、お支払い受付番号又は収納機関番号/確認番号（ゆうちょ銀行）をご確認の上、表記しております期日までにお支払いください。 <br>"
		w_shiharai_about = w_shiharai_about & "ご入金確認後、商品を発送いたします。"
	end if

	w_shiharai_about = w_shiharai_about & "    <p><a href='http://guide.soundhouse.co.jp/guide/oshiharai.asp'>お支払いについて</a></p>"
	w_shiharai_about = w_shiharai_about & "    </dd>"
	w_shiharai_about = w_shiharai_about & "  </dl>"
end if

End Function

'========================================================================
%>

<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>ご注文ありがとうございました｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/StyleOrder.css" type="text/css">
<script type="text/javascript">
<% if loan_company = "セディナ" then %>
	window.open("CFApply.asp?order_no=<%=order_no%>&order_estimate=<%=Server.URLEncode(w_order_estimate)%>")
<% end if %>
<% if loan_company = "ジャックス" then %>
	window.open("JACCSApply.asp?order_no=<%=order_no%>")
<% end if %>
</script>
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
      <li class="now">ご注文完了</li>
    </ul>
  </div></div></div>

  <h1 class="title">ご注文完了</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.ショッピングカート" width="170" height="50"></li>
    <li><img src="images/step02.gif" alt="2.お届け先、お支払方法の選択" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.ご注文内容の確認" width="170" height="50"></li>
    <li><img src="images/step04_now.gif" alt="4.ご注文完了" width="170" height="50" /></li>
  </ol>

  <div id="thanks">
    <p><strong>THANK YOU!</strong></p>
    <p><%= w_thanks_msg %></p>
    <img src="images/ojigi-2.gif" alt="" height="300" width="150">
  </div>
<%
'2013/05/20 GV #1505
fEAgency_CreateRecommendOrderSubmitJS(order_no)
%>

<%= w_shiharai_about %>

  <dl class="about">
    <dt>商品の納期について</dt>
    <dd>
      <ul>
        <li>ウェブサイト上、およびご注文やお見積り時点でご案内しております納期につきましては、あくまでも予定となっており、諸事情により変更となる場合がございます。</li>
        <li>商品の納期につきましては、メールやお電話でのお問い合わせも承っております。指定日までに納品が必要なご注文は、遠慮なく事前にご相談ください。</li>
        <li>なお、納期遅延によって生じた問題につきましては、当社では一切の責を負うことができません。 あらかじめご了承ください。</li>
      </ul>
      <p><a href="http://guide.soundhouse.co.jp/guide/kaimono.asp#nissuu">商品の納期について</a></p>
    </dd>
  </dl>
  <dl class="about">
    <dt>ご購入後のサポートについて</dt>
    <dd>
      <ul>
        <li>商品がお手元に届きましたら、すぐに納品書と商品内容および数量をご確認ください。</li>
        <li>万一商品が破損していたり注文と異なる商品だった場合は、すぐにご連絡ください。</li>
        <li>商品到着後、一週間以上経過してからお申し出があった場合、受付できないことがございます。</li>
      </ul>
      <p><a href="http://guide.soundhouse.co.jp/guide/support.asp">ご購入後のサポートについて</a></p>
    </dd>
  </dl>
  
<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->

<% if w_order_estimate = "ご注文" then %>
<!-- インタースペース アフィリエイトプログラム用タグ -->
	<img src="https://is.accesstrade.net/cgi-bin/isatV2/soundhouse/isatWeaselV2.cgi?result_id=2&verify=<%=order_no%>&value=<%=product_am%>" width="1" height="1" />
<% end if %>

<!--#include file="../Navi/NaviScript.inc"-->
</body>
<!-- SmarterJP Conversion Code -->
<!--#include file="../3rdParty/SmarterMerchantOrder.class.asp"-->
<%
Dim oSMO
Dim oRtnCode

set oSMO = new SmarterMerchantOrder

oSMO.MerchantID = "SM1201A10083"		'Merchant ID (SH)
oSMO.Key = "sh10083050206"			'the Key
oSMO.OrderNum = order_no		'Order Number
oSMO.OrderAmount = product_am	'Price Total

oRtnCode = oSMO.send

%>
</html>
