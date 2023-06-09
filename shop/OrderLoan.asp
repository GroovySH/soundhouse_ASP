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

'	ローンページ
'
'2012/06/18 ok デザイン変更のため旧版を元に新規作成
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache

'---- Session情報
Dim wUserID
Dim wUserName
Dim wMsg

Dim wLoanDownPaymentFl
Dim wLoanDownPaymentAm
Dim wLoanTermPayment
Dim wLoanTerm
Dim wLoanAm
Dim wLoanApplyFl
Dim wLoanCompany
Dim wErrDesc

'---- DB
Dim Connection

'=======================================================================
'	受け渡し情報取り出し
'=======================================================================
'---- Session変数
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session("msg")

'---- 受け渡し情報取り出し

Session("msg") = ""

'---- セッション切れチェック
If wUserID = ""Then
	Response.Redirect g_HTTP
End If

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()
Call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderLoan.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection= Nothing    '2011/08/01 an add

End Function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

' ローン情報の取得
Call GetLoanInfo()

End Function

'========================================================================
'
'	Function	ローン情報取得
'
'========================================================================
Function GetLoanInfo

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    ローン頭金ありフラグ"
vSQL = vSQL & "  , ローン頭金"
vSQL = vSQL & "  , 希望ローン回数"
vSQL = vSQL & "  , ローン金額"
vSQL = vSQL & "  , オンラインローン申込フラグ"
vSQL = vSQL & "  , ローン会社"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    仮受注"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

wLoanDownPaymentFl  = RSv("ローン頭金ありフラグ")
wLoanDownPaymentAm = RSv("ローン頭金")
wLoanTerm = RSv("希望ローン回数")
wLoanAm = RSv("ローン金額")
If wLoanAm <> 0 Then
	wLoanTermPayment = "P"
Else
	wLoanTermPayment = "T"
End If
wLoanApplyFl = RSv("オンラインローン申込フラグ")
wLoanCompany = RSv("ローン会社")

RSv.Close

End Function

'========================================================================

%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>ローンのお申し込み｜サウンドハウス</title>
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
<!--#include file="../Navi/NaviStyle.inc"-->

<script type="text/javascript">
//=====================================================================
//	次へボタン onClick
//=====================================================================
function Next_onClick() {
	document.f_data.action = "OrderLoanStore.asp";
	document.f_data.submit();
}
//=====================================================================
//	キャンセルボタン onClick
//=====================================================================
function Cancel_onClick() {
	document.f_data.action = "OrderInfoEnter.asp";
	document.f_data.submit();
}
//=====================================================================
//	ラジオボタン、ドロップダウンリストの選択
//=====================================================================
function preset_values(){

	// 頭金あり／なし
	if (document.f_data.i_loan_downpayment_fl.value == "Y"){
		document.f_data.loan_downpayment_fl[1].checked = true;
	}
	if (document.f_data.i_loan_downpayment_fl.value == "N"){
		document.f_data.loan_downpayment_fl[0].checked = true;
	}

	// オンラインで申し込む／使用しない
	if (document.f_data.i_loan_apply_fl.value == "Y"){
		document.f_data.loan_apply_fl[0].checked = true;
		// ローン会社
		if (document.f_data.i_loan_company.value == "ジャックス"){
			document.f_data.loan_company[0].checked = true;
		}
		if (document.f_data.i_loan_company.value == "セディナ"){
			document.f_data.loan_company[1].checked = true;
		}
	}
	if (document.f_data.i_loan_apply_fl.value == "N") {
		document.f_data.loan_apply_fl[1].checked = true;
		// 希望ローン回数／月額支払金額
		if (document.f_data.i_loan_term_payment.value == "T"){
			document.f_data.loan_term_payment[0].checked = true;
			for (var i=0; i<document.f_data.loan_term.length; i++){
				if (document.f_data.loan_term[i].value == document.f_data.i_loan_term.value){
					document.f_data.loan_term.options[i].selected = true;
					break;
				}
			}
		}
		if (document.f_data.i_loan_term_payment.value == "P"){
			document.f_data.loan_term_payment[1].checked = true;
		}
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

  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li>お届け先、お支払い方法の選択</li>
      <li class="now">ローンのお申し込み</li>
    </ul>
  </div></div></div>

  <h1 class="title">お届け先、お支払い方法の選択</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.ショッピングカート" width="170" height="50"></li>
    <li><img src="images/step02_now.gif" alt="2.お届け先、お支払方法の選択" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.ご注文内容の確認" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.ご注文完了" width="170" height="50"></li>
  </ol>

  <p class="error"><% = wMsg %></p>

  <h2 class="cart_title">ローンのお申し込み</h2>        

  <form name="f_data" method="post">
    <table id="address">
      <tr>
        <td class="main">
          <ul class="loan_choice">
            <li><input id="loan_downpayment_fl_n" name="loan_downpayment_fl" type="radio" value="N"><label for="loan_downpayment_fl_n">頭金なし</label></li>
            <li><input id="loan_downpayment_fl_y" name="loan_downpayment_fl" type="radio" value="Y"><label for="loan_downpayment_fl_y">頭金あり</label></li>
          </ul>
          <ul>
            <li>頭金<input name="loan_downpayment_am" type="text" value="<% = wLoanDownPaymentAm %>" size="12" class="field_r">円</li>
          </ul>
        </td>
      </tr>
      <tr>
        <td class="main">
          <span class="loan_choice"><input id="loan_apply_fl_y" name="loan_apply_fl" type="radio" value="Y"><label for="loan_apply_fl_y" class="radio_strong">オンラインでローンを申し込む</label></span>
          <ul>
            <li><input id="loan_company_1" name="loan_company" type="radio" value="ジャックス"><label for="loan_company_1"><img src="images/jaccs.gif" alt="ジャックス"></label><label for="loan_company_1">ジャックス</label></li>
            <li><input id="loan_company_2" name="loan_company" type="radio" value="セディナ"><label for="loan_company_2"><img src="images/cedyna.gif" alt="セディナ"></label><label for="loan_company_2">セディナ</label></li>
          </ul>
          <ul class="attention">
            <li>オンラインローンの場合､お申込後のご注文内容の変更を承ることができません。</li>
            <li>ご注文内容と､オンラインローン申込フォームの内容をご確認の上、ご注文ください。</li>
            <li>ジャックスでお申し込みの場合は、頭金なしとなります。</li>
          </ul>
        </td>
      </tr>
      <tr>
        <td class="main">
          <span class="loan_choice"><input id="loan_apply_fl_n" name="loan_apply_fl" type="radio" value="N"><label for="loan_apply_fl_n" class="radio_strong">オンラインを使用しない</label>（ローン回数または月額を指定してください）</span>
          <ul>
            <li><input id="loan_term_payment_t" name="loan_term_payment" type="radio" value="T"><label for="loan_term_payment_t">希望ローン回数</label>
              <select name="loan_term" size="1">
                <option value="0"></option>
                <option value="1">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="6">6</option>
                <option value="10">10</option>
                <option value="12">12</option>
                <option value="15">15</option>
                <option value="18">18</option>
                <option value="20">20</option>
                <option value="24">24</option>
                <option value="30">30</option>
                <option value="36">36</option>
                <option value="42">42</option>
                <option value="48">48</option>
                <option value="54">54</option>
                <option value="60">60</option>
              </select>
            </li>
            <li><input id="loan_term_payment_p" name="loan_term_payment" type="radio" value="P"><label for="loan_term_payment_p">月額支払金額</label><input name="loan_am" type="text" value="<% = wLoanAm %>" size="12" class="field_r">円</li>
          </ul>
          <ul class="attention">
            <li>ローン会社によりご希望のお支払い回数を指定できない場合がございます。</li>
          </ul>
        </td>
      </tr>
    </table>
    <div id="btn_box">
      <ul class="btn">
        <li><a href="javascript:Cancel_onClick();"><img src="images/btn_back.png" alt="戻る" class="opover"></a></li>
        <li class="last"><a href="javascript:Next_onClick();"><img src="images/btn_next.png" alt="次へ" class="opover"></a></li>
      </ul>
    </div>
    <input type="hidden" name="i_loan_downpayment_fl" value="<% = wLoanDownPaymentFl %>">
    <input type="hidden" name="i_loan_apply_fl" value="<% = wLoanApplyFl %>">
    <input type="hidden" name="i_loan_company" value="<% = wLoanCompany %>">
    <input type="hidden" name="i_loan_term_payment" value="<% = wLoanTermPayment %>">
    <input type="hidden" name="i_loan_term" value="<% = wLoanTerm %>">
  </form>

<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	preset_values();
</script>
</body>
</html>