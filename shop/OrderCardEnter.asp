<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
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
'	カード入力、確認ページ
'
'
'========================================================================

On Error Resume Next

Dim userID
Dim userName
Dim w_SessionID

Dim payment_method
Dim Skey

Dim CardCompany
Dim CardNo
Dim CardExpMM
Dim CardExpYY
Dim CardName
Dim CardHoji

Dim Connection
Dim RS

Dim wSQL
Dim wMSG
Dim wHTML

'=======================================================================

Response.Expires = -1			' Do not cache

'---- UserID 取り出し
userID = Session("userID")
userName = Session("userName")
w_sessionID = Session.SessionID

'---- 入力データーの取り出し
wMSG = Session("msg")
Session("msg") = ""

payment_method = ReplaceInput(Request("payment_method"))
if payment_method <> "クレジットカード" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

Skey = ReplaceInput(Request("Skey"))

CardCompany = ReplaceInput(Request("CardCompany"))
CardNo = ReplaceInput(Request("CardNo"))
CardExpMM = ReplaceInput(Request("CardExpMM"))
CardExpYY = ReplaceInput(Request("CardExpYY"))
CardName = ReplaceInput(Request("CardName"))
CardHoji = ReplaceInput(Request("CardHoji"))

'---- メイン処理
if wMSG = "" then
	call connect_db()
	call main()
	call close_db()
end if

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'=======================================================================

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

'---- カード情報チェック
wSQL = ""
wSQL = wSQL & "SELECT カード会社"
wSQL = wSQL & "     , カード番号"
wSQL = wSQL & "     , カード有効期限"
wSQL = wSQL & "     , カード名義人"
wSQL = wSQL & "  FROM Web顧客"
wSQL = wSQL & " WHERE 顧客番号 = '" & UserID & "'"
  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = true then
	wMSG = "処理が異常終了しました。"
	exit function
end if

if RS("カード番号") <> "" then
	CardCompany = RS("カード会社")
	CardNo = RS("カード番号")
	CardExpMM = left(RS("カード有効期限"),2)
	CardExpYY = right(RS("カード有効期限"),2)
	CardName = RS("カード名義人")
	CardHoji = "Y"
end if

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

End function

'========================================================================
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>サウンドハウス  ご注文受付　カード</title>

<!--#include file="../Navi/NaviStyle.inc"-->

<script language="JavaScript">
//=====================================================================
//	ラジオボタン、ドロップダウンリストを以前に選択した状態にする
//=====================================================================
function preset_values(){

// カード会社
	for (var i=0; i<document.fData.CardCompany.options.length; i++){
		if (document.fData.CardCompany.options[i].value == document.fData.iCardCompany.value)		{
			document.fData.CardCompany.options[i].selected = true;
			break;
		}
	}

//	カード有効期限
	for (var i=0; i<document.fData.CardExpMM.options.length; i++){
		if (document.fData.CardExpMM.options[i].value == document.fData.iCardExpMM.value)		{
			document.fData.CardExpMM.options[i].selected = true;
			break;
		}
	}

	for (var i=0; i<document.fData.CardExpYY.options.length; i++){
		if (document.fData.CardExpYY.options[i].value == document.fData.iCardExpYY.value)		{
			document.fData.CardExpYY.options[i].selected = true;
			break;
		}
	}

// カード保持
	if (document.fData.iCardHoji.value == "Y"){
		document.fData.CardHoji[0].checked = true;
	}
}

</script>

</head>

<body background="../Navi/Images/back_ground.gif" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!--#include file="../Navi/NaviTop.inc"-->

<table width="940" height="26" border="0" cellpadding="0" cellspacing="0">
  <tr>

<!--#include file="../Navi/NaviLeft.inc"-->

    <td width="798" align="left" valign="top" bgcolor="#ffffff">

<!------------ ページメイン部分の記述 START ------------>

<!-- エラーメッセージ -->

<% if wMSG <> "" then %>
  <table width="99%" border="1" cellspacing="0" cellpadding="3" bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff" >
    <tr align="center" valign="top" class="honbun">
      <td align="left" bgcolor="#D2FFFF"><font color="#FF0000"><%=wMSG%></font></td>
    </tr>
  </table>
  <br>
 <% end if %>

      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="100%" align="center" valign="middle" bgcolor="#fff2e6">

<!-- カード情報 -->
            <form name="fData" method="post" action="OrderCardUpdate.asp">
            <table width="100%" height="220" border="0" cellpadding="3" cellspacing="3">
              <tr align="left" class="honbun">
                <td>カード会社</td>
                <td align="left" valign="middle">
                  <select name="CardCompany">
                    <option value="" SELECTED>
                    <option value="VISA">ビザ
                    <option value="MASTER CARD">マスターカード
                    <option value="AMEX">アメリカンエクスプレス
                    <option value="DC">DC カード
                  </select>
                </td>
              </tr>
              <tr align="left" class="honbun">
                <td>カード番号</td>
                <td align="left" valign="middle">
                  <input name="CardNo" type="text" size="25" maxsize="20" value="<%=CardNo%>">
                </td>
              </tr>
              <tr align="left" class="honbun">
                <td>カード有効期限</td>
                <td align="left" valign="middle">
                  <select name="CardExpMM">
                    <option value="" SELECTED>
                    <option value="01">01
                    <option value="02">02
                    <option value="03">03
                    <option value="04">04
                    <option value="05">05
                    <option value="06">06
                    <option value="07">07
                    <option value="08">08
                    <option value="09">09
                    <option value="10">10
                    <option value="11">11
                    <option value="12">12
                  </select>
									月
                  <select name="CardExpYY">
                    <option value="" SELECTED>
                    <option value="08">2008
                    <option value="09">2009
                    <option value="10">2010
                    <option value="11">2011
                    <option value="12">2012
                    <option value="13">2013
                    <option value="14">2014
                    <option value="15">2015
                    <option value="16">2016
                    <option value="17">2017
                  </select>
									年
                </td>
              </tr>
              <tr align="left" class="honbun">
                <td>カード名義</td>
                <td align="left" valign="middle">
                  <input name="CardName" type="text" size="25" maxsize="60" value="<%=CardName%>">
                </td>
              </tr>
              <tr align="left" class="honbun">
                <td>カード情報を保持する</td>
                <td align="left" valign="middle">
                  <input name="CardHoji" type="radio" value="Y">する　
                  <input name="CardHoji" type="radio" value="N" CHECKED>しない
                </td>
              </tr>
              <tr align="left" class="honbun">
                <td colspan="2" align="center">
									<input type="submit" value="このカードを使用する">
                </td>
              </tr>
						</table>
						<input type="hidden" name="iCardCompany" value="<%=CardCompany%>">
						<input type="hidden" name="iCardExpMM" value="<%=CardExpMM%>">
						<input type="hidden" name="iCardExpYY" value="<%=CardExpYY%>">
						<input type="hidden" name="iCardHoji" value="<%=CardHoji%>">

		        <input type="hidden" name="payment_method" value="<%=payment_method%>">
		        <input type="hidden" name="Skey" value="<%=Skey%>">

            </form>

          </td>
        <td align="center" valign="middle">&nbsp;</td>
      </tr>
    </table>

<!------------ ページメイン部分の記述 END ------------>

    </td>
  </tr>
</table>

<!--#include file="../Navi/NaviBottom.inc"-->

<!--#include file="../Navi/NaviScript.inc"-->

</body>
</html>

<script language="JavaScript">

	preset_values();

</script>

