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
'	カード登録
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

Dim NextURL

Dim wSQL
Dim wMSG
Dim wHTML

Dim Degub

'=======================================================================

Response.Expires = -1			' Do not cache
Response.Buffer = true

'---- セキュリティーキーセット 
payment_method = ReplaceInput(Request("payment_method"))
if payment_method <> "クレジットカード" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

Skey = ReplaceInput(Request("Skey"))

'---- UserID 取り出し
userID = Session("userID")
userName = Session("userName")
w_sessionID = Session.SessionID

'---- 入力データーの取り出し
CardCompany = ReplaceInput(Trim(Request("CardCompany")))
CardNo = ReplaceInput(Trim(Request("CardNo")))
CardExpMM = ReplaceInput(Trim(Request("CardExpMM")))
CardExpYY = ReplaceInput(Trim(Request("CardExpYY")))
CardName = ReplaceInput(Trim(Request("CardName")))
CardHoji = ReplaceInput(Trim(Request("CardHoji")))

'---- メイン処理
call connect_db()
call main()
call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp" & Err.Description
end if

Session("msg") = ""
'---- エラーが無いときはOrderProcessing、エラーがあればカード入力ページへ
if wMSG = "" then
	NextURL = "OrderProcessing.asp"
else
	NextURL = "OrderCardEnter.asp"
	Session("msg") = wMSG
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

Dim vOldCardNo
Dim Campus

'---- 入力情報のチェック
Call ValidateData()

if wMSG <> "" then
	exit function
end if

'---- カード情報更新
wSQL = ""
wSQL = wSQL & "SELECT カード会社"
wSQL = wSQL & "     , カード番号"
wSQL = wSQL & "     , カード有効期限"
wSQL = wSQL & "     , カード名義人"
wSQL = wSQL & "  FROM Web顧客"
wSQL = wSQL & " WHERE 顧客番号 = " & UserID
  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = true then
	wMSG = "処理が異常終了しました。"
	exit function
end if

vOldCardNo = RS("カード番号")
if IsNull(vOldCardNo) = true then
	vOldCardNo = ""
end if

if vOldCardNo <> CardNo then
	if isNumeric(CardNo) = false then
		wMSG = "カード番号は数字のみで入力願います。"
		exit function
	end if
end if

if CardHoji = "Y" then
	RS("カード会社") = CardCompany
	RS("カード番号") = "************" & Right(CardNo, 4)
	RS("カード有効期限") = CardExpMM & "/" & CardExpYY
	RS("カード名義人") = CardName
else
	RS("カード会社") = ""
	RS("カード番号") = ""
	RS("カード有効期限") = ""
	RS("カード名義人") = ""
end if

RS.update
RS.close

'---- カード情報更新2(カード番号登録変更時）
if vOldCardNo <> CardNo then
	Set Campus = Server.CreateObject("WebCampusAccess.WebCampus")

	Campus.Site = g_RegForder
	Campus.CustomerNo = UserID
	Campus.CardNo = CardNo
	Campus.CardExpDt = CardExpMM & "/" & CardExpYY

	Campus.StoreCardNo()
end if

End function

'========================================================================
'
'	Function	入力データーのチェック
'
'========================================================================
'
Function ValidateData()

wMSG = ""
'---- カード会社
if CardCompany = "" then
	wMSG = wMSG & "カード会社を選択願います。<br>"
end if

'---- カード番号
if CardNo = "" then
	wMSG = wMSG & "カード番号を入力願います。<br>"
end if

'---- カード有効期限
if CardExpMM = "" OR CardExpYY = "" then
	wMSG = wMSG & "カード有効期限を選択願います。<br>"
end if

'---- カード名義
if CardName = "" then
	wMSG = wMSG & "カード名義を入力願います。<br>"
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

End function

'========================================================================
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>サウンドハウス  ご注文受付　カード</title>

</head>

<body>

<form name="fData" method="post" action="<%=NextURL%>">
<input type="hidden" name="CardCompany" value="<%=CardCompany%>">
<input type="hidden" name="CardNo" value="">
<input type="hidden" name="CardExpMM" value="<%=CardExpMM%>">
<input type="hidden" name="CardExpYY" value="<%=CardExpYY%>">
<input type="hidden" name="CardName" value="<%=CardName%>">
<input type="hidden" name="CardHoji" value="<%=CardHoji%>">

<input type="hidden" name="Skey" value="<%=Skey%>">
<input type="hidden" name="payment_method" value="<%=payment_method%>">
</form>

</body>
</html>

<script language="JavaScript">

	document.fData.submit();

</script>

