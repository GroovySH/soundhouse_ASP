<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
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
'	�J�[�h�o�^
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

'---- �Z�L�����e�B�[�L�[�Z�b�g 
payment_method = ReplaceInput(Request("payment_method"))
if payment_method <> "�N���W�b�g�J�[�h" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

Skey = ReplaceInput(Request("Skey"))

'---- UserID ���o��
userID = Session("userID")
userName = Session("userName")
w_sessionID = Session.SessionID

'---- ���̓f�[�^�[�̎��o��
CardCompany = ReplaceInput(Trim(Request("CardCompany")))
CardNo = ReplaceInput(Trim(Request("CardNo")))
CardExpMM = ReplaceInput(Trim(Request("CardExpMM")))
CardExpYY = ReplaceInput(Trim(Request("CardExpYY")))
CardName = ReplaceInput(Trim(Request("CardName")))
CardHoji = ReplaceInput(Trim(Request("CardHoji")))

'---- ���C������
call connect_db()
call main()
call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp" & Err.Description
end if

Session("msg") = ""
'---- �G���[�������Ƃ���OrderProcessing�A�G���[������΃J�[�h���̓y�[�W��
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

'---- ���͏��̃`�F�b�N
Call ValidateData()

if wMSG <> "" then
	exit function
end if

'---- �J�[�h���X�V
wSQL = ""
wSQL = wSQL & "SELECT �J�[�h���"
wSQL = wSQL & "     , �J�[�h�ԍ�"
wSQL = wSQL & "     , �J�[�h�L������"
wSQL = wSQL & "     , �J�[�h���`�l"
wSQL = wSQL & "  FROM Web�ڋq"
wSQL = wSQL & " WHERE �ڋq�ԍ� = " & UserID
  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = true then
	wMSG = "�������ُ�I�����܂����B"
	exit function
end if

vOldCardNo = RS("�J�[�h�ԍ�")
if IsNull(vOldCardNo) = true then
	vOldCardNo = ""
end if

if vOldCardNo <> CardNo then
	if isNumeric(CardNo) = false then
		wMSG = "�J�[�h�ԍ��͐����݂̂œ��͊肢�܂��B"
		exit function
	end if
end if

if CardHoji = "Y" then
	RS("�J�[�h���") = CardCompany
	RS("�J�[�h�ԍ�") = "************" & Right(CardNo, 4)
	RS("�J�[�h�L������") = CardExpMM & "/" & CardExpYY
	RS("�J�[�h���`�l") = CardName
else
	RS("�J�[�h���") = ""
	RS("�J�[�h�ԍ�") = ""
	RS("�J�[�h�L������") = ""
	RS("�J�[�h���`�l") = ""
end if

RS.update
RS.close

'---- �J�[�h���X�V2(�J�[�h�ԍ��o�^�ύX���j
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
'	Function	���̓f�[�^�[�̃`�F�b�N
'
'========================================================================
'
Function ValidateData()

wMSG = ""
'---- �J�[�h���
if CardCompany = "" then
	wMSG = wMSG & "�J�[�h��Ђ�I���肢�܂��B<br>"
end if

'---- �J�[�h�ԍ�
if CardNo = "" then
	wMSG = wMSG & "�J�[�h�ԍ�����͊肢�܂��B<br>"
end if

'---- �J�[�h�L������
if CardExpMM = "" OR CardExpYY = "" then
	wMSG = wMSG & "�J�[�h�L��������I���肢�܂��B<br>"
end if

'---- �J�[�h���`
if CardName = "" then
	wMSG = wMSG & "�J�[�h���`����͊肢�܂��B<br>"
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
<title>�T�E���h�n�E�X  ��������t�@�J�[�h</title>

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

