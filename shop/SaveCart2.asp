<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
'�T�E���h�n�E�X
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->

<%
'========================================================================
'
'	�J�[�g���e�̕ۑ�2
'
'�X�V����
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userID

Dim CartName

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- UserID ���o��
userID = Session("userID")

'---- �Ăяo��������̃f�[�^���o��
CartName = Left(ReplaceInput(Request("CartName")), 20)

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SaveCart2.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if wMSG = "" then
	Response.Redirect "SaveCartList.asp"
else
	Response.Redirect "SaveCart.asp?msg=" & wMSG
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()
Dim i

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

Dim RSv

'----���󒍃f�[�^���o��
wSQL = ""
wSQL = wSQL & "SELECT a.�󒍖��הԍ�"
wSQL = wSQL & "     , a.���[�J�[�R�[�h"
wSQL = wSQL & "     , a.���i�R�[�h"
wSQL = wSQL & "     , a.�F"
wSQL = wSQL & "     , a.�K�i"
wSQL = wSQL & "     , a.�󒍐���"
wSQL = wSQL & "  FROM ���󒍖��� a WITH (NOLOCK)"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wMSG = "�ۑ�����J�[�g��񂪂���܂���B"
	exit function
end if

'---- �ۑ��J�[�g���o�^
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM �ۑ��J�[�g"
wSQL = wSQL & " WHERE �ڋq�ԍ� = " & userID
wSQL = wSQL & "   AND �J�[�g�� = '" & CartName & "'"
	  
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = false then
	RSv.Delete
end if

RSv.AddNew

RSv("�ڋq�ԍ�") = userID
RSv("�J�[�g��") = CartName
RSv("�o�^��") = now()

RSv.Update
RSv.close


Do Until RS.EOF = true

	'---- �ۑ��J�[�g���׏��o�^
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM �ۑ��J�[�g����"
	wSQL = wSQL & " WHERE 1 = 2"
		  
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	'---- insert �J�^���O����
	RSv.AddNew

	RSv("�ڋq�ԍ�") = userID
	RSv("�J�[�g��") = CartName
	RSv("�󒍖��הԍ�") = RS("�󒍖��הԍ�")
	RSv("���[�J�[�R�[�h") = RS("���[�J�[�R�[�h")
	RSv("���i�R�[�h") = RS("���i�R�[�h")
	RSv("�F") = RS("�F")
	RSv("�K�i") = RS("�K�i")
	RSv("�󒍐���") = RS("�󒍐���")

	RSv.Update
	
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
