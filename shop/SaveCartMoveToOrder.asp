<%@ LANGUAGE="VBScript" %>
<%
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->

<%
'========================================================================
'
'	�ۑ��J�[�g�����󒍂ֈړ�
'
'�X�V����
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userID

Dim CartName

Dim wProdTermFl

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
CartName = ReplaceInput(Request("CartName"))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SaveCartMoveToOrder.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if wMSG = "" then
	Response.Redirect "Order.asp"
else
	Response.Redirect "SaveCartList.asp?msg=" & wMSG
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

'----�ۑ��J�[�g���׃f�[�^���o��
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM �ۑ��J�[�g���� WITH (NOLOCK)"
wSQL = wSQL & " WHERE �ڋq�ԍ� = " & userID
wSQL = wSQL & "   AND �J�[�g�� = '" & CartName & "'"

'@@@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wMSG = "�ۑ����ꂽ�J�[�g��񂪂���܂���B"
	exit function
end if

'---- ���󒍏��o�^
call InsertOrderHeader()

'---- ���󒍏��o�^
call InsertOrderDetail()

RS.close

End function

'========================================================================
'
'	Function	insert InsertOrderHeader
'
'========================================================================
'
Function InsertOrderHeader()
Dim i
Dim RSv

'---- �Y��SessionID�ŉ��󒍂��o�^����Ă邩�`�F�b�N
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM ����"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then

	'---- insert ����
	RSv.AddNew

	RSv("SessionID") = gSessionID		'2011/04/14 hn mod
	RSv("���͓�") = now()
	RSv("�L���R�[�h") = Session("AdID")
	RSv("�ŏI�X�V��") = now()

	RSv.update
end if

RSv.close

End function

'========================================================================
'
'	Function	insert ���󒍖���
'
'========================================================================
'
Function InsertOrderDetail()

Dim vQt
Dim RSv
Dim RSvProduct

'---- ���󒍖��ׂ�����ΑS���폜
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM ���󒍖���"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	  
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

Do until RSv.EOF = true
	RSv.Delete
	RSv.Requery
Loop

Do Until RS.EOF = true
	'---- ���󒍖��דo�^
	vQt = RS("�󒍐���")

	if vQt > 0 then
		'---- ���i����o��
		wSQL = ""
		wSQL = wSQL & "SELECT a.���[�J�[�R�[�h"
		wSQL = wSQL & "     , a.���i�R�[�h"
		wSQL = wSQL & "     , a.���i��"
		wSQL = wSQL & "     , CASE"
		wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
		wSQL = wSQL & "         ELSE a.�̔��P��"
		wSQL = wSQL & "       END AS �̔��P��"
		wSQL = wSQL & "     , a.B�i�P��"	
		wSQL = wSQL & "     , a.�����萔��"	
		wSQL = wSQL & "     , a.������󒍍ϐ���"	
		wSQL = wSQL & "     , a.ASK���i�t���O"
		wSQL = wSQL & "     , a.B�i�t���O"
		wSQL = wSQL & "     , a.�戵���~��"
		wSQL = wSQL & "     , a.�p�ԓ�"
		wSQL = wSQL & "     , a.������"
		wSQL = wSQL & "     , b.���[�J�[��"
		wSQL = wSQL & "     , c.�����\����"
		wSQL = wSQL & "     , c.B�i�����\����"

		wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
		wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
		wSQL = wSQL & "     , Web�F�K�i�ʍ݌� c WITH ( NOLOCK)"

		wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
		wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
		wSQL = wSQL & "   AND c.���i�R�[�h = a.���i�R�[�h"
		wSQL = wSQL & "   AND c.�F = '" & RS("�F") & "'"
		wSQL = wSQL & "   AND c.�K�i = '" & RS("�K�i") & "'"
		wSQL = wSQL & "   AND a.���[�J�[�R�[�h = '" & RS("���[�J�[�R�[�h") & "'"
		wSQL = wSQL & "   AND a.���i�R�[�h = '" & RS("���i�R�[�h") & "'"
		wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
		wSQL = wSQL & "   AND c.�I���� IS NULL"
			  
		Set RSvProduct = Server.CreateObject("ADODB.Recordset")
		RSvProduct.Open wSQL, Connection, adOpenStatic

		if RSvProduct.EOF = false then

			'---- �I���`�F�b�N
			wProdTermFl = "N"
			if isNull(RSvProduct("�戵���~��")) = false then		'�戵���~
				wProdTermFl = "Y"
			end if
			if isNull(RSvProduct("�p�ԓ�")) = false AND RSvProduct("�����\����") <= 0 then		'�p�Ԃō݌ɖ���
				wProdTermFl = "Y"
			end if
			if isNull(RSvProduct("������")) = false then		'�������i
				wProdTermFl = "Y"
			end if

			if RSvProduct("B�i�t���O") <> "Y" then
				if isNull(RSvProduct("�p�ԓ�")) = false AND RSvProduct("�����\����") < vQt then
					vQt = RSvProduct("�����\����")
				end if
			else
				if isNull(RSvProduct("�p�ԓ�")) = false AND RSvProduct("B�i�����\����") < vQt then
					vQt = RSvProduct("B�i�����\����")
				end if
			end if

			if wProdTermFl <> "Y" AND vQt > 0 then
				'---- insert ���󒍖���
				RSv.AddNew
				RSv("SessionID") = gSessionID		'2011/04/14 hn mod
				RSv("�󒍖��הԍ�") = RS("�󒍖��הԍ�")
				RSv("���[�J�[�R�[�h") = RS("���[�J�[�R�[�h")
				RSv("���i�R�[�h") = RS("���i�R�[�h")
				RSv("�F") = RS("�F")
				RSv("�K�i") = RS("�K�i")
				RSv("���[�J�[��") = RSvProduct("���[�J�[��")
				RSv("���i��") = RSvProduct("���i��")

				if RSvProduct("B�i�t���O") <> "Y" then
					RSv("�󒍒P��") = RSvProduct("�̔��P��")
				else
					RSv("�󒍒P��") = RSvProduct("B�i�P��")
				end if

				RSv("�󒍐���") = vQt
				RSv("�󒍋��z") = Fix(RSv("�󒍒P��")) * RSv("�󒍐���")

				if RSvProduct("�����萔��") > RSvProduct("������󒍍ϐ���") then
					RSv("������P���t���O") = "Y"
				else
					RSv("������P���t���O") = ""
				end if

				RSv("B�i�t���O") = RSvProduct("B�i�t���O")

				RSv.Update

			end if
		end if

		RSvProduct.close
	end if

	RS.MoveNext
Loop

RSv.Close

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
