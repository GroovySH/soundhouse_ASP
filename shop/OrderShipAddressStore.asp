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
<!--#include file="../common/HttpsSecurity.inc"-->

<%
'========================================================================
'
'	�I�[�_�[�͐���o�^
'		���͂��ꂽ�f�[�^�[�̃`�F�b�N�B
'		OK�Ȃ���͂��ꂽ�͐����Web�ڋq�Z���AWeb�ڋq�Z���d�b�ԍ��֒ǉ��B
'
'�ύX����
'2011/01/31 GV(ay) �V�K�쐬
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache

'---- Session���
Dim wUserID
Dim wUserName
Dim wMsg

Dim wErrMsg
Dim wErrDesc   '2011/08/01 an add

'---- �󂯓n����������ϐ�
Dim ship_name
Dim ship_zip
Dim ship_prefecture
Dim ship_address
dim ship_telephone

'---- DB
Dim Connection

'=======================================================================
'	�󂯓n�������o��
'=======================================================================
'---- Session�ϐ�
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session.contents("msg")

'---- �󂯓n�������o��
ship_name = Left(ReplaceInput(Trim(Request("ship_name"))), 30)
ship_zip = Left(ReplaceInput(Trim(Request("ship_zip"))), 10)
ship_prefecture = Left(ReplaceInput(Trim(Request("ship_prefecture"))), 4)
ship_address = Left(ReplaceInput(Trim(Request("ship_address"))), 40)
ship_telephone = Left(ReplaceInput(Trim(Request("ship_telephone"))), 20)

'---- �Z�b�V�����؂�`�F�b�N
If wUserID = ""Then
	Response.Redirect g_HTTP
End If

Session("msg") = ""
wErrMsg = ""

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()
Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderShipAddressStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- �G���[�������Ƃ��͒������e�m�F�y�[�W�A�G���[������Β������e�w��y�[�W��
If wErrMsg = "" Then
	Server.Transfer "OrderInfoEnter.asp"
Else
	Session("msg") = wErrMsg
	Server.Transfer "OrderShipAddress.asp"
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

Dim vAddNo

''---- ���̓f�[�^�[�̃`�F�b�N
Call validate_data()

If wErrMsg = "" Then
	'---- Web�ڋq�Z�����o�^
	vAddNo = insert_todokesaki()

	'---- ���󒍏��o�^
	Call insert_Order(vAddNo)

End If

End Function

'========================================================================
'
'	Function	�͐���̓o�^
'
'========================================================================
Function insert_todokesaki()

Dim vSQL
Dim RSv
Dim i
Dim vMaxNo

'---- MAX�Z���A�Ԃ̎��o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    MAX(�Z���A��) AS MAX�Z���A��"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web�ڋq�Z�� WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    �ڋq�ԍ� = " & wUserID

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

vMaxNo = RSv("MAX�Z���A��") + 1

RSv.Close

'---- insert �ڋq�Z��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web�ڋq�Z��"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    1 = 2"
 
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

RSv.AddNew

RSv("�ڋq�ԍ�") = wUserID
RSv("�Z���A��") = vMaxNo
RSv("�Z���敪") = "�͐�"
RSv("�Z������") = ship_name
RSv("�ڋq�X�֔ԍ�") = ship_zip
RSv("�ڋq�s���{��") = ship_prefecture
RSv("�ڋq�Z��") = ship_address
RSv("�Ζ���t���O") = "N"
RSv("�K��͐�t���O") = "N"
RSv("�ŏI�X�V��") = Now()
RSv("�ŏI�X�V�҃R�[�h") = "Internet"

RSv.Update
RSv.Close

'---- insert �ڋq�Z���d�b�ԍ�
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web�ڋq�Z���d�b�ԍ�"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    1 = 2"
 
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

RSv.AddNew

RSv("�ڋq�ԍ�") = wUserID
RSv("�Z���A��") = vMaxNo
RSv("�d�b�A��") = 1
RSv("�d�b�敪") = "�d�b"
RSv("�ڋq�d�b�ԍ�") = ship_telephone
RSv("�����p�ڋq�d�b�ԍ�") = cf_numeric_only(ship_telephone)
RSv("�ŏI�X�V��") = Now()
RSv("�ŏI�X�V�҃R�[�h") = "Internet"

RSv.Update
RSv.Close

insert_todokesaki = vMaxNo

End function

'========================================================================
'
'	Function	���󒍏��̓o�^
'
'========================================================================
Function insert_Order(vAddNo)

Dim RSv
Dim vSQL

'----���󒍃f�[�^���o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ����"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

RSv("�͐�敪") = "D"
RSv("�͐�Z���A��") = vAddNo
RSv("�͐於�O") = ship_name
RSv("�͐�X�֔ԍ�") = ship_zip
RSv("�͐�s���{��") = ship_prefecture
RSv("�͐�Z��") = ship_address
RSv("�͐�d�b�ԍ�") = ship_telephone

RSv.Update
RSv.Close

End Function

'========================================================================
'
'	Function	���̓f�[�^�[�̃`�F�b�N
'
'========================================================================
Function validate_data()

Dim vSQL
Dim RSv

Dim vTel
Dim vAddress
Dim vCnt
Dim vBanchFl
Const cNumber = "0123456789�O�P�Q�R�S�T�U�V�W�X���O�l�ܘZ������\"

If ship_name = "" Then
	wErrMsg = wErrMsg & "���͂���̂����O����͂��Ă��������B<br>"
Else

	If Len(ship_name) > 30 Then
		wErrMsg = wErrMsg & "���͂���̂����O��30�����ȓ��œ��͂��Ă��������B<br>"
	End If

End If

If ship_zip = "" Then
	wErrMsg = wErrMsg & "���͂���̗X�֔ԍ�����͂��Ă��������B<br>"
Else
	If IsNumeric(Replace(ship_zip, "-", "")) = False Or cf_checkHankaku2(ship_zip) = False Then
		wErrMsg = wErrMsg & "���͂���̗X�֔ԍ��𔼊p�œ��͂��Ă��������B<br>"
	Else
		If Len(ship_zip) > 10 Then
			wErrMsg = wErrMsg & "���͂���̗X�֔ԍ���10�����ȓ��œ��͂��Ă��������B<br>"
		Else
			If check_zip(ship_zip, vAddress) = False Then
				wErrMsg = wErrMsg & "���͂���̗X�֔ԍ����X�֔ԍ������ɂ���܂���B<br>"
			Else
				If InStr(vAddress, Trim(ship_prefecture)) <= 0  Then
					wErrMsg = wErrMsg & "���͂��ꂽ�X�֔ԍ��Ɠs���{������v���܂���B<br>"
				End If
			End If
		End If
	End If
End If

If ship_prefecture = "" Then
	wErrMsg = wErrMsg & "���͂���̓s���{����I�����Ă��������B<br>"
Else

	If Len(ship_prefecture) > 4 Then
		wErrMsg = wErrMsg & "���͂���̓s���{����4�����ȓ��œ��͂��Ă��������B<br>"
	End If

End If

If ship_address = "" Then
	wErrMsg = wErrMsg & "���͂���̏Z������͂��Ă��������B<br>"
Else

	If Len(ship_address) > 40 Then
		wErrMsg = wErrMsg & "���͂���̏Z����40�����ȓ��œ��͂��Ă��������B<br>"
	End If

	If Len(ship_address) > 0 Then

		vBanchFl = False

		For vCnt = 1 To Len(cNumber)

			If InStr(ship_address, Mid(cNumber, vCnt, 1)) > 0 Then
				vBanchFl = True
				Exit For
			End If

		Next

		If vBanchFl = False Then
			wErrMsg= wErrMsg & "�Ԓn����͂��Ă��������B<br>"
		End If

	End If

End If

If ship_telephone = "" Then
	wErrMsg= wErrMsg & "���͂���̓d�b�ԍ�����͂��Ă��������B<br>"
Else

	If IsNumeric(Replace(ship_telephone, "-", "")) = False Or cf_checkHankaku2(ship_telephone) = False Then
		wErrMsg = wErrMsg & "���͂���̓d�b�ԍ��͔��p�����ƃn�C�t��(�|)�œ��͂��Ă��������B<br>"
	Else

		If Len(ship_telephone) > 20 Then
			wErrMsg = wErrMsg & "���͂���̓d�b�ԍ���20�����ȓ��œ��͂��Ă��������B<br>"
		Else
			vTel = Replace(ship_telephone, "-", "")

			If Len(vTel) = 10 Or Len(vTel) = 11 Then
			Else
				wErrMsg = wErrMsg & "���͂��ꂽ�d�b�ԍ������m�F���������B<br>"
			End If

		End If

	End If

End If

'---- ����Z�������邩�ǂ����`�F�b�N
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    �Z���A��"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web�ڋq�Z�� WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    �ڋq�ԍ� = " & wUserID
vSQL = vSQL & "    AND �Z������ = '" & ship_name & "'"
vSQL = vSQL & "    AND �ڋq�X�֔ԍ� = '" & ship_zip & "'"
vSQL = vSQL & "    AND �ڋq�s���{�� = '" & ship_prefecture & "'"
vSQL = vSQL & "    AND �ڋq�Z�� = '" & ship_address & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = False Then				'����Z������
	wErrMsg = wErrMsg & "����Z�������ɓo�^����Ă��܂��B<br>"
	Exit Function
End If

RSv.Close

'If wErrMsg <> "" Then
'	wErrMsg = "<b>�ȉ��̓��̓G���[��������ĉ������B</b><br /><br />" & wErrMsg
'End If

End Function

'========================================================================
'
'	Function	�X�֔ԍ���������
'
'========================================================================
Function check_zip(pZip, pAddress)

Dim vSQL
Dim RSv

'---- �X�֔ԍ���������
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    �s���{��������"
vSQL = vSQL & "  , �s�撬��������"
vSQL = vSQL & "  , ���於����"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    �X�֔ԍ����� WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    �X�֔ԍ� = '" & Replace(pZip, "-", "") & "'"
	  
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = False Then
	check_zip = True
	pAddress = Trim(RSv("�s���{��������")) & Trim(RSv("�s�撬��������"))
Else
	check_zip = False
	pAddress = ""
End If

RSv.Close

End Function
%>
