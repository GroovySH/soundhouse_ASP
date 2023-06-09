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
'	�I�[�_�[���[�����o�^
'		���͂��ꂽ�f�[�^�[�̃`�F�b�N�B
'		OK�Ȃ���͂��ꂽ���[���������󒍂֒ǉ��B
'
'�ύX����
'2011/01/28 GV(ay) �V�K�쐬
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache

'---- Session���
Dim wUserID
Dim wUserName
Dim wMsg

Dim wErrMsg

'---- �󂯓n����������ϐ�
Dim loan_downpayment_fl
Dim loan_downpayment_am
Dim loan_term
Dim loan_am
Dim loan_term_payment
Dim loan_apply_fl
Dim loan_company
Dim wErrDesc   '2011/08/01 an add

'---- DB
Dim Connection

'=======================================================================
'	�󂯓n�������o��
'=======================================================================
'---- Session�ϐ�
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session("msg")

'---- �󂯓n�������o��
loan_downpayment_fl = Left(ReplaceInput(Trim(Request("loan_downpayment_fl"))), 1)
loan_downpayment_am = ReplaceInput(Trim(Request("loan_downpayment_am")))
loan_apply_fl = Left(ReplaceInput(Trim(Request("loan_apply_fl"))), 1)
loan_company = Left(ReplaceInput(Trim(Request("loan_company"))), 10)
loan_term_payment = Left(ReplaceInput(Trim(Request("loan_term_payment"))), 1)
loan_term = ReplaceInput(Trim(Request("loan_term")))
loan_am = ReplaceInput(Trim(Request("loan_am")))

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
	wErrDesc = "OrderLoanStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- �G���[�������Ƃ��͒������e�m�F�y�[�W�A�G���[������Β������e�w��y�[�W��
If wErrMsg = "" Then
	Server.Transfer "OrderConfirm.asp"
Else
	Session("msg") = wErrMsg
	Server.Transfer "OrderLoan.asp"
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

'---- ���󒍏��X�V
Call update_order_header()

''---- ���̓f�[�^�[�̃`�F�b�N
Call validate_data()

End Function

'========================================================================
'
'	Function	���󒍏��̍X�V
'
'========================================================================
Function update_order_header()

Dim RSv
Dim vSQL

'---- ����Recordset���o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ����"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

RSv("���[����������t���O") = loan_downpayment_fl
If loan_downpayment_fl = "Y" Then
	If IsNumeric(loan_downpayment_am) = False Then
		RSv("���[������") = 0
	Else
		RSv("���[������") = CCur(loan_downpayment_am)
	End If
Else
	RSv("���[������") = 0
End If

RSv("�I�����C�����[���\���t���O") = loan_apply_fl

Select Case loan_apply_fl
	Case "Y"
		RSv("���[�����") = loan_company
		RSv("��]���[����") = 0
		RSv("���[�����z") = 0

	Case "N"
		RSv("���[�����") = ""
		Select Case loan_term_payment
			Case "T"		' ��]���[���񐔂̏ꍇ
				RSv("��]���[����") = CLng(loan_term)
				RSv("���[�����z") = 0

			Case "P"		' ���z�x�����z�̏ꍇ
				RSv("��]���[����") = 0
				If IsNumeric(loan_am) = False Then
					RSv("���[�����z") = 0
				Else
					RSv("���[�����z") = CCur(loan_am)
				End If
			Case Else
				RSv("��]���[����") = 0
				RSv("���[�����z") = 0

		End Select

End Select

RSv("�ŏI�X�V��") = Now()

RSv.Update
RSv.Close

End Function

'========================================================================
'
'	Function	���̓f�[�^�[�̃`�F�b�N
'
'========================================================================
Function validate_data()

If isNumeric(loan_term) = False Then
	loan_term = 0
End If

' ��������^�Ȃ�
If loan_downpayment_fl = "" Then
	wErrMsg = wErrMsg & "���[����������/�Ȃ���I�����Ă��������B<br>"
End If

' ��������̏ꍇ���[�������̃`�F�b�N
If loan_downpayment_fl = "Y" Then
	If loan_downpayment_am = "" Then
		wErrMsg = wErrMsg & "���[����������͂��Ă��������B<br>"
	Else
		If isNumeric(loan_downpayment_am) = False Then
			wErrMsg = wErrMsg & "���[�������𐔎��݂̂œ��͂��Ă��������B<br>"
			loan_downpayment_am = 0
		Else
			If loan_downpayment_am = 0 Then
				wErrMsg = wErrMsg & "���[����������͂��Ă��������B<br>"
			End If
		End If
	End If
End If

' �I�����C�����[��
If loan_apply_fl = "" Then
	wErrMsg = wErrMsg & "�I�����C���Ń��[����\�����ނ��ǂ�����I�����Ă��������B<br>"
End If

' �I�����C�����[����\���ޏꍇ���[����Ђ̃`�F�b�N
If loan_apply_fl = "Y" Then

	' ���g�p���ڂ��N���A
	loan_am = 0
	loan_term = 0

	Select Case loan_company
		Case ""
			wErrMsg = wErrMsg & "���[����Ђ�I�����Ă��������B<br>"

		Case "�W���b�N�X"
			If loan_downpayment_fl = "Y" Then
				wErrMsg = wErrMsg & "�W���b�N�X�ł̂��\�����݂̏ꍇ�A�������w�肷�邱�Ƃ͂ł��܂���B<br>�����Ȃ������I�����������B<br>"
			End If

	End Select

Else

	Select Case loan_term_payment
		Case ""
			wErrMsg = wErrMsg & "���[���񐔂����z�x���z��I�����Ă��������B<br>"

		Case "T"		' ��]���[����
			' ���g�p���ڂ��N���A
			loan_am = 0

			If loan_term = 0 Then
				wErrMsg = wErrMsg & "��]���[���񐔂�I�����Ă��������B<br>"
				loan_term = 0
			End If

		Case "P"		' ���z�x�����z
			' ���g�p���ڂ��N���A
			loan_term = 0

			If loan_am = "" Then
				wErrMsg = wErrMsg & "���z�x�����z����͂��Ă��������B<br>"
			Else
				If IsNumeric(loan_am) = False Then
					wErrMsg = wErrMsg & "���z�x�����z�𐔎��݂̂œ��͂��Ă��������B<br>"
					loan_am = 0
				End If
				If loan_am = 0 Then
					wErrMsg = wErrMsg & "���z�x�����z����͂��Ă��������B<br>"
				End If
			End If

	End Select

End If

'If wErrMsg <> "" Then
'	wErrMsg = "<b>�ȉ��̓��̓G���[��������ĉ������B</b><br /><br />" & wErrMsg
'End If

End Function
%>
