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
<!--#include file="../3rdParty/aspJSON1.17.asp"-->
<%
'========================================================================
'
'	�w�������ꗗ�y�[�W��ɔ�\��/�\���̃t���O���Z�b�g����B
'
'
'�ύX����
'2016.02.12 GV �V�K�쐬
'
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wFlg						' ���s�t���O
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim wDetailNo					' ���הԍ�
Dim wMode						' ���[�h(Y ... �ǉ��AN ... �폜)
Dim oJSON						' JSON�I�u�W�F�N�g


' �����ݒ�
wFlg = True

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
' Get�p�����[�^
' �ڋq�ԍ�
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
	wFlg = False
End If

' �����ԍ�
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wFlg = False
End If

'���הԍ�
wDetailNo = ReplaceInput_NoCRLF(Trim(Request("dno")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wDetailNo) = False) Or (cf_checkNumeric(wDetailNo) = False) Then
	wFlg = False
End If

' ���[�h
wMode = ReplaceInput_NoCRLF(Trim(Request("mode")))
wMode = UCase(wMode) ' �啶����
' �`�F�b�N
If cf_checkHankaku2(wMode) = False Then
	wFlg = False
Else
	If (wMode = "Y") Or (wMode = "N") Then
	Else
		wFlg = False
	End If
End If

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   ' member�n�̑��̃y�[�W�����ɂȂ炤
If Err.Description <> "" Then
'	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
'	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If

Call close_db()

If Err.Description <> "" Then

End If


'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

ConnectionEmax.close
Set ConnectionEmax= Nothing

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

Dim vSQL
Dim i
Dim j
Dim vRS1
Dim vRS2
Dim vOrderNo
Dim vDetailNo
Dim vCustomerNo

Set oJSON = New aspJSON


' ���͒l������̏ꍇ
If (wFlg = True) Then
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "	 T1.�ڋq�ԍ� "
	vSQL = vSQL & " ,T1.�󒍔ԍ� "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "  �� T1 WITH (NOLOCK)"
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "  T1.�ڋq�ԍ� = " & wCustomerNo
	vSQL = vSQL & " AND T1.�󒍔ԍ� = " & wOrderNo

	Set vRS1 = Server.CreateObject("ADODB.Recordset")
	vRS1.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂���ꍇ
	If vRS1.EOF = False Then
		vSQL = ""
		vSQL = vSQL & "SELECT "
		vSQL = vSQL & "	 T1.�ڋq�ԍ� "
		vSQL = vSQL & " ,T1.�󒍔ԍ� "
		vSQL = vSQL & " ,T1.�󒍖��הԍ� "
		vSQL = vSQL & " ,T1.��\���t���O "
		vSQL = vSQL & "FROM "
		vSQL = vSQL & "  �󒍔�\�����X�g T1 "
		vSQL = vSQL & "WHERE "
		vSQL = vSQL & "  T1.�ڋq�ԍ� = " & wCustomerNo
		vSQL = vSQL & " AND T1.�󒍔ԍ� = " & wOrderNo
		vSQL = vSQL & " AND T1.�󒍖��הԍ� = " & wDetailNo

		Set vRS2 = Server.CreateObject("ADODB.Recordset")
		vRS2.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

		'���R�[�h�����݂���ꍇ
		If vRS2.EOF = False Then
			' ���[�h����\�����X�g�ǉ��̏ꍇ
			If (wMode = "Y") Then
				'�������Ȃ�
			ElseIf (wMode = "N") Then
				' ���[�h����\�����X�g����폜�̏ꍇ�A�폜
				vRS2.Delete
			End If
		Else
			' ���R�[�h�����݂��Ȃ������ꍇ
			' ���[�h����\�����X�g�ǉ��̏ꍇ
			If (wMode = "Y") Then
				vRS2.AddNew
				vRS2("�ڋq�ԍ�") = wCustomerNo
				vRS2("�󒍔ԍ�") = wOrderNo
				vRS2("�󒍖��הԍ�") = wDetailNo
				vRS2.Update
			End If
		End If

		'���R�[�h�Z�b�g�����
		vRS2.Close

		'���R�[�h�Z�b�g�̃N���A
		Set vRS2 = Nothing
	Else
		'�󒍂����݂��Ȃ������ꍇ
		wFlg = false
	End If

	'���R�[�h�Z�b�g�����
	vRS1.Close

	'���R�[�h�Z�b�g�̃N���A
	Set vRS1 = Nothing
End If
	' ����
	oJSON.data.Add "result" ,wFlg

	'�ڋq�ԍ�
	If (IsNull(wCustomerNo)) Then
		vCustomerNo = ""
	Else
		vCustomerNo = CStr(Trim(wCustomerNo))
	End If

	oJSON.data.Add "cno" ,vCustomerNo


	'�󒍔ԍ�
	If (IsNull(wOrderNo)) Then
		vOrderNo = ""
	Else
		vOrderNo = CStr(Trim(wOrderNo))
	End If

	'�󒍖��הԍ�
	If (IsNull(wDetailNo)) Then
		vDetailNo = ""
	Else
		vDetailNo = CStr(Trim(wDetailNo))
	End If

	oJSON.data.Add "ono" ,vOrderNo
	oJSON.data.Add "dno" ,vDetailNo

	'���[�h
	oJSON.data.Add "mode" ,wMode

' -------------------------------------------------
' JSON�f�[�^�̕ԋp
' -------------------------------------------------
' �w�b�_�o��
Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "X-Content-Type-Options", "nosniff"

' JSON�f�[�^�̏o��
Response.Write oJSON.JSONoutput()

End Function

'========================================================================
%>
