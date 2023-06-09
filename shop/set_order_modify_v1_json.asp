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
'	�w�������ꗗ�y�[�W
'	Emax��[��].[Web�����ύX�L�����Z�����t���O]��"Y"�܂���"N"�ɍX�V����
'
'�ύX����
'2016/02/04 GV �V�K�쐬�B(�����ύX�L�����Z���@�\)
'2020.12.07 GV �󒍖��ׂ�Web�L�����Z���t���O�X�V���C�B(#2619)
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim wDefer						' �ύX���[�h(Y/N)
Dim wFlg						' ���s�t���O
Dim oJSON						' JSON�I�u�W�F�N�g

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
wFlg = True

' Get�p�����[�^
' �ڋq�ԍ�
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
	wFlg = False
End If

' �󒍔ԍ�
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wFlg = False
End If

'�ۗ����[�h
wDefer = ReplaceInput_NoCRLF(Trim(Request("defer")))
wDefer = UCase(wDefer)
If wFlg = True Then
	Select Case wDefer
		Case "Y"
			wFlg = True
		Case "N"
			wFlg = True
		Case Else
			wFlg = False
	End Select
End If

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   ' member�n�̑��̃y�[�W�����ɂȂ炤
If Err.Description <> "" Then
End If

Call close_db()

Call sendResponse()

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
'�ύX����
'2020.12.07 GV �󒍖��ׂ�Web�L�����Z���t���O�X�V���C�B(#2619)
'========================================================================
Function main()

Dim vSQL
Dim vRS
Dim vRS2 '2020.12.07 GV add
Dim i '2020.12.07 GV add
Set oJSON = New aspJSON


' ���͒l������̏ꍇ
If (wFlg = True) Then
	'---- �g�����U�N�V�����J�n
	ConnectionEmax.BeginTrans

	'�󒍂̎��o��
	vSQL = ""
	vSQL = vSQL & "SELECT a.* "
	'vSQL = vSQL & "  FROM �� a WITH(UPDLOCK) "
	vSQL = vSQL & "  FROM �� a "
	vSQL = vSQL & " WHERE �󒍔ԍ� = " & wOrderNo
	vSQL = vSQL & "  AND �ڋq�ԍ� = " & wCustomerNo
	vSQL = vSQL & "  AND �폜�� IS NULL "
	'@@@@Response.Write vSQL & "<br>"

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂��Ă���ꍇ
	If vRS.EOF = False Then
		'vSQL = ""
		'vSQL = vSQL & "UPDATE �� "
		'vSQL = vSQL & " SET "
		'vSQL = vSQL & " Web�����ύX�L�����Z�����t���O = " & wDefer
		'vSQL = vSQL & " WHERE �󒍔ԍ� = " & wOrderNo
		'vSQL = vSQL & "  AND �ڋq�ԍ� = " & wCustomerNo
		'vSQL = vSQL & "  AND �폜�� = IS NULL "
		vRS("Web�����ύX�L�����Z�����t���O") = wDefer
		vRS.update

		wFlg = True
	Else
		wFlg = False
	End If

	'2020.12.07 GV add start
	If (wDefer = "N") Then
		'�󒍖��ׂ̎��o��
		vSQL = ""
		vSQL = vSQL & "SELECT a.* "
		vSQL = vSQL & "  FROM �󒍖��� a "
		vSQL = vSQL & " WHERE �󒍔ԍ� = " & wOrderNo
		vSQL = vSQL & " AND  Web�L�����Z���t���O = 'Y' "
		vSQL = vSQL & " ORDER BY �󒍖��הԍ� "

		Set vRS2 = Server.CreateObject("ADODB.Recordset")
		vRS2.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

		'���R�[�h�����݂��Ă���ꍇ
		If vRS2.EOF = False Then
			For i = 0 To (vRS2.RecordCount - 1)
				vRS2("Web�L�����Z���t���O") = wDefer
				vRS2.update

				' ���̃��R�[�h�s�ֈړ�
				vRS2.MoveNext

				If vRS2.EOF Then
					Exit For
				End If
			Next
		End If
	End If
	' 2020.12.07 GV add end


	'�����̏ꍇ
	If (wFlg = True) Then
		'�R�~�b�g
		ConnectionEmax.CommitTrans
	Else
		'���[���o�b�N
		ConnectionEmax.RollbackTrans
	End If

	'���R�[�h�Z�b�g�����
	vRS.Close

	'���R�[�h�Z�b�g�̃N���A
	Set vRS = Nothing

	'2020.12.07 GV add start
	'���R�[�h�Z�b�g�����
	vRS2.Close
	'���R�[�h�Z�b�g�̃N���A
	Set vRS2 = Nothing
	'2020.12.07 GV add end

End If
End Function


'========================================================================
'
'	Function	JSON�ԋp
'
'========================================================================
Function sendResponse()

	' �S������JSON�f�[�^�ɃZ�b�g
	oJSON.data.Add "ono" ,wOrderNo
	oJSON.data.Add "cno" ,wCustomerNo
	oJSON.data.Add "defer" ,wDefer

	If wFlg = True Then
		oJSON.data.Add "result" ,"Y"
	Else
		oJSON.data.Add "result" ,"N"
	End If

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
