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
'	�w������
'	�L�����Z���������i�Ƀt���O���Z�b�g����B
'
' �ύX�\���
'   Web�����i�C���^�[�l�b�g�A�X�}�[�g�t�H���j�ł���B
'   �x�������@���u���[���v�ȊO�ł���B
'   �󒍃X�e�[�^�X���u�󒍁v(�o�׎w������)�łȂ��B
'   ���[�J�[�����i���܂܂�Ă��Ȃ��B
'   ���c���K���݌ɐ���=0�łȂ��B
'
'
'�ύX����
'2016/02/04 GV �V�K�쐬�B(�����ύX�L�����Z���@�\)
'2020.10.20 GV �����ύX�L�����Z������Emax�w�������f�[�^�C���B(#2572)
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim wOrderDetailNo				' �󒍔ԍ�
Dim wSetFlg						' �ύX���[�h(Y/N)
Dim wFlg						' ���s�t���O
Dim oJSON						' JSON�I�u�W�F�N�g
Dim modifyFlag					' �ύX�\�t���O
Dim wNgReason					' �s���R
Dim wDepositFlag   				' ���������t���O
Dim wDepositAmount 				' �������v���z
Dim isUpdateOrderTable			' �󒍃e�[�u�����X�V����(Y/N) ' 2020.10.20 GV add

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


'�t���O
wSetFlg = ReplaceInput_NoCRLF(Trim(Request("f")))
wSetFlg = UCase(wSetFlg)
If (wSetFlg <> "Y") And (wSetFlg <> "N") And (wSetFlg <> "") Then
	wFlg = False
End If


' �󒍖��הԍ�(�J���}��؂�)
wOrderDetailNo = ReplaceInput_NoCRLF(Trim(Request("od")))
If (wOrderDetailNo = "") Then
	wFlg = False
Else
	Dim pos
	' �A���_�[�X�R�A�̈ʒu���擾
	pos = InStr(wOrderDetailNo, "_")

	'�����񂪂���A�A���_�[�X�R�A���܂܂�Ă���ꍇ
	If (Len(wOrderDetailNo) > 0) And (pos > 0) Then
		wOrderDetailNo = Replace(wOrderDetailNo, "_", ",")
	End If
End If

'2020.10.20 GV add
'�󒍃e�[�u���X�V�t���O
isUpdateOrderTable = ReplaceInput_NoCRLF(Trim(Request("up_od")))
If (isUpdateOrderTable <> "") Then
	isUpdateOrderTable = UCase(isUpdateOrderTable)

	If (isUpdateOrderTable <> "Y") And (isUpdateOrderTable <> "N") Then
		wFlg = False
	End If
End If


wNgReason = ""
wDepositFlag = ""
wDepositAmount = 0


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
'========================================================================
Function main()

Dim vSQL
Dim vRS1
Dim vRS2
Dim okFlag
Dim wSQL
Dim orderDate
Dim deleteDate
Dim promote
Dim i

' JSON�I�u�W�F�N�g����
Set oJSON = New aspJSON

okFlag = True

' ���͒l������̏ꍇ
If (wFlg = True) Then
	'---- �g�����U�N�V�����J�n
	ConnectionEmax.BeginTrans

	'�󒍂̎��o��
	wSQL = ""
	wSQL = wSQL & "SELECT "
	wSQL = wSQL & "  �󒍌`�� "
	wSQL = wSQL & " ,�x�����@ "
	wSQL = wSQL & " ,�󒍓� "
	wSQL = wSQL & " ,�폜�� "
	wSQL = wSQL & " ,�폜�� "
	wSQL = wSQL & " ,Web�����ύX�L�����Z�����t���O "
	wSQL = wSQL & " ,�������v���z "
	wSQL = wSQL & " ,���������t���O "
'	wSQL = wSQL & "  FROM �� WITH(NOLOCK) " ' 2020.10.20 GV mod
	wSQL = wSQL & "  FROM �� " ' 2020.10.20 GV add

	'2020.10.20 GV add start
	' �󒍃f�[�^�X�V�t���O��Y�łȂ�
	If (isUpdateOrderTable <> "Y") Then
		wSQL = wSQL & " WITH (NOLOCK) "
	End If
	'2020.10.20 GV add end

	wSQL = wSQL & " WHERE �󒍔ԍ� = " & wOrderNo
	wSQL = wSQL & "  AND �ڋq�ԍ� = " & wCustomerNo
	'Response.Write wSQL & "<br>"

	Set vRS1 = Server.CreateObject("ADODB.Recordset")
	vRS1.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂��Ă���ꍇ
	If vRS1.EOF = False Then
		If (okFlag = True) Then
			'2020.10.20 GV add start
			' �󒍃f�[�^�X�V�t���O��Y�ł���
			If (isUpdateOrderTable = "Y") Then
				vRS1("Web�����ύX�L�����Z�����t���O") = wSetFlg
				vRS1.update
			End If
			'2020.10.20 GV add end

			wSQL = ""
			wSQL = wSQL & "SELECT "
			wSQL = wSQL & "  �󒍖��הԍ� "
			wSQL = wSQL & " ,Web�L�����Z���t���O "
			wSQL = wSQL & "FROM "
			wSQL = wSQL & "  �󒍖��� "
			wSQL = wSQL & "WHERE "
			wSQL = wSQL & "   �󒍔ԍ� = " & wOrderNo
			wSQL = wSQL & "   AND �󒍖��הԍ� IN (" & wOrderDetailNo & ") "

			'@@@@Response.Write wSQL & "<br>"

			Set vRS2 = Server.CreateObject("ADODB.Recordset")
			vRS2.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
	
			'���R�[�h�����݂��Ă���ꍇ
			If vRS2.EOF = False Then
				For i = 0 To (vRS2.RecordCount - 1)
					vRS2("Web�L�����Z���t���O") = wSetFlg
					vRS2.update
					okFlag = True

					' ���̃��R�[�h�s�ֈړ�
					vRS2.MoveNext

					If vRS2.EOF Then
						Exit For
					End If
				Next
			Else
				okFlag = False
				wNgReason = "6"
			End If
		End If
	Else
		'���R�[�h���Ȃ��ꍇ�ANG
		okFlag = False
		wNgReason = "7"
	End If

	If okFlag = True Then
		'�R�~�b�g
		ConnectionEmax.CommitTrans

		'���R�[�h�Z�b�g�����
		vRS2.Close

		'���R�[�h�Z�b�g�̃N���A
		Set vRS2 = Nothing
	Else
		'���[���o�b�N
		ConnectionEmax.RollbackTrans
	End If

	'���R�[�h�Z�b�g�����
	vRS1.Close

	'���R�[�h�Z�b�g�̃N���A
	Set vRS1 = Nothing
Else
	'���͒l��NG�̏ꍇ
	okFlag = False
	wNgReason = "99"
End If

wFlg = okFlag

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
	oJSON.data.Add "od_no" ,wOrderDetailNo
	oJSON.data.Add "set_flg" ,wSetFlg
	oJSON.data.Add "reason" ,wNgReason
	oJSON.data.Add "is_up_od_tbl" ,isUpdateOrderTable

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
