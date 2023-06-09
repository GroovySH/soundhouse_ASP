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
' ���[�U�[�̃|�C���g�����擾
'
'
'�ύX����
'2015/01/26 GV �V�K�쐬
'2016.04.27 GV ���t�C��
'
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wCustomerNo					' �ڋq�ԍ�
Dim oJSON						' JSON�I�u�W�F�N�g


'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
' Get�p�����[�^
wCustomerNo = ReplaceInput(Trim(Request("cno")))

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

Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.close
Set Connection= Nothing

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
Dim vRS
Dim vPointDate
Dim vPointZan

Set oJSON = New aspJSON

'-----------------------------------------------------------
' �l���|�C���g���̎擾
'-----------------------------------------------------------
If (IsNumeric(wCustomerNo)) Then
	vSQL = createPointSql(wCustomerNo)
	'@@@@Response.Write(vSQL&"<br>")

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	// JSON�I�u�W�F�N�g���쐬
	createJsonObject vRS

	'���R�[�h�Z�b�g�����
	vRS.Close

	'���R�[�h�Z�b�g�̃N���A
	Set vRS = Nothing
Else
	' �|�C���g�c
	oJSON.data.Add "point_zan" ,"0"

	' �|�C���g����
	oJSON.data.Add "point_expire_date" ,""
End If


' -------------------------------------------------
' JSON�f�[�^�̕ԋp
' -------------------------------------------------
' �w�b�_�o��
Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
Response.AddHeader "Pragma", "no-cache"
' JSON�f�[�^�̏o��
Response.Write oJSON.JSONoutput()

End Function


'========================================================================
'
'	Function	���t���̃t�H�[�}�b�g (YYYY�NMM��DD��)
'
'========================================================================
Function formatDateYYYYMMDD(pdatDate)

Dim vDate

If IsNull(pdatDate) = True Then
	' Null �͌v�Z�s�\
	Exit Function
End If

If IsDate(pdatDate) = False Then
	' ���t���łȂ���Όv�Z�s�\
	Exit Function
End If

vDate = DatePart("yyyy", pdatDate) & "�N"

If DatePart("m", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("m", pdatDate)
Else
	vDate = vDate & DatePart("m", pdatDate)
End If

vDate = vDate & "��"

If DatePart("d", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("d", pdatDate)
Else
	vDate = vDate & DatePart("d", pdatDate)
End If

vDate = vDate & "��"

formatDateYYYYMMDD = vDate

End Function

'========================================================================
'
'	Function	�|�C���g���̎擾SQL
'
'========================================================================
Function createPointSql(customerNo)
	Dim vSQL

	vSQL = ""
	vSQL = "SELECT sum(�|�C���g�c) AS point_zan "
	vSQL = vSQL & " , min(�|�C���g����) AS point_expire_date "
	vSQL = vSQL & " FROM " & gLinkServer & "�|�C���g���� WITH (NOLOCK) "
	vSQL = vSQL & " WHERE "
	vSQL = vSQL & " (�|�C���g���� = "
	vSQL = vSQL & "  (SELECT min(�|�C���g����) FROM " & gLinkServer & "�|�C���g���� WITH (NOLOCK)"
	vSQL = vSQL & "   WHERE �ڋq�ԍ� = " & customerNo
	vSQL = vSQL & "     AND �|�C���g���t Is Not Null "
	vSQL = vSQL & "     AND �|�C���g�c Is Not Null "
	vSQL = vSQL & "     AND �|�C���g�c <> 0 "
	vSQL = vSQL & "     AND (�|�C���g���� Is Null "
'	vSQL = vSQL & "      OR �|�C���g���� >= CONVERT(datetime, '" & Now() & "')))) " ' 2016.04.27 GV mod
	vSQL = vSQL & "      OR �|�C���g���� >= CONVERT(datetime, '" & Date() & "')))) " '2016.04.27 GV add
	vSQL = vSQL & " AND �ڋq�ԍ� = " & customerNo
	vSQL = vSQL & " AND �|�C���g���t Is Not Null "
	vSQL = vSQL & " AND �|�C���g�c Is Not Null "
	vSQL = vSQL & " AND �|�C���g�c <> 0 "

	createPointSql = vSQL
End Function

'========================================================================
'
'	Function	DB����擾�����f�[�^����I�u�W�F�N�g�𐶐�
'
'========================================================================
Function createJsonObject(vRS)
	Dim pointZan
	Dim pointExpireDate

	pointZan = 0
	pointExpireDate = ""

	If vRS.EOF = False Then
		' �|�C���g�c
		If (IsNull(vRS("point_zan"))) Then
			pointZan = 0
		Else
			pointZan = CStr(Trim(vRS("point_zan")))
		End If

		' �|�C���g����
		If (IsNull(vRS("point_expire_date"))) Then
			pointExpireDate = ""
		Else
			pointExpireDate = CStr(Trim(vRS("point_expire_date")))
		End If
	End If


	' �|�C���g�c
	oJSON.data.Add "point_zan" ,pointZan

	' �|�C���g����
	oJSON.data.Add "point_expire_date" ,pointExpireDate
End Function
'========================================================================
%>
