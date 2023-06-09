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
'	�w�������ꗗ�y�[�W (�S���҈ꗗ)
'
'
'�ύX����
'2022.03.23 GV �V�K�쐬�B(�ƎҌ����T�C�g)(#3110)
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrDesc
Dim wFlg						' ���s�t���O
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderHidden				' ��\���t���O
Dim wOrderCancelled				' �L�����Z�������t���O
Dim wOrderShipping				' �����������t���O
Dim wOrderGift					' �M�t�g�����t���O
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

'��\���t���O
wOrderHidden = ReplaceInput_NoCRLF(Trim(Request("hide")))
If ((IsNull(wOrderHidden) = True) Or (UCase(wOrderHidden) <> "Y")) Then
	wOrderHidden = "N"
Else
	wOrderHidden = "Y"
End If

'�L�����Z�������t���O
wOrderCancelled = ReplaceInput_NoCRLF(Trim(Request("cancelled")))
If ((IsNull(wOrderCancelled) = True) Or (UCase(wOrderCancelled) <> "Y")) Then
	wOrderCancelled = "N"
Else
	wOrderCancelled = "Y"
End If

'�����������t���O
wOrderShipping = ReplaceInput_NoCRLF(Trim(Request("shipping")))
If ((IsNull(wOrderShipping) = True) Or (UCase(wOrderShipping) <> "Y")) Then
	wOrderShipping = "N"
Else
	wOrderShipping = "Y"
End If

'�M�t�g�����t���O
wOrderGift = ReplaceInput_NoCRLF(Trim(Request("gift")))
If ((IsNull(wOrderGift) = True) Or (UCase(wOrderGift) <> "Y")) Then
	wOrderGift = "N"
Else
'	wOrderGift = "Y" 'TODO: �M�t�g�����t���O��L���ɂ���ꍇ�A���̍s�̃R�����g�A�E�g���O��
	wOrderGift = "N" 'TODO: �M�t�g�����t���O��L���ɂ���ꍇ�A���̍s������
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
Dim vRS
Dim tantouParam
Dim tantouColumn

Set oJSON = New aspJSON

' ���͒l������̏ꍇ
If (wFlg = True) Then
	'-----------------------------------------------------------
	' �Y���ڋq�̎󒍂̒S���Ҏ����ꗗ���o��
	'-----------------------------------------------------------
	tantouParam  = "tantou_name"
	tantouColumn = "�����S����"

	vSQL = createTantouListSql(tantouParam, tantouColumn)

	'@@@@Response.Write vSQL & "<br>"

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂��Ă���ꍇ
	If vRS.EOF = False Then
		createJsonObject vRS, tantouParam
	End If

	'���R�[�h�Z�b�g�����
	vRS.Close

	'-----------------------------------------------------------
	' �Y���ڋq�̎󒍂̒S����e_mail�ꗗ���o��
	'-----------------------------------------------------------
	tantouParam  = "tantou_email"
	tantouColumn = "�ڋqE_mail"

	vSQL = createTantouListSql(tantouParam, tantouColumn)

	'@@@@Response.Write vSQL & "<br>"

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂��Ă���ꍇ
	If vRS.EOF = False Then
		createJsonObject vRS, tantouParam
	End If

	'���R�[�h�Z�b�g�����
	vRS.Close


	'���R�[�h�Z�b�g�̃N���A
	Set vRS = Nothing
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
'
'	Function	�S���҈ꗗ�̎擾SQL
'
'========================================================================
Function createTantouListSql(tantouParam, tantouColumn)
	Dim vSQL
	Dim orderType

	' �󒍌`��(�J���}��؂�Ŏw��)
	orderType = ""
	orderType = orderType & "  'E-mail'"
	orderType = orderType & " ,'FAX'"
	orderType = orderType & " ,'�C���^�[�l�b�g'"
	orderType = orderType & " ,'�g��'"
	orderType = orderType & " ,'�d�b'"
	orderType = orderType & " ,'�X��'"
	orderType = orderType & " ,'���X'"
	orderType = orderType & " ,'�X�}�[�g�t�H��'"
	orderType = orderType & " ,'�M�t�g'"

	vSQL = ""
	vSQL = vSQL & "SELECT DISTINCT o." & tantouParam & " "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & " (SELECT DISTINCT "
	vSQL = vSQL & "   o1.�ڋq�ԍ� "
	vSQL = vSQL & "  ,o1.�󒍔ԍ� "
	vSQL = vSQL & "  ,o1.���ϓ� "
	vSQL = vSQL & "  ,o1.�폜�� "
	vSQL = vSQL & "  ,ov.��\���t���O "
	vSQL = vSQL & "  ,ISNULL(o1." & tantouColumn & ", '') AS " & tantouParam & " "
	vSQL = vSQL & "  FROM �� AS o1 "
	vSQL = vSQL & "    INNER JOIN �󒍖��� od1 WITH (NOLOCK) "
	vSQL = vSQL & "      ON od1.�󒍔ԍ� = o1.�󒍔ԍ� "
	vSQL = vSQL & "     AND od1.�Z�b�g�i�e���הԍ� = 0 "
	vSQL = vSQL & "    LEFT JOIN �󒍔�\�����X�g ov WITH (NOLOCK) "
	vSQL = vSQL & "      ON ov.�󒍔ԍ� = od1.�󒍔ԍ� "
	vSQL = vSQL & "     AND ov.�󒍖��הԍ� = od1.�󒍖��הԍ� "
	vSQL = vSQL & "  WHERE o1.�ڋq�ԍ� = " & wCustomerNo & " "
	vSQL = vSQL & "    AND o1.�󒍌`�� IN (" & orderType & ") "

	' �����������t���O
	If wOrderShipping = "Y" Then
		vSQL = vSQL & "    AND od1.�󒍐��� > od1.�o�׍��v���� "
	End If

	' ��\���t���O
	If wOrderHidden = "Y" Then
		vSQL = vSQL & "    AND ov.��\���t���O = 'Y' "
	Else
		'�M�t�g���[�h�ł͂Ȃ�
		If (wOrderGift = "N") Then
			vSQL = vSQL & "    AND ov.��\���t���O IS NULL "
		End If
	End If

	' �L�����Z�������t���O
	If wOrderCancelled = "Y" Then
		'vSQL = vSQL & "  AND o1.�폜�� IS NOT NULL "
		vSQL = vSQL & "  AND od1.Web�L�����Z���t���O = 'Y' "
	Else
		If wOrderHidden = "Y" Then
		'��\���t���O��Y�̏ꍇ�A���w��
		Else
			vSQL = vSQL & "  AND o1.�폜�� IS NULL "
			vSQL = vSQL & "  AND ISNULL(od1.Web�L�����Z���t���O, 'N') <> 'Y' "
		End If
	End If

	vSQL = vSQL & " ) AS o "
	vSQL = vSQL & "WHERE o." & tantouParam & " <> '' "
	vSQL = vSQL & "ORDER BY o." & tantouParam & " ASC "

	createTantouListSql = vSQL
End Function

'========================================================================
'
'	Function	DB����擾�����f�[�^����I�u�W�F�N�g�𐶐�
'
'========================================================================
Function createJsonObject(vRS, tantouParam)
	Dim j
	Dim tantouListParam

	j = 0
	tantouListParam = tantouParam & "_list"

	' ���X�g�ǉ�
	oJSON.data.Add tantouListParam ,oJSON.Collection()

	' ���R�[�h�Z�b�g�̍Ō�܂Ń��[�v
	Do Until vRS.EOF
		'--- ���׍s����
		With oJSON.data(tantouListParam)
			.Add j, CStr(Trim(vRS(tantouParam)))
		End With

		j = j + 1

		' ���R�[�h�Z�b�g�̃|�C���^�����̍s�ֈړ�
		vRS.MoveNext
	Loop
End Function
'========================================================================
%>
