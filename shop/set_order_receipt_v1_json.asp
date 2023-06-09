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
'	[��]�e�[�u���́u�̎����v�֘A�J�����̒l���X�V����B
'
'
'�ύX����
'2020.02.05 GV �V�K�쐬
'2020.11.07 GV CStr�֐��̏C���B(#2589)
'
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wFlg						' ���s�t���O
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
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
Dim vRS1
Dim vRS2
Dim vCustomerNo
Dim vOrderNo
Dim vReceiptAmount
Dim vReceiptFlag
Dim vReceiptNo
Dim vReceiptDate
Dim vReceiptName
Dim vReceiptNote
Dim vModified
Dim vModifyTantouCd

vReceiptAmount = "0"
vReceiptFlag = ""
vReceiptNo = "-1"
vReceiptDate = null
vReceiptName = ""
vReceiptNote = ""
vModified = Now()
vModifyTantouCd = "Internet"

Set oJSON = New aspJSON


' ���͒l������̏ꍇ
If (wFlg = True) Then
	vSQL = ""
	vSQL = vSQL & "SELECT DISTINCT "
	vSQL = vSQL & "  T1.�ڋq�ԍ� "
	vSQL = vSQL & " ,c.�ڋq�� "
	vSQL = vSQL & " ,T1.�󒍔ԍ� "
	vSQL = vSQL & " ,T1.�x�����@ "
	vSQL = vSQL & " ,(CASE WHEN T1.�x�����@ = '����' AND T1.�󒍌`�� = '���X' AND T1.�������v���z = 0 "
	vSQL = vSQL & "          THEN T1.���v���z "
	vSQL = vSQL & "        WHEN T1.�x�����@ = '����' "
	vSQL = vSQL & "          THEN T1.�������v���z "
	vSQL = vSQL & "        WHEN T1.�x�����@ = '��s�U��' "
	vSQL = vSQL & "          THEN T1.�������v���z "
' 2020.11.07 GV add start
	vSQL = vSQL & "        WHEN T1.�x�����@ = '�R���r�j�x��' "
	vSQL = vSQL & "          THEN T1.�������v���z "
' 2020.11.07 GV add end
	vSQL = vSQL & "        WHEN T1.�x�����@ = '�N���W�b�g�J�[�h' "
	vSQL = vSQL & "          THEN T1.���v���z "
	vSQL = vSQL & "        WHEN T1.�x�����@ = '���[��(��������)' "
	vSQL = vSQL & "          THEN "
	vSQL = vSQL & "            (SELECT ol.���[�����������z "
	vSQL = vSQL & "               FROM ��_���[����� ol WITH (NOLOCK) "
	vSQL = vSQL & "              WHERE ol.�󒍔ԍ� = T1.�󒍔ԍ�) "
	vSQL = vSQL & "        END) �̎������z "
	vSQL = vSQL & " ,T1.�̎������s�t���O "
	vSQL = vSQL & " ,T1.�̎����ԍ� "
	vSQL = vSQL & " ,T1.�̎������s�� "
	vSQL = vSQL & " ,T1.�̎������� "
	vSQL = vSQL & " ,T1.�̎����A������ "
	vSQL = vSQL & " ,T1.�ŏI�X�V�� "
	vSQL = vSQL & " ,T1.�ŏI�X�V�҃R�[�h "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "  �� T1 WITH (NOLOCK) "
	vSQL = vSQL & "INNER JOIN �ڋq c WITH (NOLOCK) "
	vSQL = vSQL & "   ON c.�ڋq�ԍ� = T1.�ڋq�ԍ� "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "  T1.�ڋq�ԍ� = " & wCustomerNo
	vSQL = vSQL & " AND T1.�󒍔ԍ� = " & wOrderNo

	Set vRS1 = Server.CreateObject("ADODB.Recordset")
	vRS1.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂���ꍇ
	If vRS1.EOF = False Then
		'�̎������z
		'vReceiptAmount = CStr(Trim(vRS1("�̎������z"))) ' 2020.11.07 GV mod
		' 2020.11.07 GV add start
		If (IsNull(vRS1("�̎������z"))) Then
			vReceiptAmount = ""
		Else
			vReceiptAmount = CStr(Trim(vRS1("�̎������z")))
		End If
		' 2020.11.07 GV add end

		'�̎������s�t���O
		vReceiptFlag = getReceiptFlag(vRS1("�x�����@"), wOrderNo)
		If vReceiptFlag <> "" Then
			'�̎����ԍ�
			vReceiptNo = getReceiptNo()
			If vReceiptNo <> "-1" Then
				'�̎������s��
				'���x�X�V
'				If (IsNull(vRS1("�̎������s��"))) Then
					vReceiptDate = vModified
'				Else
'					vReceiptDate = CStr(Trim(vRS1("�̎������s��")))
'				End If

				'�̎�������
				If ((IsNull(vRS1("�̎�������"))) Or (Trim(vRS1("�̎�������")) = "")) Then
					vReceiptName = CStr(Trim(vRS1("�ڋq��")))
					vReceiptName = Replace(vReceiptName, """", "�h")
				Else
					vReceiptName = CStr(Trim(vRS1("�̎�������")))
					vReceiptName = Replace(vReceiptName, """", "�h")
				End If

				'�̎����A������
				If ((IsNull(vRS1("�̎�������"))) Or (Trim(vRS1("�̎����A������")) = "")) Then
					vReceiptNote = getReceiptNote()
				Else
					vReceiptNote = CStr(Trim(vRS1("�̎����A������")))
					vReceiptNote = Replace(vReceiptNote, """", "�h")
				End If

				'�X�V
				vSQL = ""
				vSQL = vSQL & "SELECT "
				vSQL = vSQL & "  T1.�󒍔ԍ� "
				vSQL = vSQL & " ,T1.�̎������s�t���O "
				vSQL = vSQL & " ,T1.�̎����ԍ� "
				vSQL = vSQL & " ,T1.�̎������s�� "
				vSQL = vSQL & " ,T1.�̎������� "
				vSQL = vSQL & " ,T1.�̎����A������ "
				vSQL = vSQL & " ,T1.�ŏI�X�V�� "
				vSQL = vSQL & " ,T1.�ŏI�X�V�҃R�[�h "
				vSQL = vSQL & "FROM "
				vSQL = vSQL & "  �� T1 "
				vSQL = vSQL & "WHERE "
				vSQL = vSQL & "  T1.�ڋq�ԍ� = " & wCustomerNo
				vSQL = vSQL & " AND T1.�󒍔ԍ� = " & wOrderNo

				Set vRS2 = Server.CreateObject("ADODB.Recordset")
				vRS2.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

				If vRS2.EOF = False Then
					vRS2("�̎������s�t���O") = vReceiptFlag
					vRS2("�̎����ԍ�") = vReceiptNo
					vRS2("�̎������s��") = vReceiptDate
					'vRS2("�̎�������") = vReceiptName �X�V���Ȃ�
					'vRS2("�̎����A������") = vReceiptNote �X�V���Ȃ�
					vRS2("�ŏI�X�V��") = vModified
					vRS2("�ŏI�X�V�҃R�[�h") = vModifyTantouCd
					vRS2.Update
				End If

				'���R�[�h�Z�b�g�����
				vRS2.Close

				'���R�[�h�Z�b�g�̃N���A
				Set vRS2 = Nothing
			End If
		End If
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

oJSON.data.Add "ono" ,vOrderNo

'�u�̎����v�֘A�J����
oJSON.data.Add "receipt_am" ,vReceiptAmount
oJSON.data.Add "receipt_flg" ,vReceiptFlag
oJSON.data.Add "receipt_no" ,vReceiptNo
oJSON.data.Add "receipt_dt" ,vReceiptDate
oJSON.data.Add "receipt_name" ,vReceiptName
oJSON.data.Add "receipt_note" ,vReceiptNote

'�u�ŏI�X�V�v�֘A�J����
oJSON.data.Add "modified" ,vModified
oJSON.data.Add "modify_tantou_cd" ,vModifyTantouCd

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
'	Function	�̎����ԍ��̐���
'
'========================================================================

Function getReceiptNo()

Dim vSQL
Dim vRS
Dim vReceiptNo

vReceiptNo = -1

vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  a.item_num1 "
vSQL = vSQL & "  FROM �R���g���[���}�X�^ a WITH (ROWLOCK) "
vSQL = vSQL & " WHERE a.sub_system_cd = '����'"
vSQL = vSQL & "   AND a.item_cd = '�ԍ�'"
vSQL = vSQL & "   AND a.item_sub_cd = '�̎���'"

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then
	vReceiptNo = CLng(vRS("item_num1")) + 1
	vRS("item_num1") = vReceiptNo
	vRS.Update
End If

vRS.Close

getReceiptNo = CStr(Trim(vReceiptNo))

End Function

'========================================================================
'
'	Function	�̎����A�������̐���
'
'========================================================================

Function getReceiptNote()

Dim vSQL
Dim vRS
Dim vReceiptNote

vReceiptNote = "�����@���Ƃ���"

vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  a.item_char1 "
vSQL = vSQL & "  FROM �R���g���[���}�X�^ a WITH (NOLOCK) "
vSQL = vSQL & " WHERE a.sub_system_cd = '�̎���'"
vSQL = vSQL & "   AND a.item_cd = '�A������'"
vSQL = vSQL & "   AND a.item_sub_cd = '1'"

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then
	vReceiptNote = CStr(Trim(vRS("item_char1")))
	vReceiptNote = Replace(vReceiptNote, """", "�h")
End If

vRS.Close

getReceiptNote = vReceiptNote

End Function

'========================================================================
%>
