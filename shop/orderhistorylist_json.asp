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
'
'
'�ύX����
'2014/09/16 GV �V�K�쐬
'
'========================================================================
'On Error Resume Next

Const PAGE_SIZE = 20						' �w����������1�y�[�W������̕\���s��

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wUserID


Dim wIPage						' �\������y�[�W�ʒu (�p�����[�^)
Dim oJSON						' JSON�I�u�W�F�N�g
Dim listType					' �擾���郊�X�g�^�C�v�i1...���ς��蒆/�o�׏������A2...�w�������j


'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
' Get�p�����[�^
wUserID = ReplaceInput(Trim(Request("cno")))
listType = ReplaceInput(Trim(Request("list_type")))
wIPage = ReplaceInput(Trim(Request("page")))	' �y�[�W�ʒu

'�y�[�W�ԍ�
If wIPage = "" Or IsNumeric(wIPage) = False Then
	wIPage = 1
Else
	wIPage = CLng(wIPage)
End If


' �擾���郊�X�g�^�C�v
If (listType = "") Or (IsNumeric(listType) = False) Then
	listType = 1
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
Dim i
Dim j
Dim vRS
Dim vParam
Dim vTitleWord
Dim vTitleWordSave
Dim vOrderDateLabel
Dim vHistoryCount
Dim vHTML
Dim orderDate
Dim shippingDate

Set oJSON = New aspJSON

' �C�e���[�^������
i = 0
j = 0


'--- �Y���ڋq�̎󒍈ꗗ���o��1 (���ϒ��E�o�׏�����)
If (listType = 1) Then
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "      a.�󒍔ԍ� "
	vSQL = vSQL & "    , a.�󒍓� "
	vSQL = vSQL & "    , a.���ϓ� "
	vSQL = vSQL & "    , a.�o�׊����� "
	vSQL = vSQL & "    , a.�󒍌`�� "
	vSQL = vSQL & "    , a.�x�����@ "
	vSQL = vSQL & "    , a.Web�󒍕ύX�J�n�� "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "    " & gLinkServer & "�� a WITH (NOLOCK) "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "        a.�폜��     IS NULL "
	vSQL = vSQL & "    AND a.�o�׊����� IS NULL "
	'vSQL = vSQL & "    AND a.�󒍌`�� in ('E-mail','FAX','�C���^�[�l�b�g','�g��','�d�b','�X��','���X')"	'2012/11/24 ok Del
	vSQL = vSQL & "    AND a.�󒍌`�� in ('E-mail','FAX','�C���^�[�l�b�g','�g��','�d�b','�X��','���X','�X�}�[�g�t�H��')"	'2012/11/24 ok Add
	vSQL = vSQL & "    AND a.�ڋq�ԍ�   = " & wUserID & " "
	vSQL = vSQL & "ORDER BY "
	vSQL = vSQL & "      CASE WHEN a.�󒍓� IS NULL "
	vSQL = vSQL & "          THEN 1 "
	vSQL = vSQL & "          ELSE 2 "
	vSQL = vSQL & "      END "
	vSQL = vSQL & "    , ���ϓ� DESC "

	'@@@@Response.Write(vSQL)

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	If vRS.EOF = False Then

		' ���X�g�ǉ�
		oJSON.data.Add "list" ,oJSON.Collection()

		Do Until vRS.EOF = True
			'--- �o�׏�(�^�C�g��) �̔���
			vTitleWord = make_titleWord(vRS("�󒍓�"), vRS("�o�׊�����"))

			'--- ��������̃^�C�g�����x������
			If vTitleWord = "������" Then
				vOrderDateLabel = "�����ϓ�"
			ElseIf vTitleWord = "�o�׏�����" Then
				vOrderDateLabel = "��������"
			ElseIf vTitleWord = "���w������" Then
				vOrderDateLabel = "��������"
			Else
				vOrderDateLabel = "��������"
			End If

			' �󒍓�
			If (IsNull(vRS("�󒍓�"))) Then
				orderDate = ""
			Else
				orderDate = CStr(Trim(vRS("�󒍓�")))
			End If


			' �o�׊�����
			If (IsNull(vRS("�o�׊�����"))) Then
				shippingDate = ""
			Else
				shippingDate = CStr(Trim(vRS("�o�׊�����")))
			End If


			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
'					.Add "title" ,vTitleWord
'					.Add "list" ,oJSON.Collection()
'					With .item("list")
						.Add "order_date" ,orderDate
						.Add "estimate_date" ,CStr(Trim(vRS("���ϓ�")))
						.Add "order_no" ,CStr(Trim(vRS("�󒍔ԍ�")))
						.Add "order_type" ,CStr(Trim(vRS("�󒍌`��")))
						.Add "payment_method" ,get_paymetMethodWord(vRS("�x�����@"))
						.Add "shipping_date" , shippingDate
'					End With
				End With
			End With

			vRS.MoveNext
			j = j + 1
		Loop
	End If

	'���R�[�h�Z�b�g�����
	vRS.Close
End If

If (listType = 2) Then
'--- �Y���ڋq�̎󒍈ꗗ���o��2 (���w������)
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "      a.�󒍔ԍ� "
	vSQL = vSQL & "    , a.�󒍓� "
	vSQL = vSQL & "    , a.���ϓ� "
	vSQL = vSQL & "    , a.�o�׊����� "
	vSQL = vSQL & "    , a.�󒍌`�� "
	vSQL = vSQL & "    , a.�x�����@ "
	vSQL = vSQL & "    , a.Web�󒍕ύX�J�n�� "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "    " & gLinkServer & "�� a WITH (NOLOCK) "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "        a.�폜��     IS NULL "
	vSQL = vSQL & "    AND a.�o�׊����� IS NOT NULL "
	'vSQL = vSQL & "    AND a.�󒍌`�� in ('E-mail','FAX','�C���^�[�l�b�g','�g��','�d�b','�X��','���X')"	'2012/11/24 ok Del
	vSQL = vSQL & "    AND a.�󒍌`�� in ('E-mail','FAX','�C���^�[�l�b�g','�g��','�d�b','�X��','���X','�X�}�[�g�t�H��')"	'2012/11/24 ok Add
	vSQL = vSQL & "    AND a.�ڋq�ԍ�   = " & wUserID & " "
	vSQL = vSQL & "ORDER BY "
	vSQL = vSQL & "    ���ϓ� DESC "

	'@@@@Response.Write(vSQL)

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	If vRS.EOF = False Then

		' �S������JSON�f�[�^�ɃZ�b�g
		oJSON.data.Add "count" ,vRS.RecordCount

		' ���X�g�ǉ�
		oJSON.data.Add "list" ,oJSON.Collection()

		'--- �o�׏�(�^�C�g��) ��������
		vTitleWord = "���w������"
		vOrderDateLabel = "��������"

		'--- �w��y�[�W��\������ׂ̃��R�[�h�ʒu�t��(SearchList�̏����ɕ키)
		vRS.PageSize = PAGE_SIZE
		If wIPage > ((vRS.RecordCount + (PAGE_SIZE - 1)) / PAGE_SIZE) Then		'MAX�y�[�W�𒴂���ꍇ�͍ŏI�y�[�W��
			wIPage = Fix(vRS.RecordCount / PAGE_SIZE)
		End If

		' ���R�[�h�ʒu�̈ʒu�t��
		vRS.AbsolutePage = wIPage

		For i = 0 To (vRS.PageSize - 1)

			' �󒍓�
			orderDate = vRS("�󒍓�")
			If (IsNull(vRS("�󒍓�"))) Then
				orderDate = ""
			Else
				orderDate = CStr(Trim(vRS("�󒍓�")))
			End If

			' �o�׊�����
			If (IsNull(vRS("�o�׊�����"))) Then
				shippingDate = ""
			Else
				shippingDate = CStr(Trim(vRS("�o�׊�����")))
			End If


			'--- ���׍s����
			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
'					.Add "title" ,vTitleWord
'					.Add "list" ,oJSON.Collection()
'					With .item("list")
						.Add "order_date" ,orderDate
						.Add "estimate_date" ,CStr(Trim(vRS("���ϓ�")))
						.Add "order_no" ,CStr(Trim(vRS("�󒍔ԍ�")))
						.Add "order_type" ,CStr(Trim(vRS("�󒍌`��")))
						.Add "payment_method" ,get_paymetMethodWord(vRS("�x�����@"))
						.Add "shipping_date" , shippingDate
'					End With
				End With
			End With


			vRS.MoveNext

			If vRS.EOF Then
				Exit For
			End If

			j = j + 1
		Next
	End If

	'���R�[�h�Z�b�g�����
	vRS.Close
End If

'���R�[�h�Z�b�g�̃N���A
Set vRS = Nothing

' -------------------------------------------------
' JSON�f�[�^�̕ԋp
' -------------------------------------------------
' �w�b�_�o��
Response.AddHeader "Content-Type", "application/json"
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
'	Function	�\���p�x�������@�����̐���
'
'	Note
'	  �x�����@              �\������
'	��������������������������������������������
'	  �R���r�j�x��       �� "�R���r�j����"
'	  �l�b�g�o���L���O   �� "�R���r�j����"
'	  �䂤����           �� "�R���r�j����"
'	  ���[��(��������)   �� "���[��"
'	  ���[��(�����Ȃ�)   �� "���[��"
'	  ���[��(��������)   �� "���[��"
'	  ��s�U��           �� "��s�U��"
'	  �����             �� "�������"
'	  ����               �� (�x�����@���̂܂�)
'	  ���|               �� (�x�����@���̂܂�)
'	  �A�}�]��           �� (�x�����@���̂܂�)
'	  �N���W�b�g�J�[�h   �� (�x�����@���̂܂�)
'
'========================================================================
Function get_paymetMethodWord(pstrPaymetMethod)

Dim vDisplayWord

If IsNull(pstrPaymetMethod) = True Then
	' Null �͔���s�\
	Exit Function
End If

If pstrPaymetMethod = "�����" Then
	vDisplayWord = "�������"
ElseIf pstrPaymetMethod = "�R���r�j�x��" Then
	vDisplayWord = "�R���r�j����"
ElseIf pstrPaymetMethod = "�l�b�g�o���L���O" Then
	vDisplayWord = "�R���r�j����"
ElseIf pstrPaymetMethod = "�䂤����" Then
	vDisplayWord = "�R���r�j����"
ElseIf pstrPaymetMethod = "��s�U��" Then
	vDisplayWord = "��s�U��"
ElseIf InStr(pstrPaymetMethod, "���[��") > 0 Then
	vDisplayWord = "���[��"
Else
	vDisplayWord = pstrPaymetMethod
End If

get_paymetMethodWord = vDisplayWord

End Function

'========================================================================
'
'	Function	�w�������̃^�C�g����������
'
'========================================================================
Function make_titleWord(pdatOrderDate, pdatShipCompleteDate)

Dim vTitleWord

If IsNull(pdatOrderDate) Then
	'--- �󒍓���Null�̏ꍇ
	vTitleWord = "������"
ElseIf IsNull(pdatOrderDate) = False And IsNull(pdatShipCompleteDate) Then
	'--- �󒍓���Null�łȂ��A�o�׊�������Null�̏ꍇ
	vTitleWord = "�o�׏�����"
ElseIf IsNull(pdatShipCompleteDate) = False Then
	'--- �o�׊�������Null�̏ꍇ
	vTitleWord = "���w������"
Else
	vTitleWord = "���w������"
End If

make_TitleWord = vTitleWord

End Function

'========================================================================
%>
