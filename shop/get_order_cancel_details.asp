get_order_cancel<%@ LANGUAGE="VBScript" %>
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
'	Emax�󒍖��ג����L�����Z�� �擾API
'========================================================================
'On Error Resume Next

Const PAGE_SIZE = 10			' �w����������1�y�[�W������̕\���s��

Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim oJSON						' JSON�I�u�W�F�N�g
Dim wFlg						' ���s�t���O

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

'�󒍔ԍ�
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wOrderNo = null
Else
	wOrderNo = CLng(wOrderNo)
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
Dim vWHERE
Dim i
Dim j
Dim vRS
Dim point

Set oJSON = New aspJSON

' ������
i = 0
j = 0

' ���͒l������̏ꍇ
If (wFlg = True) Then

	'--- �󒍖��ג����L�����Z������o��
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "      order_cancel_details.�����L�����Z���ԍ� "
	vSQL = vSQL & "    , order_cancel_details.�����L�����Z�����הԍ� "
	vSQL = vSQL & "    , order_cancel_details.�󒍔ԍ� "
	vSQL = vSQL & "    , order_cancel_details.�󒍖��הԍ� "
	vSQL = vSQL & "    , order_cancel_details.���[�J�[�R�[�h "
	vSQL = vSQL & "    , order_cancel_details.���i�R�[�h "
	vSQL = vSQL & "    , order_cancel_details.���i�� "
	vSQL = vSQL & "    , order_cancel_details.�F "
	vSQL = vSQL & "    , order_cancel_details.�K�i "
	vSQL = vSQL & "    , order_cancel_details.�Z�b�g�i�t���O "
	vSQL = vSQL & "    , order_cancel_details.�Z�b�g�i�e���הԍ� "
	vSQL = vSQL & "    , order_cancel_details.�����L�����Z������ "
	vSQL = vSQL & "    , order_cancel_details.�����L�����Z���P�� "
	vSQL = vSQL & "    , order_cancel_details.�����L�����Z�����z "
	vSQL = vSQL & "    , order_cancel_details.�O�� "
	vSQL = vSQL & "    , order_cancel_details.���� "
	vSQL = vSQL & "    , order_cancel_details.�|�C���g "
	vSQL = vSQL & "    , order_cancel_details.�N�[�|���l���� "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "      " & gLinkServer & "�󒍒����L�����Z�� order_cancel WITH (NOLOCK) "
	vSQL = vSQL & " INNER JOIN " & gLinkServer & "�� orders WITH (NOLOCK) "
	vSQL = vSQL & "   ON orders.�󒍔ԍ� = order_cancel.�󒍔ԍ� "
	vSQL = vSQL & " INNER JOIN " & gLinkServer & "�󒍖��ג����L�����Z�� order_cancel_details WITH (NOLOCK) "
	vSQL = vSQL & "   ON order_cancel.�󒍔ԍ� = order_cancel.�󒍔ԍ� "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "        orders.�󒍔ԍ� = " & wOrderNo
	vSQL = vSQL & "    AND orders.�ڋq�ԍ� = " & wCustomerNo & " "

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂��Ă���ꍇ
	If vRS.EOF = False Then

		' ���X�g�ǉ�
		oJSON.data.Add "list" ,oJSON.Collection()

		For i = 0 To (vRS.RecordCount - 1)
			' �|�C���g
			If (IsNull(vRS("�|�C���g"))) Then
				point = 0
			Else
				point = CDbl(vRS("�|�C���g"))
			End If

			'--- ���׍s����
			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
					.Add "order_cancel_no" ,CStr(Trim(vRS("�����L�����Z���ԍ�")))
					.Add "order_cancel_detail_no" ,CStr(Trim(vRS("�����L�����Z�����הԍ�")))
					.Add "o_no" ,CStr(Trim(vRS("�󒍔ԍ�")))
					.Add "od_no" ,CStr(Trim(vRS("�󒍖��הԍ�")))
					.Add "maker_code" ,CStr(Trim(vRS("���[�J�[�R�[�h")))
					.Add "i_cd" ,CStr(Trim(vRS("���i�R�[�h")))
					.Add "i_name" ,CStr(Trim(vRS("���i��")))
					.Add "iro" ,CStr(Trim(vRS("�F")))
					.Add "kikaku" ,CStr(Trim(vRS("�K�i")))
					.Add "set_item_flg" ,CStr(Trim(vRS("�Z�b�g�i�t���O")))
					.Add "set_item_detail_no" ,CStr(Trim(vRS("�Z�b�g�i�e���הԍ�")))
					.Add "i_suu" ,CStr(Trim(vRS("�����L�����Z������")))
					.Add "i_tanka" ,CStr(Trim(vRS("�����L�����Z���P��")))
					.Add "i_cancel_am" ,CStr(Trim(vRS("�����L�����Z�����z")))
					.Add "ext_tax" ,CStr(Trim(vRS("�O��")))
					.Add "inc_tax" ,CStr(Trim(vRS("����")))
					.Add "point" ,point '�|�C���g
					.Add "coupon_discount" ,CStr(Trim(vRS("�N�[�|���l����")))
				End With
			End With

			' ���̃��R�[�h�s�ֈړ�
			vRS.MoveNext

			If vRS.EOF Then
				Exit For
			End If

			j = j + 1
		Next
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
%>
