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
'	Emax�󒍒����L�����Z�� �擾API
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)

Dim oJSON						' JSON�I�u�W�F�N�g
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim wGiftCustomerNo				' �M�t�g�ڋq�ԍ�
Dim wGiftNo						' �M�t�g�ԍ�
Dim wOrderGift					' �M�t�g�����t���O

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
' Get�p�����[�^
wCustomerNo = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))

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

Dim createDate
Dim totalOrderAmount
Dim usedPoint
Dim orderTotalOrderAmount

Set oJSON = New aspJSON

'--- �󒍒����L�����Z������o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      order_cancel.�����L�����Z���ԍ� "
vSQL = vSQL & "    , order_cancel.�󒍔ԍ� "
vSQL = vSQL & "    , order_cancel.�{�o�^�� "
vSQL = vSQL & "    , order_cancel.���i���v���z "
vSQL = vSQL & "    , order_cancel.���̑����v���z "
vSQL = vSQL & "    , order_cancel.���� "
vSQL = vSQL & "    , order_cancel.����萔�� "
vSQL = vSQL & "    , order_cancel.�O�ō��v���z "
vSQL = vSQL & "    , order_cancel.���ō��v���z "
vSQL = vSQL & "    , order_cancel.�󒍍��v���z "
vSQL = vSQL & "    , order_cancel.�ߕs�����E���z "
vSQL = vSQL & "    , order_cancel.���p�|�C���g "
vSQL = vSQL & "    , order_cancel.���v���z "
vSQL = vSQL & "    , order_cancel.�l���������� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "�󒍒����L�����Z�� order_cancel WITH (NOLOCK) "
vSQL = vSQL & " INNER JOIN " & gLinkServer & "�� orders WITH (NOLOCK) "
vSQL = vSQL & "   ON orders.�󒍔ԍ� = order_cancel.�󒍔ԍ� "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        orders.�󒍔ԍ� = " & wOrderNo
vSQL = vSQL & "    AND orders.�ڋq�ԍ� = " & wCustomerNo & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' ���X�g�ǉ�
	oJSON.data.Add "data" ,oJSON.Collection()

	' �{�o�^��
	If (IsNull(vRS("�{�o�^��"))) Then
		createDate = ""
	Else
		createDate = CStr(Trim(vRS("�{�o�^��")))
	End If

	' ���v���z
	If (IsNull(vRS("���v���z"))) Then
		totalOrderAmount = 0
	Else
		totalOrderAmount = CDbl(vRS("���v���z"))
	End If

	' ���p�|�C���g
	If (IsNull(vRS("���p�|�C���g"))) Then
		usedPoint = 0
	Else
		usedPoint = CDbl(vRS("���p�|�C���g"))
	End If

	' �󒍍��v���z
	If (IsNull(vRS("�󒍍��v���z"))) Then
		orderTotalOrderAmount = 0
	Else
		orderTotalOrderAmount = CDbl(vRS("�󒍍��v���z"))
	End If

	With oJSON.data("data")
		.Add "order_no", CStr(Trim(vRS("�󒍔ԍ�")))
		.Add "order_cancel_no", CStr(Trim(vRS("�����L�����Z���ԍ�")))
		.Add "create_date", createDate
		.Add "item_total_amount", CStr(Trim(vRS("���i���v���z")))
		.Add "other_total_amount", CStr(Trim(vRS("���̑����v���z")))
		.Add "ff_charge", CDbl(vRS("����")) 
		.Add "cod_charge", CStr(Trim(vRS("����萔��")))
		.Add "tax_am", CStr(Trim(vRS("�O�ō��v���z")))
		.Add "tax_in", CStr(Trim(vRS("���ō��v���z")))
		.Add "order_total_order_amount", orderTotalOrderAmount ' �󒍍��v���z
		.Add "kabusoku_am", CStr(Trim(vRS("�ߕs�����E���z")))
		.Add "used_point", usedPoint ' ���p�|�C���g
		.Add "total_order_amount", totalOrderAmount ' ���v���z
		.Add "after_discount_tax", CStr(Trim(vRS("�l����������")))
	End With
End If

'���R�[�h�Z�b�g�����
vRS.Close

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
%>
