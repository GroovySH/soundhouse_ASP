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

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wUserID

Dim oJSON						' JSON�I�u�W�F�N�g
Dim wOrderNo					' �󒍔ԍ�

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
' Get�p�����[�^
wUserID = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("order_no")))

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

Dim orderDate
Dim shippingDate
Dim estimateDate
Dim one_time_todokesaki
Dim final_nouki_date_time
Dim receiptName
Dim receiptNote

Set oJSON = New aspJSON


one_time_todokesaki = ""
final_nouki_date_time = ""
receiptName = ""
receiptNote = ""

'--- �w�b�_�����̏���o��
vSQL = ""
vSQL = vSQL & "SELECT TOP 1 "
vSQL = vSQL & "      a.�󒍔ԍ� "
vSQL = vSQL & "    , a.���ϓ� "
vSQL = vSQL & "    , a.�󒍓� "
vSQL = vSQL & "    , a.�o�׊����� "
vSQL = vSQL & "    , a.�󒍌`�� "
vSQL = vSQL & "    , a.�x�����@ "
vSQL = vSQL & "    , a.���i���v���z "
vSQL = vSQL & "    , a.���� "
vSQL = vSQL & "    , a.����萔�� "
vSQL = vSQL & "    , a.�󒍍��v���z "
vSQL = vSQL & "    , a.�ꊇ�o�׃t���O "
vSQL = vSQL & "    , a.�̎������� "
vSQL = vSQL & "    , a.�̎����A������ "
vSQL = vSQL & "    , a.Web�󒍕ύX�J�n�� "
vSQL = vSQL & "    , a.����ŗ� "
vSQL = vSQL & "    , a.�^����ЃR�[�h "
vSQL = vSQL & "    , a.�S���҃R�[�h "
vSQL = vSQL & "    , b.�������͐�X�֔ԍ� "
vSQL = vSQL & "    , b.�������͐�s���{�� "
vSQL = vSQL & "    , b.�������͐�Z�� "
vSQL = vSQL & "    , b.�������͐於�O "
vSQL = vSQL & "    , b.�ŏI�w��[�� "
vSQL = vSQL & "    , b.�ŏI���Ԏw�� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "��     a WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "�󒍖��� b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        b.�󒍔ԍ� = a.�󒍔ԍ� "
vSQL = vSQL & "    AND a.�󒍔ԍ� = " & wOrderNo & " "
vSQL = vSQL & "    AND a.�ڋq�ԍ� = " & wUserID & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' ���X�g�ǉ�
	oJSON.data.Add "data" ,oJSON.Collection()

	' �󒍓�
	If (IsNull(vRS("�󒍓�"))) Then
		orderDate = ""
	Else
		orderDate = CStr(Trim(vRS("�󒍓�")))
	End If

	' ���ϓ�
	If (IsNull(vRS("���ϓ�"))) Then
		estimateDate = ""
	Else
		estimateDate = CStr(Trim(vRS("���ϓ�")))
	End If

	' �o�׊�����
	If (IsNull(vRS("�o�׊�����"))) Then
		shippingDate = ""
	Else
		shippingDate = CStr(Trim(vRS("�o�׊�����")))
	End If

	' �������͐�
	one_time_todokesaki = vRS("�������͐�X�֔ԍ�")&"^"&_
		vRS("�������͐�s���{��")&"^"&_
		vRS("�������͐�Z��")&"^"&_
		vRS("�������͐於�O")

	one_time_todokesaki = Replace(one_time_todokesaki, """", "�h")

	' �ŏI�w��[���Ǝ���
	final_nouki_date_time = vRS("�ŏI�w��[��")&"_"&vRS("�ŏI���Ԏw��") 

	' �̎�������
	If (IsNull(vRS("�̎�������"))) Then
		receiptName = ""
	Else
		receiptName = CStr(Trim(vRS("�̎�������")))
		receiptName = Replace(receiptName, """", "�h")
	End If

	' �̎����A������
	If (IsNull(vRS("�̎����A������"))) Then
		receiptNote = ""
	Else
		receiptNote = CStr(Trim(vRS("�̎����A������")))
		receiptNote = Replace(receiptNote, """", "�h")
	End If

	With oJSON.data("data")
		.Add "order_no", CStr(Trim(vRS("�󒍔ԍ�")))
		.Add "estimate_date", estimateDate
		.Add "order_date", orderDate
		.Add "shipping_date", shippingDate
		.Add "order_type", CStr(Trim(vRS("�󒍌`��")))
		.Add "payment_method",  CStr(Trim(vRS("�x�����@")))
		.Add "total_item_amount", CDbl(Trim(vRS("���i���v���z")))
		.Add "freight_charge", CDbl(vRS("����")) 
		.Add "daibiki_charge", CDbl(vRS("����萔��"))
		.Add "total_order_amount", CDbl(vRS("�󒍍��v���z"))
		.Add "combined_shipping_flag", CStr(Trim(vRS("�ꊇ�o�׃t���O")))
		.Add "receipt_name", receiptName
		.Add "receipt_note", receiptNote
'				.Add "web_order_modify_start_date", vRS("Web�󒍕ύX�J�n��")
		.Add "tax_rate", CDbl(vRS("����ŗ�"))
		.Add "freight_forwarder_cd", CStr(vRS("�^����ЃR�[�h"))
		.Add "tantou_cd", CStr(vRS("�S���҃R�[�h"))
		.Add "one_time_todokesaki", one_time_todokesaki
		.Add "final_nouki_date_time", final_nouki_date_time
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
