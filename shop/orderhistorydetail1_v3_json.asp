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
'2015.05.07 GV �ߕs�����E���z��ǉ�
'2016.02.05 GV �s�v�������폜�B(Web�����ύX�L�����Z���@�\)
'2016.06.01 GV ��\���t���O�̗L����ǉ��B
'2018.12.21 GV PayPal�Ή��B
'2020.02.05 GV ������DL�Ή��B
'2020.03.18 GV ������DL�Ή��B
'2020.06.30 GV �~���������X�g�Ή��B(#2841)
'
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

'�M�t�g�����t���O
wOrderGift = ReplaceInput_NoCRLF(Trim(Request("gift")))
If ((IsNull(wOrderGift) = True) Or (UCase(wOrderGift) <> "Y")) Then
	wOrderGift = "N"
Else
	wOrderGift = "Y"
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

Dim orderDate
Dim shippingDate
Dim estimateDate
Dim one_time_todokesaki1
Dim one_time_todokesaki2
Dim final_nouki_date_time
Dim receiptName
Dim receiptNote
Dim totalOrderAmount2
Dim usedPoint
Dim furikomiMeigi
Dim pos
Dim storeStop
Dim webModCancelFlg
Dim deleteDate
Dim hide
Dim paymentMethodDetail '2018.12.21 GV add
Dim receiptFlag '2020.02.05 GV add
Dim receiptDate '2020.02.05 GV add
Dim displayReceiptDate '2020.03.18 GV add
Dim giftCustomerNo '2021.06.30 GV add
Dim giftNo '2021.06.30 GV add

Set oJSON = New aspJSON


one_time_todokesaki1 = ""
one_time_todokesaki2 = ""
final_nouki_date_time = ""
receiptName = ""
receiptNote = ""
totalOrderAmount2 = 0
hide = ""

'-- ��\���t���O�����݂��Ă��邩
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  count(�󒍔ԍ�) as cnt "
vSQL = vSQL & " FROM "
vSQL = vSQL & "   �󒍔�\�����X�g ov WITH (NOLOCK) "
vSQL = vSQL & " WHERE "
vSQL = vSQL & " ov.�󒍔ԍ� = " & wOrderNo & " "
Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then
	If (CDbl(vRS("cnt")) > 0) Then
		hide = "Y"
	End If
End If


'--- �w�b�_�����̏���o��
vSQL = ""
vSQL = vSQL & "SELECT TOP 1 "
vSQL = vSQL & "      a.�󒍔ԍ� "
vSQL = vSQL & "    , a.�ڋq�ԍ� "
vSQL = vSQL & "    , a.���ϓ� "
vSQL = vSQL & "    , a.�󒍓� "
vSQL = vSQL & "    , a.�o�׊����� "
vSQL = vSQL & "    , a.�󒍌`�� "
vSQL = vSQL & "    , a.�x�����@ "
vSQL = vSQL & "    , a.���i���v���z "
vSQL = vSQL & "    , a.���� "
vSQL = vSQL & "    , a.����萔�� "
vSQL = vSQL & "    , a.�󒍍��v���z "
vSQL = vSQL & "    , a.���v���z "
vSQL = vSQL & "    , a.�O�ō��v���z "
vSQL = vSQL & "    , a.���p�|�C���g "
vSQL = vSQL & "    , a.�ꊇ�o�׃t���O "
vSQL = vSQL & "    , a.�̎������� "
vSQL = vSQL & "    , a.�̎����A������ "
vSQL = vSQL & "    , a.Web�󒍕ύX�J�n�� "
vSQL = vSQL & "    , a.����ŗ� "
vSQL = vSQL & "    , a.�^����ЃR�[�h "
vSQL = vSQL & "    , a.�S���҃R�[�h "
vSQL = vSQL & "    , a.�U�����`�l "
vSQL = vSQL & "    , a.�c�Ə��~�߃t���O "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.�󒍌`�� = '�M�t�g' THEN '' "
vSQL = vSQL & "     ELSE b.�������͐�X�֔ԍ� END "
vSQL = vSQL & "   ) AS �������͐�X�֔ԍ� "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.�󒍌`�� = '�M�t�g' THEN '' "
vSQL = vSQL & "     ELSE b.�������͐�s���{�� END "
vSQL = vSQL & "   ) AS �������͐�s���{�� "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.�󒍌`�� = '�M�t�g' THEN '' "
vSQL = vSQL & "     ELSE b.�������͐�Z�� END "
vSQL = vSQL & "   ) AS �������͐�Z�� "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.�󒍌`�� = '�M�t�g' THEN gift_c.�n���h���l�[�� "
vSQL = vSQL & "     ELSE b.�������͐於�O END "
vSQL = vSQL & "   ) AS �������͐於�O "
vSQL = vSQL & ", (CASE "
vSQL = vSQL & "     WHEN a.�󒍌`�� = '�M�t�g' THEN '' "
vSQL = vSQL & "     ELSE b.�������͐�d�b�ԍ� END "
vSQL = vSQL & "   ) AS �������͐�d�b�ԍ� "
vSQL = vSQL & "    , b.�������͐�X�֔ԍ� AS ORG_�������͐�X�֔ԍ� "
vSQL = vSQL & "    , b.�������͐�s���{�� AS ORG_�������͐�s���{�� "
vSQL = vSQL & "    , b.�������͐�Z�� AS ORG_�������͐�Z�� "
vSQL = vSQL & "    , b.�������͐於�O AS ORG_�������͐於�O "
vSQL = vSQL & "    , b.�������͐�d�b�ԍ� AS ORG_�������͐�d�b�ԍ� "
vSQL = vSQL & "    , b.�ŏI�w��[�� "
vSQL = vSQL & "    , b.�ŏI���Ԏw�� "
vSQL = vSQL & "    , a.�ߕs�����E���z " ' 2015.05.07 GV add
vSQL = vSQL & "    , a.�폜�� "
vSQL = vSQL & "    , a.Web�����ύX�L�����Z�����t���O "
vSQL = vSQL & "    , a.�x�����@�ڍ� " '2018.12.21 GV add
vSQL = vSQL & "    , a.�̎����ԍ� " '2020.02.05 GV add
vSQL = vSQL & "    , a.�̎������s�� " '2020.02.05 GV add
vSQL = vSQL & "    , (CASE WHEN a.�ŏI������ IS NULL THEN a.�󒍓� " '2020.03.18 GV add
vSQL = vSQL & "            ELSE a.�ŏI������ " '2020.03.18 GV add
vSQL = vSQL & "       END) AS �̎��� " '2020.03.18 GV add
vSQL = vSQL & "    , a.�M�t�g�ڋq�ԍ� "
vSQL = vSQL & "    , a.�M�t�g�ԍ� "

vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "��     a WITH (NOLOCK) "
vSQL = vSQL & " INNER JOIN " & gLinkServer & "�󒍖��� b WITH (NOLOCK) "
vSQL = vSQL & "   ON b.�󒍔ԍ� = a.�󒍔ԍ� "

vSQL = vSQL & " LEFT JOIN " & gLinkServer & "�ڋq gift_c WITH (NOLOCK) "
vSQL = vSQL & "   ON gift_c.�ڋq�ԍ� = a.�M�t�g�ڋq�ԍ� "

vSQL = vSQL & "WHERE "
If (wOrderGift = "N") Then
	vSQL = vSQL & "        a.�󒍔ԍ� = " & wOrderNo
	vSQL = vSQL & "    AND a.�ڋq�ԍ� = " & wCustomerNo & " "
Else
	vSQL = vSQL & "        a.�M�t�g�ԍ� = " & wOrderNo
	vSQL = vSQL & "    AND a.�M�t�g�ڋq�ԍ� = " & wCustomerNo & " "
End If

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
	one_time_todokesaki1 = vRS("�������͐�X�֔ԍ�") & "^" &_
		vRS("�������͐�s���{��") & "^" &_
		vRS("�������͐�Z��") & "^" &_
		vRS("�������͐於�O") & "^" &_
		vRS("�������͐�d�b�ԍ�")

	one_time_todokesaki1 = Replace(one_time_todokesaki1, """", "�h")

	' ORG_�������͐�
	one_time_todokesaki2 = vRS("ORG_�������͐�X�֔ԍ�") & "^" &_
		vRS("ORG_�������͐�s���{��") & "^" &_
		vRS("ORG_�������͐�Z��") & "^" &_
		vRS("ORG_�������͐於�O") & "^" &_
		vRS("ORG_�������͐�d�b�ԍ�")

	one_time_todokesaki2 = Replace(one_time_todokesaki2, """", "�h")


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

	' ���v���z
	If (IsNull(vRS("���v���z"))) Then
		totalOrderAmount2 = 0
	Else
		totalOrderAmount2 = CDbl(vRS("���v���z"))
	End If

	' ���p�|�C���g
	If (IsNull(vRS("���p�|�C���g"))) Then
		usedPoint = 0
	Else
		usedPoint = CDbl(vRS("���p�|�C���g"))
	End If

	' �U�����`�l
	If (IsNull(vRS("�U�����`�l"))) Then
		furikomiMeigi = ""
	Else
		furikomiMeigi = CStr(Trim(vRS("�U�����`�l")))
		furikomiMeigi = Replace(furikomiMeigi, """", "�h")
	End If

	'�c�Ə��~�߃t���O
	If (IsNull(vRS("�c�Ə��~�߃t���O"))) Then
		storeStop = ""
	Else
		storeStop = CStr(Trim(vRS("�c�Ə��~�߃t���O")))
	End If

	'Web�����ύX�L�����Z�����t���O
	If (IsNull(vRS("Web�����ύX�L�����Z�����t���O"))) Then
		webModCancelFlg = "N"
	Else
		If (Trim(vRS("Web�����ύX�L�����Z�����t���O")) <> "Y") Then
			webModCancelFlg = "N"
		Else
			webModCancelFlg = "Y"
		End If
	End If

	' �폜��
	If (IsNull(vRS("�폜��"))) Then
		deleteDate = ""
	Else
		deleteDate = CStr(Trim(vRS("�폜��")))
		webModCancelFlg = "N"
	End If

	' 2018.12.21 GV add start
	'�x�������@�ڍ�
	If (IsNull(vRS("�x�����@�ڍ�"))) Then
		paymentMethodDetail = ""
	Else
		paymentMethodDetail = CStr(vRS("�x�����@�ڍ�"))
	End If
	' 2018.12.21 GV add end

	'2020.02.05 GV add start
	'�̎������s�t���O
	receiptFlag = getReceiptFlag(vRS("�x�����@"), wOrderNo)

	'�̎������s��
	If (IsNull(vRS("�̎������s��"))) Then
		receiptDate = ""
	Else
		receiptDate = CStr(Trim(vRS("�̎������s��")))
	End If
	'2020.02.05 GV add end

	'2020.03.18 GV add start
	'�̎���
	If (IsNull(vRS("�̎���"))) Then
		displayReceiptDate = ""
	Else
		displayReceiptDate = CStr(Trim(vRS("�̎���")))
	End If
	'2020.03.18 GV add end

	' �M�t�g�ڋq�ԍ� 2021.06.30 GV add
	If (IsNull(vRS("�M�t�g�ڋq�ԍ�"))) Then
		giftCustomerNo = 0
	Else
		giftCustomerNo =CStr(vRS("�M�t�g�ڋq�ԍ�"))
	End If

	' �M�t�g�ԍ� 2021.06.30 GV add
	If (IsNull(vRS("�M�t�g�ԍ�"))) Then
		giftNo = 0
	Else
		giftNo = CStr(vRS("�M�t�g�ԍ�"))
	End If


	With oJSON.data("data")
		.Add "o_no", CStr(Trim(vRS("�󒍔ԍ�")))
		.Add "est_dt", estimateDate
		.Add "o_dt", orderDate
		.Add "ship_comp_dt", shippingDate
		.Add "o_type", CStr(Trim(vRS("�󒍌`��")))
		.Add "pay_method",  CStr(Trim(vRS("�x�����@")))
		.Add "pay_method_detail", paymentMethodDetail ' 2018.12.21 GV add
		.Add "furikomi_nm", furikomiMeigi
		.Add "tax_am", CDbl(Trim(vRS("�O�ō��v���z")))
		.Add "total_item_am", CDbl(Trim(vRS("���i���v���z")))
		.Add "ff_charge", CDbl(vRS("����")) 
		.Add "cod_charge", CDbl(vRS("����萔��"))
		.Add "kabusoku_am", CDbl(Trim(vRS("�ߕs�����E���z"))) ' 2015.05.07 GV add
		.Add "total_order_am", CDbl(vRS("�󒍍��v���z"))
		.Add "total_order_am2", totalOrderAmount2 ' ���v���z
		.Add "used_pt", usedPoint ' ���p�|�C���g
		.Add "comb_ship_flg", CStr(Trim(vRS("�ꊇ�o�׃t���O")))
		.Add "receipt_name", receiptName
		.Add "receipt_note", receiptNote
'		.Add "web_order_modify_start_date", vRS("Web�󒍕ύX�J�n��")
		.Add "tax_rate", CDbl(vRS("����ŗ�"))
		.Add "ff_cd", CStr(vRS("�^����ЃR�[�h"))
		'.Add "tantou_cd", CStr(vRS("�S���҃R�[�h"))
		.Add "one_time_todokesaki1", one_time_todokesaki1
		.Add "one_time_todokesaki2", one_time_todokesaki2
		.Add "nouki_dt", final_nouki_date_time
		.Add "store_stop", storeStop
		.Add "modifying", webModCancelFlg
		.Add "del_dt", deleteDate
		.Add "hide_ari", hide ' 2016.06.01 GV add
		.Add "receipt_flg", receiptFlag '2020.02.05 GV add
		.Add "receipt_no", CStr(Trim(vRS("�̎����ԍ�"))) '2020.02.05 GV add
		.Add "receipt_dt", receiptDate '2020.02.05 GV add
		.Add "display_receipt_dt", displayReceiptDate '2020.03.18 GV add
		.Add "gift_cst_no" , giftCustomerNo '�M�t�g�ڋq�ԍ� 2021.06.30 GV add
		.Add "gift_no" , giftNo '�M�t�g�ԍ� 2021.06.30 GV add
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
