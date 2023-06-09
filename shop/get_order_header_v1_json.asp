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
'	Emax�󒍁@�擾API
'
'
'�ύX����
'2016/03/29 GV �V�K�쐬
'2016.06.22 GV �F�؃A�V�X�g�g�p�t���O�Ή��B
'2016.09.06 GV �L�����Z�����̈������߂������̉��C�Ή��B
'2016.11.17 GV 3D�Z�L���A�Ή�
'
'========================================================================
'On Error Resume Next

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

Dim orderDate
Dim customerEmail
Dim estimateNote
Dim estimateDate
Dim shipSodFlg
Dim noukiDt
Dim noukiTm
Dim storeStopFlg
Dim furikomiMeigi
Dim ccTotalAm
Dim ccCreditNo
Dim ccSlipNo
Dim receiptFlg
Dim receiptName
Dim ritouFlg
Dim customerTel
Dim receiptNote
Dim adCd
Dim econNo
Dim econPayUrl
Dim econTranUrl
Dim orderTel
Dim orderFax
Dim usedPoint
Dim totalOrderAmount2
Dim depositAmount ' �������v���z
Dim depositFlag   ' ���������t���O
Dim coupon   ' �N�[�|��
Dim ccAssist '�F�؃A�V�X�g�t���O 2016.06.22 GV add
Dim cc3dSecure '3D�Z�L���A�t���O 2016.11.17 GV add
Dim cc3dSecureResult '3D�Z�L���A���ʃR�[�h 2016.11.17 GV add

Set oJSON = New aspJSON

'--- �w�b�_�����̏���o��
vSQL = ""
vSQL = vSQL & "SELECT TOP 1 "
vSQL = vSQL & "  o.�󒍔ԍ� "							' order_no
vSQL = vSQL & " ,o.�ڋq�ԍ� "							' customer_no
vSQL = vSQL & " ,o.�ڋqE_mail "							' email
vSQL = vSQL & " ,o.�x�����@ "							' payment_method
vSQL = vSQL & " ,o.�^����ЃR�[�h "						' freight_forwarder_cd
vSQL = vSQL & " ,o.���ϔ��l "							' estimate_note
vSQL = vSQL & " ,o.���i���v���z "						' total_item_amount
vSQL = vSQL & " ,o.���� "								' freight_charge
vSQL = vSQL & " ,o.����萔�� "							' daibiki_charge
vSQL = vSQL & " ,o.�O�ō��v���z "						' total_tax_amount
vSQL = vSQL & " ,o.�󒍍��v���z "						' total_order_amount
vSQL = vSQL & " ,o.�󒍓� "								' order_date
vSQL = vSQL & " ,o.���ϓ� "								' input_date
vSQL = vSQL & " ,od.�͐�Z���A�� "						' todokesaki_address_renban
vSQL = vSQL & " ,od.�������͐於�O "					' todokesaki_name
vSQL = vSQL & " ,od.�������͐�X�֔ԍ� "				' todokesaki_postal_cd
vSQL = vSQL & " ,od.�������͐�s���{�� "				' todokesaki_prefecture
vSQL = vSQL & " ,od.�������͐�Z�� "					' todokesaki_address
vSQL = vSQL & " ,od.�������͐�d�b�ԍ� "				' todokesaki_tel
vSQL = vSQL & " ,od.�������͐�[�i�����t�t���O "	' todokesaki_nouhinsho_send_flag
vSQL = vSQL & " ,od.�ŏI�w��[�� "						' nouki_date
vSQL = vSQL & " ,od.�ŏI���Ԏw�� "						' nouki_time
vSQL = vSQL & " ,o.�c�Ə��~�߃t���O "					' store_stop_flag
vSQL = vSQL & " ,o.�ꊇ�o�׃t���O "						' combined_shipping_flag
vSQL = vSQL & " ,o.�U�����`�l "							' furikomi_meigi
vSQL = vSQL & " ,cc.�J�[�h�x�����z "					' card_total_amount
vSQL = vSQL & " ,cc.�J�[�h�^�M�m�F�ԍ� "				' card_credit_no
vSQL = vSQL & " ,cc.�J�[�h�l�b�g�`�[�ԍ� "				' card_net_slip_no
vSQL = vSQL & " ,o.�̎������s�t���O "					' receipt_flag
vSQL = vSQL & " ,o.�̎������� "							' receipt_name
vSQL = vSQL & " ,od.�����t���O "						' ritou_flag
vSQL = vSQL & " ,o.�����ғd�b�ԍ� "						' customer_tel
vSQL = vSQL & " ,o.�̎����A������ "						' receipt_note
vSQL = vSQL & " ,o.�L���R�[�h "							' ad_cd
vSQL = vSQL & " ,o.����ŗ� "							' tax_rate
vSQL = vSQL & " ,o.eContext��t�ԍ� "					' e_context_no
vSQL = vSQL & " ,o.eContext�x�����@URL "				' e_context_payment_method_url
vSQL = vSQL & " ,o.eContext�U���[URL "					' e_context_transfer_url
vSQL = vSQL & " ,o.�󒍌`�� "							' order_type
vSQL = vSQL & " ,o.�ߕs�����E���z "						' kabusoku_sousai_amount
vSQL = vSQL & " ,o.�����Җ��O "							' order_name
vSQL = vSQL & " ,o.�����җX�֔ԍ� "						' order_postal_cd
vSQL = vSQL & " ,o.�����ғs���{�� "						' order_prefecture
vSQL = vSQL & " ,o.�����ҏZ�� "							' order_address
vSQL = vSQL & " ,o.�����ғd�b�ԍ� "						' order_tel
vSQL = vSQL & " ,o.������FAX "							' order_fax
vSQL = vSQL & " ,o.���p�|�C���g "						' used_point
vSQL = vSQL & " ,o.���v���z "							' total_used_point_order_amount
vSQL = vSQL & " ,o.�������v���z "						' work_order.old_deposit_amount
vSQL = vSQL & " ,o.���������t���O "						' work_order.old_deposit_flag
vSQL = vSQL & " ,o.�N�[�|�� "							' coupon
vSQL = vSQL & " ,o.�F�؃A�V�X�g�t���O "					' cc_assist_flag 2016.06.22 GV add
vSQL = vSQL & " ,o.�Z�L���A3D�t���O "					' cc_3d_secure_flag 2016.11.17 GV add
vSQL = vSQL & " ,o.�Z�L���A3D���ʃR�[�h "				' cc_3d_secure_result_cd 2016.11.17 GV add


vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "��     o WITH (NOLOCK) "
vSQL = vSQL & "INNER JOIN " & gLinkServer & "�󒍖��� od WITH (NOLOCK) "
vSQL = vSQL & "   ON od.�󒍔ԍ� = o.�󒍔ԍ� "
vSQL = vSQL & "  AND od.�󒍐��� > 0 "

vSQL = vSQL & "LEFT JOIN " & gLinkServer & "�󒍃J�[�h��� cc WITH (NOLOCK) "
vSQL = vSQL & "  ON cc.�󒍔ԍ� = o.�󒍔ԍ� "

vSQL = vSQL & "WHERE "
vSQL = vSQL & "      o.�󒍔ԍ� = " & wOrderNo & " "
vSQL = vSQL & "  AND o.�ڋq�ԍ� = " & wUserID & " "

vSQL = vSQL & " ORDER BY "
vSQL = vSQL & "        od.�󒍖��הԍ� ASC "


'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' ���X�g�ǉ�
	oJSON.data.Add "data" ,oJSON.Collection()

	' --------------------
	' �󒍓�
	If (IsNull(vRS("�󒍓�"))) Then
		orderDate = ""
	Else
		orderDate = CStr(Trim(vRS("�󒍓�")))
	End If

	'�ڋqE_mail
	If (IsNull(vRS("�ڋqE_mail"))) Then
		customerEmail = ""
	Else
		customerEmail = CStr(Trim(vRS("�ڋqE_mail")))
	End If

	'���ϔ��l
	If (IsNull(vRS("���ϔ��l"))) Then
		estimateNote = ""
	Else
		estimateNote = CStr(Trim(vRS("���ϔ��l")))
	End If

	' ���ϓ�
	If (IsNull(vRS("���ϓ�"))) Then
		estimateDate = ""
	Else
		estimateDate = CStr(Trim(vRS("���ϓ�")))
	End If

	'�������͐�[�i�����t�t���O
	If (IsNull(vRS("�������͐�[�i�����t�t���O"))) Then
		shipSodFlg = ""
	Else
		shipSodFlg = CStr(Trim(vRS("�������͐�[�i�����t�t���O")))
	End If

	' �z����
	If (IsNull(vRS("�ŏI�w��[��"))) Then
		noukiDt = ""
	Else
		noukiDt = CStr(Trim(vRS("�ŏI�w��[��")))
	End If

	' �z�����ԑ�
	If (IsNull(vRS("�ŏI�w��[��"))) Then
		noukiTm = ""
	Else
		noukiTm = CStr(Trim(vRS("�ŏI���Ԏw��"))) 
	End If

	'�c�Ə��~�߃t���O
	If (IsNull(vRS("�c�Ə��~�߃t���O"))) Then
		storeStopFlg = ""
	Else
		storeStopFlg = CStr(Trim(vRS("�c�Ə��~�߃t���O")))
	End If

	' �U�����`�l
	If (IsNull(vRS("�U�����`�l"))) Then
		furikomiMeigi = ""
	Else
		furikomiMeigi = CStr(Trim(vRS("�U�����`�l")))
		'furikomiMeigi = Replace(furikomiMeigi, """", "�h")
	End If

	' �J�[�h�x�����z
	If (IsNull(vRS("�J�[�h�x�����z"))) Then
		ccTotalAm = ""
	Else
		ccTotalAm = CStr(CDbl(vRS("�J�[�h�x�����z")))
	End If

	'�J�[�h�^�M�m�F�ԍ�
	If (IsNull(vRS("�J�[�h�^�M�m�F�ԍ�"))) Then
		ccCreditNo = ""
	Else
		ccCreditNo = CStr(Trim(vRS("�J�[�h�^�M�m�F�ԍ�")))
	End If

	'�J�[�h�l�b�g�`�[�ԍ�
	If (IsNull(vRS("�J�[�h�l�b�g�`�[�ԍ�"))) Then
		ccSlipNo = ""
	Else
		ccSlipNo = CStr(Trim(vRS("�J�[�h�l�b�g�`�[�ԍ�")))
	End If

	'�̎������s�t���O
	If (IsNull(vRS("�̎������s�t���O"))) Then
		receiptFlg = ""
	Else
		receiptFlg = CStr(Trim(vRS("�̎������s�t���O")))
	End If

	' �̎�������
	If (IsNull(vRS("�̎�������"))) Then
		receiptName = ""
	Else
		receiptName = CStr(Trim(vRS("�̎�������")))
		receiptName = Replace(receiptName, """", "�h")
	End If

	'�����t���O
	If (IsNull(vRS("�����t���O"))) Then
		ritouFlg = ""
	Else
		ritouFlg = CStr(Trim(vRS("�����t���O")))
	End If

	'�����ғd�b�ԍ�
	If (IsNull(vRS("�����ғd�b�ԍ�"))) Then
		customerTel = ""
	Else
		customerTel = CStr(Trim(vRS("�����ғd�b�ԍ�")))
	End If

	' �̎����A������
	If (IsNull(vRS("�̎����A������"))) Then
		receiptNote = ""
	Else
		receiptNote = CStr(Trim(vRS("�̎����A������")))
		receiptNote = Replace(receiptNote, """", "�h")
	End If

	'�L���R�[�h
	If (IsNull(vRS("�L���R�[�h"))) Then
		adCd = ""
	Else
		adCd = CStr(Trim(vRS("�L���R�[�h")))
	End If

	'eContext��t�ԍ�
	If (IsNull(vRS("eContext��t�ԍ�"))) Then
		econNo = ""
	Else
		econNo = CStr(Trim(vRS("eContext��t�ԍ�")))
	End If

	'eContext�x�����@URL
	If (IsNull(vRS("eContext�x�����@URL"))) Then
		econPayUrl = ""
	Else
		econPayUrl = CStr(Trim(vRS("eContext�x�����@URL")))
	End If

	'eContext�U���[URL
	If (IsNull(vRS("eContext�U���[URL"))) Then
		econTranUrl = ""
	Else
		econTranUrl = CStr(Trim(vRS("eContext�U���[URL")))
	End If

	'�����ғd�b�ԍ�
	If (IsNull(vRS("�����ғd�b�ԍ�"))) Then
		orderTel = ""
	Else
		orderTel = CStr(Trim(vRS("�����ғd�b�ԍ�")))
	End If

	'������FAX
	If (IsNull(vRS("������FAX"))) Then
		orderFax = ""
	Else
		orderFax = CStr(Trim(vRS("������FAX")))
	End If

	' ���p�|�C���g
	If (IsNull(vRS("���p�|�C���g"))) Then
		usedPoint = 0
	Else
		usedPoint = CDbl(vRS("���p�|�C���g"))
	End If

	' ���v���z
	If (IsNull(vRS("���v���z"))) Then
		totalOrderAmount2 = 0
	Else
		totalOrderAmount2 = CDbl(vRS("���v���z"))
	End If

	'���������t���O
	If (IsNull(vRS("���������t���O"))) Then
		depositFlag = ""
	Else
		depositFlag = CStr(Trim(vRS("���������t���O")))
	End If

	' �������v���z
	If (IsNull(vRS("�������v���z"))) Then
		depositAmount = 0
	Else
		depositAmount = CDbl(vRS("�������v���z"))
	End If

	'�N�[�|��
	If (IsNull(vRS("�N�[�|��"))) Then
		coupon = ""
	Else
		coupon = CStr(Trim(vRS("�N�[�|��")))
		coupon = Replace(coupon, """", "�h")
	End If

	'2016.06.22 GV add start
	'�F�؃A�V�X�g�t���O
	If (IsNull(vRS("�F�؃A�V�X�g�t���O"))) Then
		ccAssist = ""
	Else
		ccAssist = CStr(Trim(vRS("�F�؃A�V�X�g�t���O")))
	End If
	'2016.06.22 GV add end

	'2016.11.17 GV add start
	'3D�Z�L���A�t���O
	If (IsNull(vRS("�Z�L���A3D�t���O"))) Then
		cc3dSecure = ""
	Else
		cc3dSecure = CStr(Trim(vRS("�Z�L���A3D�t���O")))
	End If

	'3D�Z�L���A���ʃR�[�h
	If (IsNull(vRS("�Z�L���A3D���ʃR�[�h"))) Then
		cc3dSecureResult = ""
	Else
		cc3dSecureResult = CStr(Trim(vRS("�Z�L���A3D���ʃR�[�h")))
	End If
	'2016.11.17 GV add end

	With oJSON.data("data")
		.Add "o_no", CStr(Trim(vRS("�󒍔ԍ�")))
		.Add "cstm_mail", customerEmail
		.Add "pay_method",  CStr(Trim(vRS("�x�����@")))
		.Add "ff_cd", CStr(vRS("�^����ЃR�[�h"))
		.Add "est_nt", estimateNote
		.Add "total_item_am", CDbl(Trim(vRS("���i���v���z")))
		.Add "ff_charge", CDbl(vRS("����")) 
		.Add "cod_charge", CDbl(vRS("����萔��"))
		.Add "tax_am", CDbl(vRS("�O�ō��v���z"))
		.Add "total_order_am", CDbl(vRS("�󒍍��v���z"))
		.Add "est_dt", estimateDate '���ϓ�(input_date)
		.Add "ship_addr_no", CDbl(vRS("�͐�Z���A��"))
		.Add "ship_sod_flg", shipSodFlg '�������͐�[�i�����t�t���O(todokesaki_nouhinsho_send_flag)
		.Add "nouki_dt", noukiDt '�ŏI�w��[��(nouki_date)
		.Add "nouki_tm", noukiTm '�ŏI���Ԏw��(nouki_time)
		.Add "store_stop", storeStopFlg '�c�Ə��~�߃t���O(store_stop_flag)
		.Add "comb_ship_flg", CStr(Trim(vRS("�ꊇ�o�׃t���O"))) ' combined_shipping_flag
		.Add "furikomi_nm", furikomiMeigi '�U�����`�l(furikomi_meigi)
		.Add "cc_pay_am", ccTotalAm '�J�[�h�x�����z(card_total_amount)
		.Add "cc_c_no", ccCreditNo '�J�[�h�^�M�m�F�ԍ�(card_credit_no)
		.Add "cc_slip", ccSlipNo '�J�[�h�l�b�g�`�[�ԍ�(card_net_slip_no)
		.Add "receipt_flg", receiptFlg '�̎������s�t���O(receipt_flag)
		.Add "receipt_nm", receiptName '�̎�������(receipt_name)
		.Add "ritou", ritouFlg '�����t���O(ritou_flag)
		.Add "cstm_tel", customerTel '�����ғd�b�ԍ�(customer_tel)
		.Add "receipt_nt", receiptNote '�̎����A������(receipt_note)
		.Add "ad_cd", adCd '�L���R�[�h(ad_cd)
		.Add "tax_rate", CDbl(vRS("����ŗ�"))
		.Add "econ_no", econNo 'eContext��t�ԍ�(e_context_no)
		.Add "econ_pay", econPayUrl 'eContext�x�����@URL(e_context_payment_method_url)
		.Add "econ_tran", econTranUrl 'eContext�U���[URL(e_context_transfer_url)
		.Add "o_type", CStr(Trim(vRS("�󒍌`��")))
		.Add "kabusoku_am", CDbl(Trim(vRS("�ߕs�����E���z"))) ' 2015.05.07 GV add
		.Add "o_nm", CStr(Trim(vRS("�����Җ��O")))
		.Add "o_zip", CStr(Trim(vRS("�����җX�֔ԍ�")))
		.Add "o_pref", CStr(Trim(vRS("�����ғs���{��")))
		.Add "o_addr", CStr(Trim(vRS("�����ҏZ��")))
		.Add "o_tel", orderTel '�����ғd�b�ԍ�(order_tel)
		.Add "o_fax", orderFax '������FAX(order_fax)
		.Add "used_pt", usedPoint ' ���p�|�C���g
		.Add "total_order_am2", totalOrderAmount2 ' ���v���z(total_used_point_order_amount)
		.Add "deposit_flg", depositFlag ' ���������t���O(work_order.old_deposit_flag)
		.Add "deposit_am", depositAmount ' �������v���z(work_order.old_deposit_amount)
		.Add "coupon", coupon ' �N�[�|��(work_order.coupon)
		.Add "cc_assist", ccAssist ' �F�؃A�V�X�g�t���O(work_order.old_cc_assist_flag) 2016.06.22 GV add
		.Add "o_dt", orderDate '�󒍓� 2016.09.06 GV add
		.Add "cc_3d_secure", cc3dSecure ' 3D�Z�L���A�t���O 2016.11.17 GV add
		.Add "cc_3d_secure_result", cc3dSecureResult ' 3D�Z�L���A���ʃR�[�h 2016.11.17 GV add
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
Response.AddHeader "X-Content-Type-Options", "nosniff"

' JSON�f�[�^�̏o��
Response.Write oJSON.JSONoutput()

End Function
%>
