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
<!--#include file="../common/HttpsSecurity.inc"-->
<%
'========================================================================
'
'	���͂���A���x�������@�̑I��
'
'2012/06/14 ok �f�U�C���ύX�̂��ߋ��ł����ɐV�K�쐬
'2012/08/08 nt ���d�ʕi���̉�ʐ����ǉ�
'2012/08/25 nt ���������֎~�n��̐���@�\��ǉ�
'2012/09/05 ok �����w��\���ύX
'2012/10/25 ok �U���ȊO�Łu�̎����K�v�v�N���b�N�Ń|�b�v�A�b�v�\���ǉ�
'2014/08/05 GV �[�i���\���ύX�Ή�
'2014/08/20 GV �̎������K�v�ȏꍇ�ɔ[�i���s�v��I���ł��Ȃ��悤�C��
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

'---- Session���
Dim wUserID
Dim wMsg
Dim wErrMsg

'---- �󂯓n����������ϐ�

'---- Web�ڋq���
Dim wCustomerNm
Dim wCustomerKn
Dim wCustomerEmail
Dim wCustomerKabusokuAm
Dim wCustomerClass
Dim wCustomerZip
Dim wCustomerPref
Dim wCustomerAddress
Dim wCustomerTel
Dim wCustomerRitouFl   '2012/08/08 nt add
Dim wCustomerSagawaLTFl '2012/08/25 nt add

'---- ����
Dim wPaymentMethod
Dim wShipAddressNo
Dim wFurikomiNm
Dim wShipInvoiceFl
Dim wFreightForwarder
Dim wIkkatsuFl
Dim wDeliveryMM
Dim wDeliveryDD
Dim wDeliveryTM
Dim wEigyoushoDomeFl
Dim wReceiptFl
Dim wReceiptNm
Dim wReceiptMemo
Dim wToriyoseFl
Dim wTokuchuuFl
Dim wDaibikiFukaFl
Dim wRebateFl
Dim wKuyuKinshiFl   '2012/08/08 nt add
Dim wSagawaLTFl     '2012/08/25 nt add

Dim wNoData
Dim wShipAddressHTML
Dim wErrDesc   '2011/08/01 an add

'---- ���͂��惊�X�g
Dim wAddressNoHTML					'�Z���A��
Dim wZipHTML						'�X�֔ԍ�
Dim wAddressHTML					'�Z��
Dim wTelephoneNoHTML				'�d�b�ԍ�
Dim wAddressNameHTML				'���͂��掁��
Dim wRitouFlHTML					'�����t���O                     2012/08/08 nt add
Dim wKuyuKinshiFlHTML				'��A�֎~�t���O�F���d�ʕi�t���O 2012/08/08 nt add
Dim wSagawaLTHTML					'���쐧���t���O                 2012/08/25 nt add

Dim wShowInvoice					'2014/08/05 GV add
Dim wInvoiceDisabled				'2014/08/05 GV add
Dim wShipAddressNo1Data				'2014/08/05 GV add
Dim wSelectedShipAddressData		'2014/08/05 GV add
Dim wInvoiceChecked					'2014/08/05 GV add


'---- �z�����Ԏw��    2011/06/29 an add
Dim wDeliveryTime01
Dim wDeliveryTime02
Dim wDeliveryTime03
Dim wDeliveryTime04
Dim wDeliveryTime05

'---- DB
Dim Connection

'=======================================================================
'	�󂯓n�������o��
'=======================================================================
'---- Session�ϐ�
wUserID = Session("userID")
wMsg = Session.contents("msg")

'---- �󂯓n�������o��

Session("msg") = ""

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()
Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderinfoEnter.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection= Nothing    '2011/08/01 an add

End Function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

wNoData = False

wInvoiceChecked = array("", "", "")		'2014/08/05 GV add

Call get_customer()				'�ڋq���̎��o��
Call get_order()				'���󒍏��̎��o��
Call get_todokesaki()			'�ڋq�͐���̎��o��
Call get_DeliveryTime()			'�z�����ԑт��R���g���[���}�X�^������o��

End Function

'========================================================================
'
'	Function	�ڋq���̎��o��
'
'========================================================================
Function get_customer()

Dim RSv
Dim vSQL
Dim bobj

'---- �ڋq�����o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.�ڋq�� "
vSQL = vSQL & "    , a.�ڋq�t���K�i "
vSQL = vSQL & "    , a.�U�����`�l "
vSQL = vSQL & "    , a.�ڋqE_mail1 "
vSQL = vSQL & "    , a.�����ߕs�����z "
vSQL = vSQL & "    , a.�ڋq�N���X "
vSQL = vSQL & "    , b.�ڋq�X�֔ԍ� "
vSQL = vSQL & "    , b.�ڋq�s���{�� "
vSQL = vSQL & "    , b.�ڋq�Z�� "
vSQL = vSQL & "    , c.�ڋq�d�b�ԍ� "
vSQL = vSQL & "    , CASE WHEN d.�X�֔ԍ� IS NOT NULL THEN 'Y' ELSE 'N' END AS �����t���O "		'2012/08/08 nt add
vSQL = vSQL & "    , CASE WHEN e.�X�֔ԍ� IS NOT NULL THEN 'Y' ELSE 'N' END AS ���쐧���t���O "	'2012/08/25 nt add
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web�ڋq                          a WITH (NOLOCK) "
vSQL = vSQL & "      INNER JOIN Web�ڋq�Z��         b WITH (NOLOCK) "
vSQL = vSQL & "        ON     b.�ڋq�ԍ� = a.�ڋq�ԍ� "
vSQL = vSQL & "           AND b.�Z���A�� = 1 "
vSQL = vSQL & "      INNER JOIN Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK) "
vSQL = vSQL & "        ON     c.�ڋq�ԍ� = a.�ڋq�ԍ� "
vSQL = vSQL & "           AND c.�Z���A�� = b.�Z���A�� "
vSQL = vSQL & "           AND c.�d�b�A�� = 1 "
vSQL = vSQL & "      LEFT  JOIN ( SELECT '�Z��' AS 'AddrTypeHouse' ) t1 "
vSQL = vSQL & "        ON     b.�Z���敪 = t1.AddrTypeHouse "
'2012/08/08 nt add Start
vSQL = vSQL & "      LEFT  JOIN ���� d "
vSQL = vSQL & "        ON     REPLACE(b.�ڋq�X�֔ԍ�, '-', '') = d.�X�֔ԍ� "
'2012/08/08 nt add End
'2012/08/25 nt add Start
vSQL = vSQL & "      LEFT  JOIN ���쐧�� e "
vSQL = vSQL & "        ON     REPLACE(b.�ڋq�X�֔ԍ�, '-', '') = e.�X�֔ԍ� AND e.����s�t���O='Y' "
'2012/08/25 nt add End
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        t1.AddrTypeHouse IS NOT NULL "
vSQL = vSQL & "    AND a.�ڋq�ԍ� = " & wUserID & " "

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	wErrMsg = "�ڋq��񂪂���܂���B"
Else
	wCustomerNm = RSv("�ڋq��")
	If RSv("�U�����`�l") <> "" Then
		wFurikomiNm = RSv("�U�����`�l")
	Else
		wFurikomiNm = RSv("�ڋq�t���K�i")
	End If

	'---- ���p��S�p�ɕϊ�		'2011/09/09 hn add
	Set bobj = Server.CreateObject("basp21")
	wFurikomiNm = bobj.StrConv(wFurikomiNm,4)

	wCustomerEmail = RSv("�ڋqE_mail1")
	wCustomerKabusokuAm = RSv("�����ߕs�����z")
	wCustomerClass = RSv("�ڋq�N���X")
	wCustomerZip = RSv("�ڋq�X�֔ԍ�")
	wCustomerPref = RSv("�ڋq�s���{��")
	wCustomerAddress = RSv("�ڋq�Z��")
	wCustomerTel = RSv("�ڋq�d�b�ԍ�")

	'2012/08/08 nt add Start
	wCustomerRitouFl =  RSv("�����t���O")
	'2012/08/08 nt add End
	'2012/08/25 nt add Start
	wCustomerSagawaLTFl =  RSv("���쐧���t���O")
	'2012/08/25 nt add End
End If

RSv.Close

End Function

'========================================================================
'
'	Function	�󒍏��̎��o��
'
'========================================================================
Function get_order()

Dim RSv
Dim vSQL

'----���󒍃f�[�^���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.�x�����@ "
vSQL = vSQL & "    , a.�U�����`�l "
vSQL = vSQL & "    , a.�͐�Z���A�� "
vSQL = vSQL & "    , a.�͐於�O "
vSQL = vSQL & "    , a.�͐�X�֔ԍ� "
vSQL = vSQL & "    , a.�͐�s���{�� "
vSQL = vSQL & "    , a.�͐�Z�� "
vSQL = vSQL & "    , a.�͐�d�b�ԍ� "
vSQL = vSQL & "    , a.�͐�[�i�����t�t���O "
vSQL = vSQL & "    , a.�^����ЃR�[�h "
vSQL = vSQL & "    , a.�w��[�� "
vSQL = vSQL & "    , a.���Ԏw�� "
vSQL = vSQL & "    , a.�c�Ə��~�߃t���O "
vSQL = vSQL & "    , a.�ꊇ�o�׃t���O "
vSQL = vSQL & "    , a.�̎������s�t���O "
vSQL = vSQL & "    , a.�̎������� "
vSQL = vSQL & "    , a.�̎����A������ "
vSQL = vSQL & "    , a.���x�[�g�g�p�t���O "
vSQL = vSQL & "    , c.�����\���� "
vSQL = vSQL & "    , d.���[�J�[�������敪 "
vSQL = vSQL & "    , d.����s�t���O "
'2012/08/08 nt add Start
vSQL = vSQL & "    , d.��A�֎~�t���O "
'2012/08/08 nt add End
vSQL = vSQL & "FROM "
vSQL = vSQL & "    ����                       AS a WITH (NOLOCK) "
vSQL = vSQL & "      INNER JOIN ���󒍖���      AS b WITH (NOLOCK) "
vSQL = vSQL & "        ON     b.SessionID      = a.SessionID "
vSQL = vSQL & "      INNER JOIN Web�F�K�i�ʍ݌� AS c WITH (NOLOCK) "
vSQL = vSQL & "        ON     c.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
vSQL = vSQL & "           AND c.���i�R�[�h     = b.���i�R�[�h "
vSQL = vSQL & "           AND c.�F             = b.�F "
vSQL = vSQL & "           AND c.�K�i           = b.�K�i "
vSQL = vSQL & "      INNER JOIN Web���i         AS d WITH (NOLOCK) "
vSQL = vSQL & "        ON     d.���[�J�[�R�[�h = c.���[�J�[�R�[�h "
vSQL = vSQL & "           AND d.���i�R�[�h     = c.���i�R�[�h "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.SessionID = '" & gSessionID & "' "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "      b.�󒍖��הԍ� "

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = False Then

	'---- �w�b�_���Z�b�g
	wShipAddressNo = RSv("�͐�Z���A��")
	If isNumeric(wShipAddressNo) = False Then
		wShipAddressNo = 1
	ElseIf wShipAddressNo <= 0 Then
		wShipAddressNo = 1
	End If

	wShipInvoiceFl = RSv("�͐�[�i�����t�t���O")
	wPaymentMethod = RSv("�x�����@")

	'---- ���󒍂ɐU�����`�l��񂪂���Ώ㏑��   '2011/09/09 an mod s
	if RSv("�U�����`�l") <> "" then
		wFurikomiNm = RSv("�U�����`�l")
	end if                                       '2011/09/09 an mod e

	wIkkatsuFl = RSv("�ꊇ�o�׃t���O")

	wFreightForwarder = RSv("�^����ЃR�[�h")
	If wFreightForwarder = "" Then
		wFreightForwarder = "5"		'���Z �����l  '2011/06/29 an mod
	End If

	If isNull(RSv("�w��[��")) = False Then
		wDeliveryMM = cf_NumToChar(DatePart("m", RSv("�w��[��")),2)
		wDeliveryDD = cf_NumToChar(DatePart("d", RSv("�w��[��")),2)
	End If

	wDeliveryTM = RSv("���Ԏw��")

	wEigyoushoDomeFl = RSv("�c�Ə��~�߃t���O")

	wReceiptFl = RSv("�̎������s�t���O")
	wReceiptNm = RSv("�̎�������")
	wReceiptMemo = RSv("�̎����A������")
	If wReceiptFl = "Y" Then
		If wReceiptNm = "" Then
			wReceiptNm = wCustomerNm
		End If
		If wReceiptMemo = "" Then
			wReceiptMemo = "�����@���Ƃ���"
		End If
	End If

	wRebateFl = RSv("���x�[�g�g�p�t���O")

	wToriyoseFl = "N"
	wTokuchuuFl = "N"
	wDaibikiFukaFl = "N"

	'Do While RSv.EOF		'2011/03/04 na del
	Do Until RSv.EOF	'2011/03/04 na mod

		If RSv("�����\����") <= 0 Then				'�v����
			wToriyoseFl = "Y"
		End If
		If RSv("���[�J�[�������敪") = "����" Then	'���ʒ���
			wToriyoseFl = "Y"
			wTokuchuuFl = "Y"
		End If
		If RSv("����s�t���O") = "Y" Then				'������s��
			wDaibikiFukaFl = "Y"
		End If

		'2012/08/08 nt add Start
		'---- ��ʐ�������Z�b�g�i��A�֎~�t���O�F���d�ʕi�t���O�j
		If RSv("��A�֎~�t���O") = "Y" Then
			wKuyuKinshiFl = "Y"
		End If
		'2012/08/08 nt add End

		RSv.MoveNext

	Loop

Else

	wNoData = True

End If

RSv.Close

End Function

'========================================================================
'
'	Function	�ڋq�͐���̎��o��
'
'	Note		���͂���I��p�̃h���b�v�_�E�����X�g�𐶐�
'
'========================================================================
Function get_todokesaki()

Dim RSv
Dim vSQL

'---- �ڋq�͐�����o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "    a.�Z���A�� "
vSQL = vSQL & "  , a.�Z������ "
vSQL = vSQL & "  , a.�ڋq�X�֔ԍ� "
vSQL = vSQL & "  , a.�ڋq�s���{�� "
vSQL = vSQL & "  , a.�ڋq�Z�� "
vSQL = vSQL & "  , b.�ڋq�d�b�ԍ� "
vSQL = vSQL & "  , CASE WHEN c.�X�֔ԍ� IS NOT NULL THEN 'Y' ELSE 'N' END AS �����t���O "		'2012/08/08 nt add
vSQL = vSQL & "  , CASE WHEN d.�X�֔ԍ� IS NOT NULL THEN 'Y' ELSE 'N' END AS ���쐧���t���O "	'2012/08/25 nt add
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web�ڋq�Z��                      a WITH (NOLOCK) "
vSQL = vSQL & "      INNER JOIN Web�ڋq�Z���d�b�ԍ� b WITH (NOLOCK) "
vSQL = vSQL & "        ON     b.�ڋq�ԍ� = a.�ڋq�ԍ� "
vSQL = vSQL & "           AND b.�Z���A�� = a.�Z���A�� "
vSQL = vSQL & "      LEFT  JOIN ( SELECT '�d�b' AS 'PhoneTypeTel' ) t1 "
vSQL = vSQL & "        ON     b.�d�b�敪 = t1.PhoneTypeTel "
'2012/08/08 nt add Start
vSQL = vSQL & "      LEFT  JOIN ���� c "
vSQL = vSQL & "        ON     REPLACE(a.�ڋq�X�֔ԍ�, '-', '') = c.�X�֔ԍ� "
'2012/08/08 nt add End
'2012/08/25 nt add Start
vSQL = vSQL & "      LEFT  JOIN ���쐧�� d "
vSQL = vSQL & "        ON     REPLACE(a.�ڋq�X�֔ԍ�, '-', '') = d.�X�֔ԍ� AND d.����s�t���O='Y' "
'2012/08/25 nt add End
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        t1.PhoneTypeTel IS NOT NULL "
vSQL = vSQL & "    AND a.�폜�� IS NULL "
vSQL = vSQL & "    AND a.�ڋq�ԍ� = " & wUserID & " "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "    a.�Z���A�� "

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

wAddressNoHTML = "0"						'�Z���A��
wZipHTML = "''"								'�X�֔ԍ�
wAddressHTML = "''"							'�Z��
wTelephoneNoHTML = "''"						'�d�b�ԍ�
wAddressNameHTML = "'���͂����ύX����'"	'���͂��掁��

wRitouFlHTML = "''"							'�����t���O 2012/08/08 nt add
wKuyuKinshiFlHTML = "''"					'��A�֎~�t���O�F���d�ʕi�t���O 2012/08/08 nt add
wSagawaLTHTML = "''"						'���쐧���t���O 2012/08/25 nt add

wShipAddressHTML = "<option value=""0"">���͂����ύX����</option>"

Do While RSv.EOF = False
	'2014/08/05 GV add start
	If RSv("�Z���A��") = 1 Then
		wShipAddressNo1Data = RSv("�ڋq�s���{��") & RSv("�ڋq�Z��")
	End If
	
	If RSv("�Z���A��") = wShipAddressNo Then
		wSelectedShipAddressData =  RSv("�ڋq�s���{��") & RSv("�ڋq�Z��")
	End If
	'2014/08/05 GV add end

	wShipAddressHTML = wShipAddressHTML & _
					   "<option value=""" & RSv("�Z���A��") & """>" & _
						   RSv("�Z������") & _
						   " ��" & RSv("�ڋq�X�֔ԍ�") & _
						   " " & RSv("�ڋq�s���{��") & RSv("�ڋq�Z��") & _
						   " Tel. " & RSv("�ڋq�d�b�ԍ�") & _
						   "</option>" & vbNewLine

	' JavaScript�p�̂��͂����񃊃X�g���쐬
	wAddressNoHTML = wAddressNoHTML & "," & Replace(Replace(RSv("�Z���A��"),vbCR,""),vbLF,"") & ""
	wZipHTML = wZipHTML & ",'��" & Replace(Replace(RSv("�ڋq�X�֔ԍ�"),vbCR,""),vbLF,"") & "'"
	wAddressHTML = wAddressHTML & ",'" & Replace(Replace(RSv("�ڋq�s���{��") & RSv("�ڋq�Z��"),vbCR,""),vbLF,"") & "'"
	wTelephoneNoHTML = wTelephoneNoHTML & ",'" & Replace(Replace(RSv("�ڋq�d�b�ԍ�"),vbCR,""),vbLF,"") & "'"
	wAddressNameHTML = wAddressNameHTML & ",'" & Replace(Replace(RSv("�Z������"),vbCR,""),vbLF,"") & "'"
	wRitouFlHTML = wRitouFlHTML & ",'" & Replace(Replace(RSv("�����t���O"),vbCR,""),vbLF,"") & "'"			'2012/08/08 nt add
	wKuyuKinshiFlHTML = wKuyuKinshiFlHTML & ",'" & Replace(Replace(wKuyuKinshiFl,vbCR,""),vbLF,"") & "'"	'2012/08/08 nt add
	wSagawaLTHTML = wSagawaLTHTML & ",'" & Replace(Replace(RSv("���쐧���t���O"),vbCR,""),vbLF,"") & "'"	'2012/08/25 nt add

	RSv.MoveNext

Loop

RSv.Close

wShipAddressHTML = "<select name=""select_ship_address_no"" id=""select_ship_address_no"" onChange=""changeShipAddress();"">" & vbNewLine & _
				   wShipAddressHTML & _
				   "</select>"

End Function

'========================================================================
'
'	Function	�z�����Ԏw����擾
'
'========================================================================
Function get_DeliveryTime()

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'---- ���Z���Ԏw��01
call getCntlMst("��","���Ԏw��_���Z","01", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime01 = vItemChar1
'---- ���Z���Ԏw��02
call getCntlMst("��","���Ԏw��_���Z","02", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime02 = vItemChar1
'---- ���Z���Ԏw��03
call getCntlMst("��","���Ԏw��_���Z","03", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime03 = vItemChar1
'---- ���Z���Ԏw��04
call getCntlMst("��","���Ԏw��_���Z","04", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime04 = vItemChar1
'---- ���Z���Ԏw��05
call getCntlMst("��","���Ԏw��_���Z","05", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wDeliveryTime05 = vItemChar1

End Function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>���͂���A���x�������@�̑I���b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css?20120629" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120831" type="text/css">
<link rel="stylesheet" href="style/jquery.fancybox-1.3.4.css" type="text/css">
<script type="text/javascript">
//=====================================================================
//	OrderInfoInsert.asp��Submit
//=====================================================================
function OrderSubmit(pCmd){

	document.f_data.cmd.value = pCmd;
	document.f_data.action = "OrderInfoInsert.asp";
	document.f_data.submit();

}

//=====================================================================
//	���W�I�{�^���A�h���b�v�_�E�����X�g���ȑO�ɑI��������Ԃɂ���
//=====================================================================
function preset_values(){

	// �x�����@
	for (var i=0; i<document.f_data.payment_method.length; i++){
		if (document.f_data.payment_method[i].value == document.f_data.i_payment_method.value){
			document.f_data.payment_method[i].checked = true;
			break;
		}
	}

	// �x�����@�ύX
	checkPaymentMethod();

	// �͐�ꗗ
	for (var i=0; i<document.f_data.select_ship_address_no.options.length; i++){
		if (document.f_data.select_ship_address_no.options[i].value == document.f_data.i_ship_address_no.value){
			document.f_data.select_ship_address_no.options[i].selected = true;
			break;
		}
	}
	changeShipAddress();

	//2014/08/05
	// �[�i�����t
	if (document.f_data.ship_invoice_fl) {
		if (document.f_data.i_ship_invoice_fl.value == "Y"){
			document.f_data.ship_invoice_fl[2].checked = true;
		}
		if (document.f_data.i_ship_invoice_fl.value == "N"){
			document.f_data.ship_invoice_fl[1].checked = true;
		}
		if (document.f_data.i_ship_invoice_fl.value == "X"){
			document.f_data.ship_invoice_fl[0].checked = true;
		}
	}

//2011/06/01 if-web del start
	// �^�����

//	for (var i=0; i<document.f_data.freight_forwarder.length; i++){
//		if (document.f_data.freight_forwarder[i].value == document.f_data.i_freight_forwarder.value){
//			document.f_data.freight_forwarder[i].checked = true;
//			break;
//		}
//	}
//2011/06/01 if-web del end

	// �^����Ўw��ύX
	//sel_FreightForwarder();

	// �����w�肠��A�Ȃ�
	if (((document.f_data.i_delivery_mm.value != "") && (document.f_data.i_delivery_dd.value != "")) || (document.f_data.i_delivery_tm.value != "")) {
		document.f_data.delivery_fl[1].checked = true;
	} else {
		document.f_data.delivery_fl[0].checked = true;
	}

	// �����w��ύX
	checkDeliveryDate();

	for (var i=0; i<document.f_data.delivery_mm.options.length; i++){
		if (document.f_data.delivery_mm.options[i].value == document.f_data.i_delivery_mm.value){
			document.f_data.delivery_mm.options[i].selected = true;
			break;
		}
	}

	for (var i=0; i<document.f_data.delivery_dd.options.length; i++){
		if (document.f_data.delivery_dd.options[i].value == document.f_data.i_delivery_dd.value){
			document.f_data.delivery_dd.options[i].selected = true;
			break;
		}
	}

	// ���Ԏw��
	for (var i=0; i<document.f_data.delivery_tm.options.length; i++){
		if (document.f_data.delivery_tm.options[i].value == document.f_data.i_delivery_tm.value){
			document.f_data.delivery_tm.options[i].selected = true;
			break;
		}
	}

	// �c�Ə��~��
	if (document.f_data.i_eigyousho_dome_fl.value == "Y"){
		document.f_data.eigyousho_dome_fl.checked = true;
	}

	// �ꊇ�o��
	if (document.f_data.i_ikkatsu_fl.value == "Y"){
		document.f_data.ikkatsu_fl[0].checked = true;
	}
	if (document.f_data.i_ikkatsu_fl.value == "N"){
		document.f_data.ikkatsu_fl[1].checked = true;
	}

	// �̎���
	if (document.f_data.receipt_fl.type != "hidden"){
		if (document.f_data.i_receipt_fl.value == "N"){
			document.f_data.receipt_fl[0].checked = true;
			// 2014/08/20
			document.getElementById('ship_invoice_fl_x').disabled = false;
		}
		if (document.f_data.i_receipt_fl.value == "Y"){
			document.f_data.receipt_fl[1].checked = true;
			// 2014/08/20
			if (document.getElementById('ship_invoice_fl_x').checked == true) {
				document.getElementById('ship_invoice_fl_n').checked = true;
			}
			document.getElementById('ship_invoice_fl_x').disabled = true;
		}
	}

	// �̎����ύX
	checkReceipt();

	// �ߕs�������g�p����
	if (document.f_data.i_rebate_fl.value == "Y"){
		document.f_data.RebateFl.checked = true;
	}
}

//=====================================================================
//	�^����� �I��ύX��  2011/06/01 mod �i����݂̂Ɂj
//=====================================================================
function sel_FreightForwarder(){

	// ����}��
//	if (document.getElementById('freight_forwarder_1').checked == true){
//		document.f_data.delivery_tm.options.length = 6;
//		document.f_data.delivery_tm.options[0].value = "";
//		document.f_data.delivery_tm.options[1].value = "�ߑO��";
//		document.f_data.delivery_tm.options[2].value = "12������14���܂�";
//		document.f_data.delivery_tm.options[3].value = "14������16���܂�";
//		document.f_data.delivery_tm.options[4].value = "16������18���܂�";
//		document.f_data.delivery_tm.options[5].value = "18������21���܂�";
//		document.f_data.delivery_tm.options[0].text = "";
//		document.f_data.delivery_tm.options[1].text = "�ߑO��";
//		document.f_data.delivery_tm.options[2].text = "12������14���܂�";
//		document.f_data.delivery_tm.options[3].text = "14������16���܂�";
//		document.f_data.delivery_tm.options[4].text = "16������18���܂�";
//		document.f_data.delivery_tm.options[5].text = "18������21���܂�";
//	}

	// ���}�g�^�A
//	if (document.getElementById('freight_forwarder_2').checked == true){
//		document.f_data.delivery_tm.options.length = 7;
//		document.f_data.delivery_tm.options[0].value = "";
//		document.f_data.delivery_tm.options[1].value = "�ߑO��";
//		document.f_data.delivery_tm.options[2].value = "12������14��";
//		document.f_data.delivery_tm.options[3].value = "14������16��";
//		document.f_data.delivery_tm.options[4].value = "16������18��";
//		document.f_data.delivery_tm.options[5].value = "18������20��";
//		document.f_data.delivery_tm.options[6].value = "20������21��";
//		document.f_data.delivery_tm.options[0].text = "";
//		document.f_data.delivery_tm.options[1].text = "�ߑO��";
//		document.f_data.delivery_tm.options[2].text = "12������14��";
//		document.f_data.delivery_tm.options[3].text = "14������16��";
//		document.f_data.delivery_tm.options[4].text = "16������18��";
//		document.f_data.delivery_tm.options[5].text = "18������20��";
//		document.f_data.delivery_tm.options[6].text = "20������21��";
//	}

//	document.f_data.delivery_tm.options[0].selected = true;

}

//=====================================================================
//	�x�����@ �I��ύX��
//=====================================================================
function checkPaymentMethod(){

	// ��s�U���̏ꍇ�A�U���l���`���͉A�̎����K�v�I����
	if(document.getElementById('radio_ginkou').checked == true){
		document.getElementById('furikomi_nm').disabled = false;
		document.getElementById('receipt_fl_y').disabled = false;

		$("#receipt1").css("display","inline");
		$("#receipt2").css("display","none");

	}else{
		document.getElementById('furikomi_nm').disabled = true;
		document.getElementById('receipt_fl_y').disabled = true;
		document.getElementById('furikomi_nm').value='';
		document.getElementById('receipt_fl_n').checked=true;
		// �̎����ύX
		checkReceipt();

		$("#receipt1").css("display","none");
		$("#receipt2").css("display","inline");
	}

	// ����̏ꍇ�A�݌ɏ��i����o�וs��
	if(document.getElementById('radio_daibiki').checked == true){
		document.getElementById('ikkatsu_fl_n').disabled = true;
	}else{
		document.getElementById('ikkatsu_fl_n').disabled = false;
	}

	// 2014/08/05
	if(document.f_data.radio_daibiki.checked){
		$("#ship_invoice").css("display", "none");
		document.getElementById('ship_invoice_fl_x').disabled = true;
		document.getElementById('ship_invoice_fl_y').disabled = true;
		document.getElementById('ship_invoice_fl_n').disabled = true;
	} else {
		if (document.f_data.ship_address_no.value != 1) {
			if ($("#i_ship_address_no1").val() == $("#i_selected_ship_address").val()) {
				$("#ship_invoice").css("display", "none");
				document.getElementById('ship_invoice_fl_x').disabled = true;
				document.getElementById('ship_invoice_fl_y').disabled = true;
				document.getElementById('ship_invoice_fl_n').disabled = true;
			} else {
				if ($("#ship_invoice").css("display") == 'none') {
					$("#ship_invoice").css("display", "inline");
					// 2014/08/20
					document.getElementById('ship_invoice_fl_y').disabled = false;
					document.getElementById('ship_invoice_fl_n').disabled = false;
					if (document.getElementById('receipt_fl_n').checked == true) {
						document.getElementById('ship_invoice_fl_x').disabled = false;
						document.getElementById('ship_invoice_fl_x').checked  = true;
					} else {
						if (document.getElementById('ship_invoice_fl_x').checked == true) {
							document.getElementById('ship_invoice_fl_n').checked = true;
						}
					}
				}
			}
		} else {
			$("#ship_invoice").css("display", "none");
		}
	}

}

//=====================================================================
//	�����w�� �I��ύX��
//=====================================================================
function checkDeliveryDate(){

	// �����w�肠��̏ꍇ�A���t�A���ԑI����
	if(document.getElementById('delivery_fl_y').checked == true){
		// 2012/09/06 ok Add
		if(document.getElementById('delivery_mm').disabled){
			$('a[href=#delivery_time]').click();
		}

		document.getElementById('delivery_mm').disabled = false;
		document.getElementById('delivery_dd').disabled = false;
		document.getElementById('delivery_tm').disabled = false;
	}else{
		document.getElementById('delivery_mm').disabled = true;
		document.getElementById('delivery_dd').disabled = true;
		document.getElementById('delivery_tm').disabled = true;
		document.getElementById('delivery_mm').options[0].selected=true;
		document.getElementById('delivery_dd').options[0].selected=true;
		document.getElementById('delivery_tm').options[0].selected=true;
	}

}

//=====================================================================
//	�̎����w�� �I��ύX��
//=====================================================================
function checkReceipt(){

	// �̎����K�v�̏ꍇ�A�̎�������A�̎����A���������͉�
	if(document.getElementById('receipt_fl_y').checked==true){
		document.getElementById('receipt_nm').disabled=false;
		document.getElementById('receipt_memo').disabled=false;

		// 2014/08/20
		if (document.getElementById('ship_invoice_fl_x').checked == true) {
			document.getElementById('ship_invoice_fl_n').checked = true;
		}
		document.getElementById('ship_invoice_fl_x').disabled=true;
	}else{
		document.getElementById('receipt_nm').disabled=true;
		document.getElementById('receipt_memo').disabled=true;
		document.getElementById('receipt_nm').value='';
		document.getElementById('receipt_memo').value='';
		document.getElementById('ship_invoice_fl_x').disabled=false;	// 2014/08/20
	}

}

//=====================================================================
//	���͂���ύX��
//=====================================================================

function changeShipAddress(){

	var i;
	var vAddressNo = new Array(<% = wAddressNoHTML %>);
	var vZip = new Array(<% = wZipHTML %>);
	var vAddress = new Array(<% = wAddressHTML %>);
	var vTelephoneNo = new Array(<% = wTelephoneNoHTML %>);
	var vAddressName = new Array(<% = wAddressNameHTML %>);

	//2012/08/08 nt add Start
	var flag;
	var vRitouFl = new Array(<% = wRitouFlHTML %>);
	var vKuyuKinshiFl = new Array(<% = wKuyuKinshiFlHTML %>);
	//2012/08/08 nt add End

	//2012/08/25 nt add Start
	var vSagawaLTFlg = new Array(<% = wSagawaLTHTML %>);
	//2012/08/25 nt add End

	var idx = document.f_data.select_ship_address_no.selectedIndex;
	if(idx <= 0){
		return;
	}
	var AddrNo = document.f_data.select_ship_address_no.options[idx].value;
	if(AddrNo <= 0){
		return;
	}

	for ( i=0; i<vAddressNo.length; i++){
		if(AddrNo==vAddressNo[i]){
			idx = i;
			break;
		}
	}

	//2014/08/05
	document.getElementById('i_selected_ship_address').value = vAddress[idx];
	if(document.f_data.radio_daibiki.checked){
		$("#ship_invoice").css("display", "none");
		document.getElementById('ship_invoice_fl_x').disabled = true;
		document.getElementById('ship_invoice_fl_y').disabled = true;
		document.getElementById('ship_invoice_fl_n').disabled = true;
	} else {
		if (AddrNo != 1) {
			if ($("#i_ship_address_no1").val() == $("#i_selected_ship_address").val()) {
				$("#ship_invoice").css("display", "none");
			} else {
				if ($("#ship_invoice").css("display") == 'none') {
					$("#ship_invoice").css("display", "inline");
					// 2014/08/20
					document.getElementById('ship_invoice_fl_y').disabled = false;
					document.getElementById('ship_invoice_fl_n').disabled = false;
					if (document.getElementById('receipt_fl_n').checked == true) {
						document.getElementById('ship_invoice_fl_x').disabled = false;
						document.getElementById('ship_invoice_fl_x').checked  = true;
					} else {
						if (document.getElementById('ship_invoice_fl_x').checked == true) {
							document.getElementById('ship_invoice_fl_n').checked = true;
						}
					}
				}
			}
		} else {
			$("#ship_invoice").css("display", "none");
			document.getElementById('ship_invoice_fl_x').disabled = true;
			document.getElementById('ship_invoice_fl_y').disabled = true;
			document.getElementById('ship_invoice_fl_n').disabled = true;
		}
	}



	document.f_data.ship_address_no.value = AddrNo;
	document.getElementById('ShipZip').innerHTML = vZip[idx];
	document.getElementById('ShipAddress').innerHTML = vAddress[idx];
	document.getElementById('ShipTel').innerHTML = 'Tel. ' + vTelephoneNo[idx];
	document.getElementById('ShipName').innerHTML = vAddressName[idx] + ' �l';
	document.f_data.select_ship_address_no.length = vAddressNo.length - 1;

	//2012/08/08 nt add Start
	//���d�ʕi�Ή�
	flag=0;
	if (vRitouFl[idx] == "Y" && vKuyuKinshiFl[idx] == "Y") {
		//�����t���O�FY + ���d�ʕi�t���O�FY�̏ꍇ�A�u����v�I��s��
		document.getElementById('radio_daibiki').disabled=true;

		//�u����v���I������Ă����ꍇ�A�u��s�U���v�փ��W�I�{�^����checked��ύX
		if(document.f_data.radio_daibiki.checked){
			flag = 1;

			if (flag = 1) {
				document.getElementById('radio_ginkou').checked=true;
			}
		}

		//�\���E��\����ؑւ�
		$("#lDaibiki2").css("display","inline");
		$("#lDaibiki3").css("display","none");
		$("#lDaibiki").css("display","none");

	//2012/08/25 nt add Start
	//���쐧���Ή�
	}else if (vSagawaLTFlg[idx] == "Y") {
		//�������������t���O�FY�̏ꍇ�A�u����v�I��s��
		document.getElementById('radio_daibiki').disabled=true;

		//�u����v���I������Ă����ꍇ�A�u��s�U���v�փ��W�I�{�^����checked��ύX
		if(document.f_data.radio_daibiki.checked){
			flag = 1;

			if (flag = 1) {
				document.getElementById('radio_ginkou').checked=true;
			}
		}

		//�\���E��\����ؑւ�
		$("#lDaibiki3").css("display","inline");
		$("#lDaibiki2").css("display","none");
		$("#lDaibiki").css("display","none");

	//2012/08/25 nt add End
	}else if (vRitouFl[idx] != "Y" && vKuyuKinshiFl[idx] == "Y") {
		//���d�ʕi�t���O�FY�݂̂̏ꍇ�A�u����v�I����
		document.getElementById('radio_daibiki').disabled=false;

		//�\���E��\����ؑւ�
		$("#lDaibiki2").css("display","none");
		$("#lDaibiki3").css("display","none");
		$("#lDaibiki").css("display","inline");

	//2012/08/25 nt add Start
	//���쐧���Ή�
	}else if (vSagawaLTFlg[idx] != "Y") {
		//�������������t���O�FY�łȂ��ꍇ�A�u����v�I����
		document.getElementById('radio_daibiki').disabled=false;

		//�\���E��\����ؑւ�
		$("#lDaibiki2").css("display","none");
		$("#lDaibiki3").css("display","none");
		$("#lDaibiki").css("display","inline");

	//2012/08/25 nt add End
	}else{
		//��L�ȊO�́u����v�I����
		document.getElementById('radio_daibiki').disabled=false;
	}
	//2012/08/08 nt add End

	idx = 0;
	for ( i=0; i<vAddressNo.length; i++){
		if(AddrNo!=vAddressNo[i]){
			document.f_data.select_ship_address_no.options[idx].value = vAddressNo[i];
			document.f_data.select_ship_address_no.options[idx].text = vAddressName[i] + '�@' + vZip[i] + ' ' + vAddress[i] + '�@' + vTelephoneNo[i];
			idx++;
		}
	}
	document.f_data.select_ship_address_no.options[0].selected = true;
}

</script>

</head>

<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>

<!-- �R���e���cstart -->
<div id="globalContents">

  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now">���͂���A���x�������@�̑I��</li>
    </ul>
  </div></div></div>

<% If wMsg <> "" Then %>
  <p class="error"><% = wMsg %></p>
<% End If %>

<% If wErrMsg <> "" Then %>
  <p class="error"><% = wErrMsg %></p>
<% End If %>

  <h1 class="title">���͂���A���x�������@�̑I��</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.�V���b�s���O�J�[�g" width="170" height="50"></li>
    <li><img src="images/step02_now.gif" alt="2.���͂���A���x�����@�̑I��" width="170" height="50"></li>
    <li><img src="images/step03.gif" alt="3.���������e�̊m�F" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.����������" width="170" height="50"></li>
  </ol>

  <h2 class="cart_title">���͂���</h2>
  <form action="JavaScript:OrderSubmit('next')" method="post" name="f_data" >
    <table id="address">
      <tr>
        <td class="main">
          <div class="box_l">
            <p>
              <div id="ShipZip"><% = "��" & wCustomerZip %></div>
              <div id="ShipAddress"><% = wCustomerPref & wCustomerAddress %></div>
              <div id="ShipTel"><% = "Tel. " & wCustomerTel %></div>
              <div id="ShipName"><% = wCustomerNm & " �l" %></div>
            </p>
          </div>
          <div class="box_r">
<% = wShipAddressHTML %>
            <p class="change"><a href="JavaScript:OrderSubmit('address');">�V�����Z����o�^����</a></p>
          </div>
        </td>
      </tr>
<%
' 2014/08/05 GV mod start
'1)�����(����悪�{�l�łȂ��Ă��j�͕\�����Ȃ�
'2)�͂��悪�{�l�ƈ�ꍇ�A�\��
'3)�o�^�Z���Ɠ��͂��ꂽ�V�Z���i�A��1�ȊO�j���قȂ�ꍇ�A�\��
'  (�o�^�Z���Ɠ������e�𖈉���͂��邨�q�l�ւ̑Ή��̂���)
wShowInvoice = "none;"

If (wShipInvoiceFl = "X") Then
	wInvoiceChecked(0) = " checked"
ElseIf (wShipInvoiceFl = "N") Then
	wInvoiceChecked(1) = " checked"
ElseIf (wShipInvoiceFl = "Y") Then
	wInvoiceChecked(2) = " checked"
Else
	wInvoiceChecked(0) = " checked"
End If

If (wPaymentMethod = "�����") Then
	wShowInvoice = "none;"
	wInvoiceDisabled = " disabled='disabled'"
Else
	If (wShipAddressNo <> "1") Then
'		If wShipAddressNo1Data = (wCustomerPref & wCustomerAddress) Then
'			wShowInvoice = "none;"
'			wInvoiceDisabled = " disabled='disabled'"
'		Else
			wShowInvoice = "inline;"
			wInvoiceDisabled = ""
'		End If
	Else
		wShowInvoice = "none;"
		wInvoiceDisabled = " disabled='disabled'"
	End If
End If
' 2014/08/05 GV mod end
%>
      <tr id="ship_invoice" style="display:<%=wShowInvoice%>">
        <td class="left">�[�i���̑��t���I�����Ă��������B<br>
          <input type="radio" name="ship_invoice_fl" id="ship_invoice_fl_x" value="X"<%=wInvoiceChecked(0)%><%=wInvoiceDisabled%>><label for="ship_invoice_fl_x">�s�v</label><br>
          <input type="radio" name="ship_invoice_fl" id="ship_invoice_fl_n" value="N"<%=wInvoiceChecked(1)%><%=wInvoiceDisabled%>><label for="ship_invoice_fl_n">�w���ҁi�ʓr�X����z�j...�v���[���g�A���蕨�̍ۂ͂���������I�����������B</label><br>
          <input type="radio" name="ship_invoice_fl" id="ship_invoice_fl_y" value="Y"<%=wInvoiceChecked(2)%><%=wInvoiceDisabled%>><label for="ship_invoice_fl_y">�͂���i�ו��֓����j...�����g�̂��������̍ۂɂ͂���������I�����������B</label>
        </td>
      </tr>
    </table>
    <p class="kome">�����i�����w�����������ۂ́A�u<a href="#notes">�������ɂ�����</a>�v��K�����m�F���������B</p>

    <h2 class="cart_title">���x�������@�̑I��</h2>
    <table id="pay">
      <tr>
        <th>���x�������@</th>
        <th>�z�����@</th>
        <th>����]�z�B����</th>
        <th>�̎���</th>
      </tr>
      <tr>
        <td>
          <ul class="select">
            <li onClick="checkPaymentMethod();">
              <input id="radio_ginkou" name="payment_method" type="radio" value="��s�U��" checked><label for="radio_ginkou">��s�U��</label>
              <span>�U���l���`</span><span><input name="furikomi_nm" type="text" id="furikomi_nm" value="<% = wFurikomiNm %>">�l</span>
            </li>
            <li onClick="checkPaymentMethod();">
              <input id="radio_netbank" name="payment_method" type="radio" value="�R���r�j�x��"><label for="radio_netbank">�l�b�g�o���L���O</label>
              <span><label for="radio_netbank">�iPay-easy<img src="images/payeasy.gif" alt="Pay-easy">�j</label></span>
              <span><label for="radio_netbank">�䂤����</label></span><span><label for="radio_netbank">�R���r�j����</label></span>
            </li>
            <li onClick="checkPaymentMethod();">

              <!-- 2012/08/25 nt mod Start -->
              <!-- <input id="radio_daibiki" name="payment_method" type="radio" value="�����"><label for="radio_daibiki">�������</label> -->
                   <input id="radio_daibiki" name="payment_method" type="radio" value="�����">
                   <label for="radio_daibiki" id="lDaibiki">�������</label>
                   <label for="radio_daibiki" id="lDaibiki2" style="display:none;"><a href="#hmethods" class="fancybox" id="aDaibiki">�������</a></label>
                   <label for="radio_daibiki" id="lDaibiki3" style="display:none;"><a href="#hmethods2" class="fancybox" id="aDaibiki">�������</a></label>
              <!-- 2012/08/25 nt mod End -->
            </li>
            <li onClick="checkPaymentMethod();"><input id="radio_loan" name="payment_method" type="radio" value="���[��"><label for="radio_loan">���[��</label></li>
          </ul>
        </td>
        <td>
          <ul class="select">
            <li><input id="ikkatsu_fl_y" name="ikkatsu_fl" type="radio" value="Y" checked><label for="ikkatsu_fl_y">�ꊇ�o��</label></li>
            <li><input id="ikkatsu_fl_n" name="ikkatsu_fl" type="radio" value="N"><label for="ikkatsu_fl_n">�݌ɏ��i����o��</label></li>
          </ul>
        </td>
        <td>
          <ul class="select">
            <li onClick="checkDeliveryDate();"><input id="delivery_fl_n" name="delivery_fl" type="radio" value="N"><label for="delivery_fl_n">�w��Ȃ�</label></li>
            <li onClick="checkDeliveryDate();">

              <!-- 2012/08/08 nt mod Start -->
              <!-- <input id="delivery_fl_y" name="delivery_fl" type="radio" value="Y"><label for="delivery_fl_y">�w�肠��</label> -->
              <% if (wKuyuKinshiFl = "Y") then %>
                   <input id="delivery_fl_y" name="delivery_fl" type="radio" value="Y" disabled>
                   <label for="delivery_fl_y"><a href="#hdelivery_time" class="fancybox">�w�肠��</a></label>
              <% else %>
                   <input id="delivery_fl_y" name="delivery_fl" type="radio" value="Y">
                   <label for="delivery_fl_y">�w�肠��</label>
              <% end if %>
              <!-- 2012/08/08 nt mod End -->

              <span>
                <select id="delivery_mm" name="delivery_mm" disabled>
                  <option value=""></option>
                  <option value="01">1</option>
                  <option value="02">2</option>
                  <option value="03">3</option>
                  <option value="04">4</option>
                  <option value="05">5</option>
                  <option value="06">6</option>
                  <option value="07">7</option>
                  <option value="08">8</option>
                  <option value="09">9</option>
                  <option value="10">10</option>
                  <option value="11">11</option>
                  <option value="12">12</option>
                </select>��
                <select id="delivery_dd" name="delivery_dd" disabled>
                  <option value=""></option>
                  <option value="01">1</option>
                  <option value="02">2</option>
                  <option value="03">3</option>
                  <option value="04">4</option>
                  <option value="05">5</option>
                  <option value="06">6</option>
                  <option value="07">7</option>
                  <option value="08">8</option>
                  <option value="09">9</option>
                  <option value="10">10</option>
                  <option value="11">11</option>
                  <option value="12">12</option>
                  <option value="13">13</option>
                  <option value="14">14</option>
                  <option value="15">15</option>
                  <option value="16">16</option>
                  <option value="17">17</option>
                  <option value="18">18</option>
                  <option value="19">19</option>
                  <option value="20">20</option>
                  <option value="21">21</option>
                  <option value="22">22</option>
                  <option value="23">23</option>
                  <option value="24">24</option>
                  <option value="25">25</option>
                  <option value="26">26</option>
                  <option value="27">27</option>
                  <option value="28">28</option>
                  <option value="29">29</option>
                  <option value="30">30</option>
                  <option value="31">31</option>
                </select>��
              </span>
              <span>����</span>
              <span>
                <select id="delivery_tm" name="delivery_tm" style="width:115px;" disabled>
                  <option value="" selected="selected"></option>
                  <option value="<%=wDeliveryTime01%>"><%=wDeliveryTime01%></option>
                  <option value="<%=wDeliveryTime02%>"><%=wDeliveryTime02%></option>
                  <option value="<%=wDeliveryTime03%>"><%=wDeliveryTime03%></option>
                  <option value="<%=wDeliveryTime04%>"><%=wDeliveryTime04%></option>
                  <option value="<%=wDeliveryTime05%>"><%=wDeliveryTime05%></option>
                </select>
              </span>
            </li>
            <li><input type="checkbox" id="eigyousho_dome_fl" name="eigyousho_dome_fl" value="Y"><label for="eigyousho_dome_fl">�^����Љc�Ə��~��</label></li>
          </ul>
        </td>
        <td>
          <ul class="select">
            <li onClick="checkReceipt();"><input id="receipt_fl_n" name="receipt_fl" type="radio" value="N" checked><label for="receipt_fl_n">�s�v</label></li>
            <li onClick="checkReceipt();">
              <input id="receipt_fl_y" name="receipt_fl" type="radio" value="Y"><label for="receipt_fl_y" id="receipt1">�K�v</label><label for="receipt_fl_y" id="receipt2" style="display:none;"><a href="#receipt" class="fancybox">�K�v</a></label>
              <span>�̎�������</span><span><input name="receipt_nm" type="text" id="receipt_nm" value="<% = wReceiptNm %>" disabled>�l</span>
              <span>�̎����A������</span><span><input type="text" name="receipt_memo" id="receipt_memo" value="<% = wReceiptMemo %>" disabled></span>
            </li>
          </ul>
        </td>
      </tr>
      <tr>
      	<td class="detail"><a href="#payment" class="fancybox">�ڍׂ͂�����</a></td>
        <td class="detail"><a href="#methods" class="fancybox">�ڍׂ͂�����</a></td>
        <td class="detail"><a href="#delivery_time" class="fancybox">�ڍׂ͂�����</a></td>
        <td class="detail"><a href="#receipt" class="fancybox">�ڍׂ͂�����</a></td>
      </tr>
    </table>

<% If wCustomerKabusokuAm > 0 And wCustomerClass = "��ʌڋq" Then %>
    <dl class="excess">
      <dt>�N���W�b�g�^�ߕs����</dt>
      <dd><% = FormatNumber(wCustomerKabusokuAm, 0) %>�~</dd>
    </dl>
    <p class="excess_use"><input type="checkbox" name='RebateFl' id="RebateFl" value="Y"><label for="RebateFl">���x�����ɃN���W�b�g/�ߕs�������g�p����</label></p>
<% End If %>

<% If wNoData = False Then %>
    <div id="btn_box">
      <ul class="btn next">
        <li><a href="JavaScript:OrderSubmit('next');"><img src="images/btn_next.png" alt="����" class="opover"></a></li>
      </ul>
    </div>
<% End If %>
    <input type="hidden" name="cmd" value="">
    <input type="hidden" name="customer_kn" value="<% = wCustomerKn %>">
    <input type="hidden" name="customer_email" value="<% = wCustomerEmail %>">
    <input type="hidden" name="telephone" value="<% = wCustomerTel %>">
    <input type="hidden" name="KabusokuAm" value="<% = wCustomerKabusokuAm %>">
    <input type="hidden" name="i_rebate_fl" value="<% = wRebateFl %>">
    <input type="hidden" name="i_payment_method" value="<% = wPaymentMethod %>">
    <input type="hidden" name="i_ship_address_no" value="<% = wShipAddressNo %>">
    <input type="hidden" name="i_ship_invoice_fl" value="<% = wShipInvoiceFl %>">
    <input type="hidden" name="i_ikkatsu_fl" value="<% = wIkkatsuFl %>">
    <input type="hidden" name="i_freight_forwarder" value="<% = wFreightForwarder %>">
    <input type="hidden" name="i_delivery_mm" value="<% = wDeliveryMM %>">
    <input type="hidden" name="i_delivery_dd" value="<% = wDeliveryDD %>">
    <input type="hidden" name="i_delivery_tm" value="<% = wDeliveryTM %>">
    <input type="hidden" name="i_eigyousho_dome_fl" value="<% = wEigyoushoDomeFl %>">
    <input type="hidden" name="i_receipt_fl" value="<% = wReceiptFl %>">
    <input type="hidden" name="i_tokuchuu_fl" value="<% = wTokuchuuFl %>">
    <input type="hidden" name="i_daibiki_fuka_fl" value="<% = wDaibikiFukaFl %>">
    <input type="hidden" name="ship_address_no" value="<% = wShipAddressNo %>">
    <input type="hidden" name="freight_forwarder" value="5">
  </form>
<% '2014/08/05 GV add start %>
  <form id='f_shipping_data' name='f_shipping_data'>
    <input type="hidden" name="i_ship_address_no1" id="i_ship_address_no1" value="<% = wShipAddressNo1Data %>">
    <input type="hidden" name="i_selected_ship_address" id="i_selected_ship_address" value="<% = wSelectedShipAddressData %>">
  </form>
<% '2014/08/05 GV add end %>
  <ul class="info left">
    <li><a href="#cancel" class="fancybox">���������i�̃L�����Z���E�ԕi�ɂ���</a></li>
    <li><a href="#delivery" class="fancybox">���i�̔[���ɂ��Ă͂�����</a></li>
  </ul>

  <h2 id="notes" class="cart_title">�������ɂ�����</h2>
  <h3 style="font-weight:bold;">�����z�B�T�[�r�X�ɂ���</h3>
  <p>���͂���͊֓���s�����i�����A�_�ސ�A��t�A��ʁA���A�Ȗ؁A�Q�n�A�R���j�Ɍ��肳��Ă���܂��B<br>�܂��A�ȉ��̏ꍇ�͓����z�B�ΏۊO�ƂȂ�܂��B</p>
  <ul id="attention" style="margin:.5em;">
    <li>��^���i���܂ނ�����</li>
    <li>�c�Ə��~�߂̂�����</li>
    <li>�z�B�����w�肪����Ă��邲����</li>
    <li>���x�������@�u��������v�ȊO�̂�����</li>
  </ul>
  <p class="notice"><span style="color:red">��</span>�����z���T�[�r�X�ł́A�z�B���ԑт̂���]�͂��󂯂��邱�Ƃ��ł��܂���B</p>

  <ul id="attention" style="border:3px solid #ccc; line-height:1.8;background-color:#f0f0f0;padding:.8em;">
    <li>�m�F���[���͎����I�ɑ��M����܂��B�����_��͏��i�̔����������Đ����ƂȂ�܂��B</li>
    <li>�g�ѓd�b�Ȃǂ̎�M����������A�h���X�ł��o�^���ꂽ�ꍇ�A ��������񂪎�M�ł��Ȃ��ꍇ���������܂��B</li>
    <li>���������i�ɂ��Ă̂��₢���킹�́A���[���₨�d�b�ɂď����Ă���܂��̂ł������O�ɂ��m�F���������܂��悤���肢�������܂��B</li>
    <li>�J�[�g�ɓ��ꂽ���i�ȊO������]�̍ۂ́A���炩���߃��[���₨�d�b�ɂĂ��₢���킹���������B</li>
  </ul>

  <div style="display:none;">

    <!-- ���x�����@ -->
    <div id="payment">
      <h2>���x�������@�ɂ���</h2>
      <div>
        <ul>
          <li>��s�U��
            <ul>
              <li>�U���l���`���͉���o�^�̂����O�ƈقȂ�ꍇ�݂̂��L�����������B</li>
              <li>�U���萔���͂��q�l�̕��S�Ƃ����Ă��������܂��B</li>
              <li>��قǁA�����Ϗ������ē��������܂��̂ŁA����������m�F��ɂ��U���݂��������܂��悤���肢�������܂��B</li>
            </ul>
          </li>
          <li>�l�b�g�o���L���O�E�䂤����E�R���r�j����
            <ul>
              <li>���[�\���A�t�@�~���[�}�[�g�A�T�[�N��K�A�T���N�X�A�Z�C�R�[�}�[�g�A�䂤�����s�A�l�b�g�o���L���O �ł��x�����������܂��B</li>
              <li>��������A���z���ύX�ƂȂ邲�����̕ύX�͏��邱�Ƃ��ł��܂���B�݌ɂ̖������i�Ȃǂ��������̍ۂ́A���O�ɂ��₢���킹���������B</li>
              <li>E-MAIL�A�h���X���g�т̏ꍇ�́A�K�v�������m�F�ł��Ȃ��ꍇ�����邽�߁A�p�\�R������̂����p���������߂��܂��B</li>
              <li>��قǁA�����Ϗ������ē��������܂��̂ŁA����������m�F��ɂ��U���݂��������܂��悤���肢�������܂��B</li>
            </ul>
          </li>
          <li>�������
            <ul>
              <li>��������ł̂��w���̏ꍇ�A���i�̔����͈ꊇ�o�ׂƂȂ�܂��B�܂��A���x�����͌����݂̂̎�t�ƂȂ�܂��B</li>
            </ul>
          </li>
          <li>���[��
            <ul>
              <li>�I�����C�����[���̏ꍇ����\����̂��������e�̕ύX�����邱�Ƃ��ł��܂���B</li>
              <li>���������e�Ƥ�I�����C�����[���\���t�H�[���̓��e�����m�F�̏ゲ�������������B</li>
              <li>�W���b�N�X�ł��\�����݂̏ꍇ�́A�����Ȃ��ƂȂ�܂��B</li>
            </ul>
          </li>
        </ul>
        <p class="info"><a href="http://guide.soundhouse.co.jp/guide/oshiharai.asp" target="_blank">���x�������@�ɂ��ďڂ����͂�������������������B</a></p>
      </div>
    </div>

    <!-- �z�����@ -->
    <div id="methods">
      <h2>�z�����@�ɂ���</h2>
      <ul>
        <%
        '<li>���w�肪�����ꍇ�͍���}�ւŔ����������܂��B</li>
        %>
        <li>
          <%
          '����ȂǗ����ւ̂��͂��̓T�C�Y�̏��������i�Ɍ��胄�}�g�^�A�Ŕ����������܂��B<br>�܂��A
	      %>
          ��^�̏��i�̏ꍇ�͂��͂��܂ł�1�T�Ԓ��x�����Ԃ����������ꍇ������܂��B
        </li>
        <li>�����́A���������i���S�đ��������_�ł܂Ƃ߂Ĕ�������u�ꊇ�o�ׁv�܂��͍݌ɂ̂�����̂���s�x��������u�݌ɏ��i����o�ׁv�̂����ꂩ�̕��@���������̍ۂɑI���ł��܂��B</li>
        <li>�������i�������������������ꍇ�A���i�ɂ���Ă͓����������ł��ʔz���ƂȂ�ꍇ������܂��B</li>
        <li>��������ł̔����͈ꊇ�����݂̂ƂȂ�܂��B
          <ul>
            <li>�����񂹏��i���܂܂��ꍇ�A���i����������̈ꊇ�o�ׂƂȂ�܂��B</li>
            <%
            '<li>���}�g�^�A�ł̑�������̔�����1���݂̂ƂȂ�܂��B</li>
            %>
          </ul>
        </li>
      </ul>
    </div>

    <!-- �z����ЁE�z������ -->
    <div id="delivery_time">
      <h2><!--�^����ЁE-->�z�������̎w��</h2>
      <ul>
        <li>�V��A��ʏ󋵂Ȃ�тɔz�B�Ǝ҂̓s���ɂ�育��]�ɓY���Ȃ��ꍇ���������܂��B�����ɂ͗]�T�������Ă��������������B</li>
        <%
        '<li>
        '  ����}�ւ̏ꍇ�A���ԑюw��͕����Ȃ�тɂ��͂��悪
        '  �s�s���ɂ��Z�܂��̌l��̏ꍇ�Ɍ���\�ł��B
        '</li>
        %>
        <li>�ꕔ�A�z�B�����̂���]�����󂯂ł��Ȃ��n�悪�������܂��B�ڍׂ͓d�b�������̓��[���ɂĂ��₢���킹���������B</li>
        <li>�^����Љc�Ə��~�߂��w�肳�ꂽ�ꍇ�A���͐�Z���ɊY�������z�։�Ђ̋K��c�Ə��ւ̗��ߒu���ƂȂ�܂��B</li>
        <li>�����Ȃ�ꍇ�ɂ����Ă��z�B�x�����琶���鑹�Q(���Ɨ��v�̑����A���Ƃ̒x���E���f�A���Ə��̑����܂��͂��̑��̋��K�I���Q��)�Ɋւ��āA�T�E���h�n�E�X�͈�؂̐ӔC�𕉂��܂���B</li>
        <li>��������̏ꍇ�A�z�����w���10���ȓ��̓��t���w�肵�Ă��������B</li>
      </ul>
    </div>

    <!-- �̎��� -->
    <div id="receipt">
      <h2>�̎����̔��s�ɂ���</h2>
      <ul>
        <li>�̎����́A�[�i���̍Ō�Ɉ������Ă���܂��B���萔�ł����A�؂藣���Ă��g�p���������B</li>
        <li>���x�������@���ȉ��̏ꍇ�A�T�E���h�n�E�X�̗̎����͔��s�������܂���B
          <ol>
            <li>�������</li>
            <li>���[��</li>
            <li>�R���r�j/�X�֋ǎx��</li>
          </ol>
        </li>
        <li>�����A�A�����́A�w�藓�ɓ��͂������������e�̂܂܍쐬�������܂��B</li>
        <li>1���̂������ɑ΂��āA�̎�����1���̂ݔ��s�����Ă��������܂��B</li>
      </ul>
      <p class="info"><a href="http://guide.soundhouse.co.jp/guide/kaimono.asp#ryousyuu" target="_blank">�̎����ɂ��ďڂ����͂�������������������B</a></p>
    </div>

    <!-- �L�����Z���E�ԕi -->
    <div id="cancel">
      <h2>�w�����ꂽ���i�̌����A�L�����Z���ɂ���</h2>
      <div>
        <p>�T�E���h�n�E�X�ł́A�����Ƃ��Ă��q�l�̂��s���ɂ��L�����Z���E�ԕi�͏��邱�Ƃ��ł��܂���B���i�̏ڍׂȎd�l��[���ȂǁA���s���ȓ_�͎��O�ɂ��₢���킹�̏�A���������������܂��悤���肢�\���グ�܂��B<br>�������A���i������7���ȓ��ɂ��\���o����������΁A���L�̏����ɊY������ꍇ�̂݌����E�L�����Z��������܂��B</p>
        <h3>�����E�L�����Z�����\�ȃP�[�X</h3>
        <ol>
          <li>�j���A�둗�̏ꍇ
            <p>���ꏤ�i���j�����Ă�����A�������ƈقȂ鏤�i���͂����ꍇ�́A���i������7���ȓ��ɂ��A�����������B�S���X�^�b�t���j���E�둗���i�̈������A�đ��̎菇�ɂ��ďڂ������������Ă��������܂��B</p>
          </li>
          <li>�����s�ǂ̏ꍇ
            <p>���w���������������i�ɕs�������A���i������7���ȓ��ɂ��A������������΁A���Ђɂď��i�̏�Ԃ��m�F��A�ʏ��i�ւ̌����A�������͂������̃L�����Z��������܂��B����́A���q�l���w�肳�ꂽ�����ɕԋ��A�������͂��a������Ƃ��Ď��񂲒������������ۂɑ��E�������܂��B<br>�������A���L�ɋL�ڂ���Ă���6�Ԃ́u���q�l�̎g�p����@�ނ̑��������R�ŏ��i���쓮���Ȃ��ꍇ�v�������܂��B</p>
          </li>
          <li>���q�l�s���ɂ��ꍇ
            <p>���͂��������i�̃L�����Z���A�܂��͑����i�ւ̌�������]�̏ꍇ�A�Y�����i�����J���A���g�p�ł���A�Ȃ������i���󂯎����������7���ȓ��ɂ��A������������΁A�L�����Z���A�܂��͑����i�ւ̌���������܂��B�܂��A���A������J������Ă���ꍇ�ł��A���g�p�ł���΁A���i�����15%�����x���������������Ƃɂ��A�L�����Z���܂��͑����i�ւ̌���������܂��B
            </p>
          </li>
        </ol>
        <h3>�����E�L�����Z�����ł��Ȃ��P�[�X</h3>
        <p>���̏ꍇ�́A���i������7���ȓ��ł����Ă���������уL�����Z���͂��󂯂ł��܂���B</p>
        <ol>
          <li>���i���J�����Ďg�p�����ꍇ</li>
          <li>�\�t�g�E�F�A���J�������ꍇ</li>
          <li>���[�J�[���ԕi���󂯕t���Ȃ��ꍇ</li>
          <li>�̂ɒ��ڐg�ɂ��鏤�i�̏ꍇ</li>
          <li>���q�l�̎w��ɂ�邨���񂹏��i������i</li>
          <li>���q�l�̎g�p����@�ނ̑��������R�ŏ��i���쓮���Ȃ��ꍇ</li>
          <li>���q�l�̌��ŉ����A�j�������������i</li>
          <li>���i�ɕt������I���W�i���̊O������э���ނ����ׂđ����Ă��Ȃ��ꍇ</li>
        </ol>
        <p>���ڍׂ͕��ЃJ�X�^�}�[�T�|�[�g�܂ł��₢���킹���������B</p>
        <h3>���i�̌����A�L�����Z���̎菇</h3>
        <ol>
          <li>���i�̌����A�܂��̓L�����Z��������]�̏ꍇ�A���O�ɏڍׂ��m�F������ŕ��Д��s��RA�i�ԕi���F�j�ԍ����K�v�ƂȂ�܂��B���i�𑗕t����O�ɕK�����Ђ̃J�X�^�}�[�T�|�[�g�܂ł��A�����������B���e�m�F�̏�ARA(�ԕi���F)�ԍ��𔭍s���܂��B</li>
          <li>���i�������肢�������ۂ́A�^����Ђ̑����̔��l����RA�ԍ������L�����������B���i�̊O�����̂��̂ɂ͋L�ڂ��Ȃ��悤���肢�������܂��BRA�ԍ����L�ڂ���Ă��Ȃ����i�͎�̂ł��܂���̂ŁA���炩���߂��������������B</li>
          <li>�ۏ؊��ԓ��̏C���Ȃǂŕ��Ђ������𕉒S����ꍇ�A�w��̉^����Ёi�ʏ�͍���}�ցj�ȊO�͗L���ƂȂ�܂��B���q�l���������ɂď��i�𔭑������ꍇ�́A�ǂ̉^����Ђł������p�\�ł��B</li>
          <li>���i�𑗕t�����ꍇ�́A�K���O���A�}�j���A���A�ۏ؏����A�S�Ă̕t���i���󂯎�莞�Ɠ�����Ԃ̂܂ܑ��t���Ă��������B</li>
          <li>���i������A���Ђ������̓��[�J�[�ɂď��i�̌��i���s���A�����A�܂��̓L�����Z���̏�����i�߂����Ă��������܂��B���ꏤ�i�ƌ����ł��Ȃ��ꍇ�́A���z�𒲐�������ő����i�ւ̌����A�������͕ԋ��ɂđΉ�����ꍇ���������܂��B</li>
          <li>���q�l�֏��i�̑����ԋ�����ꍇ�A���Ђɏ��i���������A���i�A�m�F���s������ɁA���q�l���w��̋�s�����֐U���݂ɂĕԋ��������܂��B</li>
          <li>�ԕi�E�����̂��\���o����2�T�Ԉȓ��ɂ��ԑ����������Ȃ��ꍇ�́A��U�L�����Z�������Ƃ����Ă��������܂��̂ł��炩���߂��������������B</li>
        </ol>
      </div>
    </div>

    <!-- �[���ɂ��� -->
    <div id="delivery">
      <h2>���i�̔[���ɂ���</h2>
      <ul>
        <li>�E�F�u�T�C�g��A����т������₨���ς莞�_�ł��ē����Ă���܂��[���ɂ��܂��ẮA�����܂ł��\��ƂȂ��Ă���A������ɂ��ύX�ƂȂ�ꍇ���������܂��B</li>
        <li>���i�̔[���ɂ��܂��ẮA���[���₨�d�b�ł̂��₢���킹�������Ă���܂��B�w����܂łɔ[�i���K�v�Ȃ������́A�����Ȃ����O�ɂ����k���������B</li>
        <li>�Ȃ��A�[���x���ɂ���Đ��������ɂ��܂��ẮA���Ђł͈�؂̐ӂ𕉂����Ƃ��ł��܂���B���炩���߂��������������B</li>
      </ul>
    </div>

    <!-- 2012/08/08 nt add Start -->
    <!-- ���d�ʕi ��������ɂ��� -->
    <div id="hmethods">
      <ul>
        <p align="center">���q�l�̂������́A�d�ʁA�������͑傫�����K��l�𒴂��鏤�i���܂ވׁA��������ȊO�̂��x�����@��I�����Ă��������B</p>
      </ul>
    </div>

    <!-- ���d�ʕi �z�B�����w��ɂ��� -->
    <div id="hdelivery_time">
      <ul>
        <p align="center">���q�l�̂������́A�d�ʁA�������͑傫�����K��l�𒴂��鏤�i���܂ވׁA�z�B�����w��Ȃ��ł̂��͂��ƂȂ�܂��B</p>
      </ul>
    </div>
    <!-- 2012/08/08 nt add End -->

    <!-- 2012/08/25 nt add Start -->
    <!-- ���쐧�� ��������ɂ��� -->
    <div id="hmethods2">
      <ul>
        <p align="center">���q�l�̂������́A������������󂯂ł��Ȃ��n��ׁ̈A��������ȊO�̂��x�����@��I�����Ă��������B</p>
      </ul>
    </div>
    <!-- 2012/08/08 nt add End -->

  </div>

<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/jquery.fancybox-1.3.4.pack.js"></script>
<script type="text/javascript">
$(function(){
	$(".fancybox").fancybox({
	'scrolling'		: 'no',
	'titleShow'		: false
	});
});
preset_values();
</script>
</body>
</html>