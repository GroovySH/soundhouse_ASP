<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
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
'	�I�[�_�[�y�[�W
'�X�V����
'2004/12/20 �J�[�h�L��������/���Ȃ��f�[�^�̑΍�
'2005/01/04 �s���菤�i�̒��c�o�ׂ��܂Ƃ߂�ύX(�s����ꊇ�Ή�)
'2005/01/27 ���ϋ@�\�폜�Ɋւ����C��
'2005/04/05 �J�[�h����ۑ�����`�F�b�N�{�b�N�X�ǉ�
'2005/04/25 ���̃y�[�W���烊���N���Ă����ʂ��Window�ŊJ���悤�ɕύX
'2005/06/20 �I���R���[���ǉ�
'2005/06/28 �ʐM���̔��l�ύX
'2005/06/29 �ʐM�����폜
'2005/07/06 �������𢉓�u�n��ɕύX
'2005/09/20 �̎����A�������R�����g�ύX
'2005/11/17 �J�[�h���͗���1�ɂ܂Ƃ߂�
'2005/11/18 Input��Value�ɢ�f����Ȃ��L�q������@Value�ɋ󔒂�����Ƃ����Ő؂�邽��
'2006/01/09 ���[�����W�I�{�^���������`�F�b�N���Ȃ��悤�ɕύX
'2006/06/14 �I���R���[�����폜
'2006/06/28 Hidden�ł����Ă����J�[�h�ԍ����폜
'2006/06/29 �I���R���[������
'2006/06/30 �J�[�h�R�����g�ύX
'2006/08/11 �I���R���[���폜
'2006/10/24 �R���r�j���ϒǉ�
'2006/12/08 �R���r�j���σR�����g�ύX
'2006/12/15 �̎����R�����g�ύX
'2007/01/12 �u�R���r�j�x���v���u�R���r�j/�X�֋ǎx���v�ɕ\���ύX
'2007/01/22 �Z��������������̌Ăяo�����͓s���{���݂̂��Z�b�g
'2007/03/20 ���쎞�Ԏw���ύX
'2007/08/14 �J�[�h�G���[���̃��b�Z�[�W�ύX
'2007/09/10 ��]���ԑтɃ^�C�g���ύX
'2007/12/11 �ꊇ�A�����o�בI������ɕ\��
'2008/04/14 ���x�[�g�@�\�ǉ��A�J�[�h�����o�������폜
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2008/08/28 �u�R���r�j/�X�֋ǎx���v���u�R���r�j�G���X�X�g�A/�䂤�����s�x����/Pay-easy�v�ɕ\���ύX
'2008/09/02 �u�R���r�j�G���X�X�g�A/�䂤�����s�x����/Pay-easy�v���u�l�b�g�o���L���O�E�X���E�R���r�j�����v�ɕ\���ύX
'2009/04/21 JACCS�I�����C�����[���ǉ�
'
'========================================================================

On Error Resume Next

Dim w_sessionID
Dim userID
Dim userName
Dim msg

Dim CardErrorCd

Dim customer_nm
Dim furigana
Dim customer_email
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim fax

Dim payment_method
Dim furikomi_nm
Dim loan_downpayment_fl
Dim loan_downpayment_am
Dim loan_term
Dim loan_am
Dim loan_apply_fl
Dim loan_company

Dim ship_address_no
Dim ship_name
Dim ship_zip
Dim ship_prefecture
Dim ship_address
Dim ship_telephone
Dim ship_invoice_fl
Dim freight_forwarder
Dim delivery_mm
Dim delivery_dd
Dim delivery_tm
Dim eigyousho_dome_fl

Dim receipt_fl
Dim receipt_nm
Dim receipt_memo
Dim receipt_nm_org
Dim receipt_memo_org

Dim CustomerClass
Dim KabusokuAm
Dim RebateFl

Dim ikkatsu_fl

Dim i_tokuchuu_fl
Dim i_toriyose_fl
Dim i_daibiki_fuka_fl

Dim wSalesTaxRate
Dim wPrice
Dim wNoData
Dim wShipAddressHTML
Dim wOrderProductHTML

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS
Dim RS_customer
Dim RS_order

Dim w_sql
Dim w_html
Dim wMSG

'========================================================================

Response.Expires = -1			' Do not cache

'---- UserID ���o��
userID = Session("userID")
userName = Session("userName")
w_sessionID = Session.SessionID

'---- Get input data
msg = Session.contents("msg")
wMSG = Session.contents("msg")
Session("msg") = ""
CardErrorCd = ReplaceInput(Request("CardErrorCd"))

if msg = "CardError1" then
	wMSG = "<font size='+1'><b>�J�[�h�ł̏������ł��܂���ł����B</b></font><br><br>"
	wMSG = wMSG & "���L�̃G���[�R�[�h�����Q�Ƃ̏�A�ēx��������蒼���Ē������A�ʂ̃J�[�h���̂��x�����@�ɂčēx�����������肢���܂��B<br>"
	wMSG = wMSG & "���A�J�[�h��Ђɒ��ڌ�₢���킹�̍ۂ́A�G���[�̓��e�A���������ꂽ�����i" & fFormatDate(Now()) & " " & fFormatTime(Now()) & "�j�������Ă��`�����������B<br><br>"
	wMSG = wMSG & "<font size='+1'><b>�G���[�R�[�h:" & CardErrorCd & "</b></font><br><br>"
end if

if msg = "CardError2" then
	wMSG = "<font size='+1'><b>�J�[�h�̏���������Ɏ��s�ł��܂���ł����B</b></font><br><br>"
	wMSG = wMSG & "�����͂��ꂽ�J�[�h�������m�F�̏�A�ēx�������������A�ʂ̃J�[�h�A�܂��͑��̂��x�����@�ɂĂ������肢�܂��B<br>"
	wMSG = wMSG & "���A�ǂ����Ă��䒍�������s�ł��Ȃ��ꍇ�́A���L�̃G���[�R�[�h�𖾋L�̏�A���ЃV�X�e���S���܂ł��₢���킹���������B<br><br>"
	wMSG = wMSG & "<font size='+1'><b>�G���[�R�[�h:" & CardErrorCd & "</b></font><br><br>"
end if

'---- Execute main
call connect_db()
call main()
call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()

'---- Connect database

Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
'
Function main()
'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

payment_method = ReplaceInput(Request("payment_method"))
ship_address_no = ReplaceInput(Request("ship_address_no"))

if isNumeric(ship_address_no) = false then
	ship_address_no = ""
end if

wNoData = false
call get_customer()				'�ڋq���̎��o��
call get_order()					'���󒍏��̎��o��
call createOrderHtml()		'�������i�ꗗHTML�쐬
call get_todokesaki()			'�ڋq�͐���̎��o��

End function

'========================================================================
'
'	Function	�ڋq���̎��o��
'
'========================================================================
'
Function get_customer()

'---- �ڋq�����o��
w_sql = ""
w_sql = w_sql & "SELECT a.�ڋq��"
w_sql = w_sql & "       , a.�ڋq�t���K�i"
w_sql = w_sql & "       , a.�ڋqE_mail1"
w_sql = w_sql & "       , a.�����ߕs�����z"
w_sql = w_sql & "       , a.�ڋq�N���X"
w_sql = w_sql & "       , b.�ڋq�X�֔ԍ�"
w_sql = w_sql & "       , b.�ڋq�s���{��"
w_sql = w_sql & "       , b.�ڋq�Z��"
w_sql = w_sql & "       , c.�ڋq�d�b�ԍ�"
w_sql = w_sql & "      , d.�ڋq�d�b�ԍ� AS FAX"
w_sql = w_sql & "  FROM Web�ڋq a WITH (NOLOCK)"
w_sql = w_sql & "     , Web�ڋq�Z�� b WITH (NOLOCK) LEFT JOIN Web�ڋq�Z���d�b�ԍ� d WITH (NOLOCK)"
w_sql = w_sql & "                                          ON d.�ڋq�ԍ� = b.�ڋq�ԍ�"
w_sql = w_sql & "                                         AND d.�Z���A�� = b.�Z���A��"
w_sql = w_sql & "                                         AND d.�d�b�敪 = 'FAX'"
w_sql = w_sql & "     , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
w_sql = w_sql & " WHERE a.�ڋq�ԍ� = " & userID
w_sql = w_sql & "   AND b.�ڋq�ԍ� = a.�ڋq�ԍ�"
w_sql = w_sql & "   AND b.�Z���A�� = 1"
w_sql = w_sql & "   AND c.�ڋq�ԍ� = a.�ڋq�ԍ�"
w_sql = w_sql & "   AND c.�Z���A�� = 1"
w_sql = w_sql & "   AND c.�d�b�A�� = 1"
	  
'@@@@@response.write(w_sql)

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open w_sql, Connection, adOpenStatic

if RS_customer.EOF = true then
	wMSG = "<center><font color='#ff0000'>�ڋq��񂪂���܂���B</font></center>"
	Session("msg") = wMSG
else
	customer_nm = RS_customer("�ڋq��")
	furigana = RS_customer("�ڋq�t���K�i")
	customer_email = RS_customer("�ڋqE_mail1")
	KabusokuAm = RS_customer("�����ߕs�����z")
	CustomerClass = RS_customer("�ڋq�N���X")
	zip = RS_customer("�ڋq�X�֔ԍ�")
	prefecture = RS_customer("�ڋq�s���{��")
	address = RS_customer("�ڋq�Z��")
	telephone = RS_customer("�ڋq�d�b�ԍ�")

	if isNull(RS_customer("FAX")) = true then
		fax = ""
	else
		fax = RS_customer("FAX")
	end if

end if

RS_customer.close

End function

'========================================================================
'
'	Function	�󒍏��̎��o��
'
'========================================================================
'
Function get_order()

'----���󒍃f�[�^���o��
w_sql = ""
w_sql = w_sql & "SELECT a.�x�����@"
w_sql = w_sql & "     , a.�U�����`�l"
w_sql = w_sql & "     , a.���[����������t���O"
w_sql = w_sql & "     , a.���[������"
w_sql = w_sql & "     , a.��]���[����"
w_sql = w_sql & "     , a.���[�����z"
w_sql = w_sql & "     , a.�I�����C�����[���\���t���O"
w_sql = w_sql & "     , a.���[�����"
w_sql = w_sql & "     , a.���ϔ��l"
w_sql = w_sql & "     , a.�͐�Z���A��"
w_sql = w_sql & "     , a.�͐於�O"
w_sql = w_sql & "     , a.�͐�X�֔ԍ�"
w_sql = w_sql & "     , a.�͐�s���{��"
w_sql = w_sql & "     , a.�͐�Z��"
w_sql = w_sql & "     , a.�͐�d�b�ԍ�"
w_sql = w_sql & "     , a.�͐�[�i�����t�t���O"
w_sql = w_sql & "     , a.�^����ЃR�[�h"
w_sql = w_sql & "     , a.�w��[��"
w_sql = w_sql & "     , a.���Ԏw��"
w_sql = w_sql & "     , a.�c�Ə��~�߃t���O"
w_sql = w_sql & "     , a.�ꊇ�o�׃t���O"
w_sql = w_sql & "     , a.�̎������s�t���O"
w_sql = w_sql & "     , a.�̎�������"
w_sql = w_sql & "     , a.�̎����A������"
w_sql = w_sql & "     , a.���x�[�g�g�p�t���O"
w_sql = w_sql & "     , b.�󒍖��הԍ�"
w_sql = w_sql & "     , b.���[�J�[�R�[�h"
w_sql = w_sql & "     , b.���i�R�[�h"
w_sql = w_sql & "     , b.�F"
w_sql = w_sql & "     , b.�K�i"
w_sql = w_sql & "     , b.���[�J�[��"
w_sql = w_sql & "     , b.���i��"
w_sql = w_sql & "     , b.�󒍐���"
w_sql = w_sql & "     , b.�󒍒P��" 
w_sql = w_sql & "     , b.�󒍋��z" 
w_sql = w_sql & "     , c.���[�J�[�������敪"
w_sql = w_sql & "     , c.����s�t���O" 
w_sql = w_sql & "     , d.�����\����"
w_sql = w_sql & "  FROM ���� a WITH (NOLOCK)"
w_sql = w_sql & "     , ���󒍖��� b WITH (NOLOCK)"
w_sql = w_sql & "     , Web���i c WITH (NOLOCK)"
w_sql = w_sql & "     , Web�F�K�i�ʍ݌� d WITH (NOLOCK)"
w_sql = w_sql & " WHERE a.SessionID = '" & w_sessionID & "'"
w_sql = w_sql & "   AND b.SessionID = a.SessionID"
w_sql = w_sql & "   AND c.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
w_sql = w_sql & "   AND c.���i�R�[�h = b.���i�R�[�h"
w_sql = w_sql & "   AND d.���[�J�[�R�[�h = b.���[�J�[�R�[�h"
w_sql = w_sql & "   AND d.���i�R�[�h = b.���i�R�[�h"
w_sql = w_sql & "   AND d.�F = b.�F"
w_sql = w_sql & "   AND d.�K�i = b.�K�i"
w_sql = w_sql & " ORDER BY b.�󒍖��הԍ�"

'@@@@@@response.write(w_sql)

Set RS_order = Server.CreateObject("ADODB.Recordset")
RS_order.Open w_sql, Connection, adOpenStatic

if RS_order.EOF = false then
'---- �w�b�_���Z�b�g
	payment_method = RS_order("�x�����@")
	furikomi_nm = RS_order("�U�����`�l")

	loan_downpayment_fl = RS_order("���[����������t���O")
	loan_downpayment_am = RS_order("���[������")
	loan_term = RS_order("��]���[����")
	loan_am = RS_order("���[�����z")
	loan_apply_fl = RS_order("�I�����C�����[���\���t���O")
	loan_company = RS_order("���[�����")

	if ship_address_no = "" then
		ship_address_no = RS_order("�͐�Z���A��")
	end if

	ship_name = RS_order("�͐於�O")
	ship_zip = RS_order("�͐�X�֔ԍ�")
	ship_prefecture = RS_order("�͐�s���{��")
	ship_address = RS_order("�͐�Z��")
	ship_telephone = RS_order("�͐�d�b�ԍ�")
	ship_invoice_fl = RS_order("�͐�[�i�����t�t���O")

	freight_forwarder = RS_order("�^����ЃR�[�h")
	if freight_forwarder = "" then
		freight_forwarder = "1"		'���� �����l
	end if

	if isNull(RS_order("�w��[��")) = false then
		delivery_mm = cf_NumToChar(DatePart("m", RS_order("�w��[��")),2)
		delivery_dd = cf_NumToChar(DatePart("d", RS_order("�w��[��")),2)
	end if

	delivery_tm = RS_order("���Ԏw��")

	eigyousho_dome_fl = RS_order("�c�Ə��~�߃t���O")
	ikkatsu_fl = RS_order("�ꊇ�o�׃t���O")

	payment_method = RS_order("�x�����@")

	receipt_fl = RS_order("�̎������s�t���O")
	receipt_nm = RS_order("�̎�������")

	if receipt_fl = "Y" then
		if receipt_nm = "" then
			receipt_nm = customer_nm
		end if
		receipt_memo = RS_order("�̎����A������")
		if receipt_memo = "" then
			receipt_memo = "�����@���Ƃ���"
		end if
	end if
	receipt_nm_org = customer_nm
	receipt_memo_org = "�����@���Ƃ���"

	RebateFl = RS_order("���x�[�g�g�p�t���O")

end if

End function

'========================================================================
'
'	Function	�������i�ꗗHTML�쐬
'
'========================================================================
'
Function CreateOrderHtml()

Dim v_dataCnt
Dim v_product_nm
Dim vTotalAm

v_dataCnt = 0
vTotalAm = 0
w_html = ""
i_toriyose_fl = ""
i_tokuchuu_fl = ""
i_daibiki_fuka_fl = ""

'---- ����HTML�쐬
if RS_order.EOF = true then
	w_html = w_html & "<table width='100%' border='0' cellspacing='1' cellpadding='0'>" & vbNewLine
	w_html = w_html & "<tr class='honbun'><td align='center'><b>�J�[�g�ɏ��i������܂���B</b></td></tr>" & vbNewLine
	w_html = w_html & "</table>" & vbNewLine
	wOrderProductHTML = w_html
	wNoData = true
	exit function
end if

'----- ���o��
w_html = w_html & "<table width='100%' border='0' cellspacing='1' cellpadding='0'>" & vbNewLine
w_html = w_html & "  <tr align='center' bgcolor='#d3d3d3' class='honbun'>" & vbNewLine
w_html = w_html & "    <td>���[�J�[</td>" & vbNewLine
w_html = w_html & "    <td>���i��</td>" & vbNewLine
w_html = w_html & "    <td>�P��(�ō�)</td>" & vbNewLine
w_html = w_html & "    <td>����</td>" & vbNewLine
w_html = w_html & "    <td>���z(�ō�)</td>" & vbNewLine
w_html = w_html & "  </tr>" & vbNewLine

Do Until RS_order.EOF = true
	'------------- ���[�J�[�A���i��
	v_product_nm = RS_order("���i��")
	if Trim(RS_order("�F")) <> "" then
		v_product_nm = v_product_nm & "/" & RS_order("�F")
	end if
	if Trim(RS_order("�K�i")) <> "" then
		v_product_nm = v_product_nm & "/" & RS_order("�K�i")
	end if
	w_html = w_html & "  <tr>" & vbNewLine
	w_html = w_html & "    <td align='left' width='170' nowrap class='honbun'>" & RS_order("���[�J�[��") & "</td>" & vbNewLine
	w_html = w_html & "    <td align='left' nowrap><a href='" & g_HTTP & "shop/ProductDetail.asp?Item=" & RS_order("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS_order("���i�R�[�h")) & "^" & RS_order("�F") & "^" & RS_order("�K�i") & "' class='link'>" & v_product_nm & "</a></td>" & vbNewLine

		'------------- �P���A���ʁA���z
	wPrice = calcPrice(RS_order("�󒍒P��"), wSalesTaxRate)
	vTotalAm = vTotalAm + (wPrice * RS_order("�󒍐���"))

	w_html = w_html & "    <td align='right' width='100' class='honbun'>" & FormatNumber(wPrice,0) & "�~</td>" & vbNewLine
	w_html = w_html & "    <td align='right' width='70' class='honbun'>" & RS_order("�󒍐���") & "</td>" & vbNewLine
	w_html = w_html & "    <td align='right' width='130' class='honbun'>" & FormatNumber(wPrice*RS_order("�󒍐���"),0) & "�~</td>" & vbNewLine
	w_html = w_html & "   </tr>" & vbNewLine

'----���̑����Z�b�g
	if RS_order("�����\����") <= 0 then		'�v����
		i_toriyose_fl = "Y"
	end if
	if RS_order("���[�J�[�������敪") = "����" then		'���ʒ���
		i_toriyose_fl = "Y"
		i_tokuchuu_fl = "Y"
	end if
	if RS_order("����s�t���O") = "Y" then		'������s��
		i_daibiki_fuka_fl = "Y"
	end if

	v_dataCnt = v_dataCnt + 1
	RS_order.MoveNext
Loop

'----���i���v���z
w_html = w_html & "  <tr bgcolor='#d3d3d3' class='honbun'>" & vbNewLine
w_html = w_html & "    <td height='2' colspan='5' align='left'><img src='images/blank.gif' width='1' height='2'></td>" & vbNewLine
w_html = w_html & "  </tr>" & vbNewLine
 
w_html = w_html & "  <tr class='honbun'>" & vbNewLine
w_html = w_html & "    <td align='left'><a href='" & g_HTTP & "shop/Order.asp'><img src='images/OrderItemUpdate.gif' width='120' height='19' border='0' align='absmiddle' alt='�������i�̕ύX'></a></td>" & vbNewLine
w_html = w_html & "    <td align='left'></td>" & vbNewLine
w_html = w_html & "    <td colspan='2' align='right'><b>���i���v(�ō�)</b></td>" & vbNewLine
w_html = w_html & "    <td align='right'><b>" & FormatNumber(vTotalAm,0) & "�~</b></td>" & vbNewLine
w_html = w_html & "  </tr>" & vbNewLine

'----���x�[�g���z�\��
if KabusokuAm > 0 AND CustomerClass = "��ʌڋq"then
	w_html = w_html & "  <tr bgcolor='#d3d3d3' class='honbun'>" & vbNewLine
	w_html = w_html & "    <td height='2' colspan='5' align='left'><img src='images/blank.gif' width='1' height='2'></td>" & vbNewLine
	w_html = w_html & "  </tr>" & vbNewLine
	 
	w_html = w_html & "  <tr class='honbun'>" & vbNewLine
	w_html = w_html & "    <td align='left'></td>" & vbNewLine
	w_html = w_html & "    <td align='left'></td>" & vbNewLine
	w_html = w_html & "    <td colspan='2' align='right'><b>�N���W�b�g/�ߕs����</b></td>" & vbNewLine
	w_html = w_html & "    <td align='right'><b>" & FormatNumber(KabusokuAm,0) & "�~</b></td>" & vbNewLine
	w_html = w_html & "  </tr>" & vbNewLine

	w_html = w_html & "  <tr class='honbun'>" & vbNewLine
	w_html = w_html & "    <td align='left'></td>" & vbNewLine
	w_html = w_html & "    <td align='left'></td>" & vbNewLine
	w_html = w_html & "    <td colspan='3' align='left'><input type='checkbox' name='RebateFl' value='Y' "

	if RebateFl = "Y" then
		w_html = w_html & "CHECKED"
	end if

	w_html = w_html & "><b>���x�����ɃN���W�b�g/�ߕs�������g�p����</b>" & vbNewLine
	w_html = w_html & "  </tr>" & vbNewLine
end if

w_html = w_html & "</table>" & vbNewLine

if v_dataCnt = 1 then
	i_toriyose_fl = "N"		'�f�[�^��1�������Ȃ��ꍇ�͎�񂹎��̈ꊇ���b�Z�[�W�s�v
end if

RS_order.close
wOrderProductHTML = w_html

End Function

'========================================================================
'
'	Function	�ڋq�͐���̎��o��
'
'========================================================================
'
Function get_todokesaki()

'---- �ڋq�͐�����o��
w_sql = ""
w_sql = w_sql & "SELECT b.�Z���A��" 
w_sql = w_sql & "       , b.�Z������" 
w_sql = w_sql & "       , b.�ڋq�X�֔ԍ�" 
w_sql = w_sql & "       , b.�ڋq�s���{��" 
w_sql = w_sql & "       , b.�ڋq�Z��" 
w_sql = w_sql & "       , c.�ڋq�d�b�ԍ�" 
w_sql = w_sql & "  FROM Web�ڋq�Z�� b WITH (NOLOCK)" 
w_sql = w_sql & "     , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)" 
w_sql = w_sql & " WHERE b.�ڋq�ԍ� = " & userID 
w_sql = w_sql & "   AND c.�ڋq�ԍ� = b.�ڋq�ԍ�" 
w_sql = w_sql & "   AND c.�Z���A�� = b.�Z���A��" 
w_sql = w_sql & "   AND c.�d�b�敪 = '�d�b'"
w_sql = w_sql & " ORDER BY b.�Z���A��"
	  
'@@@@@@response.write(w_sql)

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open w_sql, Connection, adOpenStatic

wShipAddressHTML = ""

Do while RS_customer.EOF = false
	wShipAddressHTML = wShipAddressHTML _
							& "<option value='" & RS_customer("�Z���A��") & "'>" _
							& RS_customer("�Z������") _
							& "�@��" & RS_customer("�ڋq�X�֔ԍ�") _
							& " " & RS_customer("�ڋq�s���{��") & RS_customer("�ڋq�Z��") _
							& "�@" & RS_customer("�ڋq�d�b�ԍ�") & vbNewLine
	
	RS_customer.MoveNext
Loop

RS_customer.close

wShipAddressHTML = "<select name='ship_address_no'>" & vbNewLine _
									& wShipAddressHTML _
									& "</select>"

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close

End function

'========================================================================
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>�T�E���h�n�E�X  ����</title>

<!--#include file="../Navi/NaviStyle.inc"-->

<script language="JavaScript">
//=====================================================================
//	���� onClick
//=====================================================================
function order_onClick(pEstimate){

	document.f_data.estimate_fl.value = pEstimate;
	document.f_data.action = "OrderInfoInsert.asp";
	document.f_data.submit();
}

//=====================================================================
//	�͐�ύX onClick
//=====================================================================
function ship_address_onClick(){

	if (document.f_data.ship_address_no.selectedIndex == 0){
		document.f_data.action = "../member/Member.asp?called_from=order";
	}else{
		document.f_data.action = "../member/MemberShipaddress.asp?called_from=order";
	}
 
	document.f_data.submit();
}

//=====================================================================
//	�Z������ onClick
//=====================================================================
function address_search_onClick(){

	var addrWin;

	if (document.f_data.ship_zip.value == ""){
		alert("�X�֔ԍ�����͂��ĉ������B");
		return;
	}
 
	AddrWin = window.open("../comasp/address_search.asp?zip=" + document.f_data.ship_zip.value +"&name_prefecture=i_ship_prefecture&name_address=ship_address","AddrSearch","width=200,height=100");

}

//=====================================================================
//	�J�[�h
//=====================================================================
function card_onClick(){

	//��]���[�������N���A
	document.f_data.loan_downpayment_am.value = "";
	document.f_data.loan_term.options[0].selected = true;
	document.f_data.loan_am.value = "";
}

//=====================================================================
//	���[��
//=====================================================================
function loan_onClick(){

//	receipt_onClick();
}

//=====================================================================
//	��]���[����
//=====================================================================
function loan_term_onChange(){

	//�J�[�h�x�����̏ꍇ�I��s��
	if (document.f_data.payment_method[3].checked == true){
		alert("�J�[�h�ł��w���̍ہC��]���[���񐔂̐ݒ�͂ł��܂���B");
		document.f_data.loan_term.options[0].selected = true;
	}
}

//=====================================================================
//	���W�I�{�^���A�h���b�v�_�E�����X�g���ȑO�ɑI��������Ԃɂ���
//=====================================================================
function preset_values(pPref){

// �Z��������������̌Ăяo�����͓s���{���݂̂��Z�b�g
	if (pPref == "pref"){
		for (var i=0; i<document.f_data.ship_prefecture.options.length; i++){
			if (document.f_data.ship_prefecture.options[i].value == document.f_data.i_ship_prefecture.value)		{
				document.f_data.ship_prefecture.options[i].selected = true;
				break;
			}
		}
		return;
	}

//	�x�����@
	for (var i=0; i<document.f_data.payment_method.length; i++){
		if (document.f_data.payment_method[i].value == document.f_data.i_payment_method.value){
			document.f_data.payment_method[i].checked = true;
			break;
		}
	}

// ���[������
	if (document.f_data.i_payment_method.value == "���[��"){
		if (document.f_data.i_loan_downpayment_fl.value == "Y"){
			document.f_data.loan_downpayment_fl[1].checked = true;
		}
		if (document.f_data.i_loan_downpayment_fl.value == "N"){
			document.f_data.loan_downpayment_fl[0].checked = true;
		}
	}

// ���[����/���z
	if (document.f_data.i_payment_method.value == "���[��"){
		if (document.f_data.loan_am.value != "0"){
			document.f_data.loan_term_payment[1].checked = true;
		}
	}

// ���[����
		for (var i=0; i<document.f_data.loan_term.options.length; i++){
			if (document.f_data.loan_term.options[i].value == document.f_data.i_loan_term.value){
				document.f_data.loan_term.options[i].selected = true;
				break;
			}
		}

//	�I�����C�����[���\��		030829 add
	if (document.f_data.i_payment_method.value == "���[��"){
		if (document.f_data.i_loan_apply_fl.value == "Y"){
			document.f_data.loan_apply_fl[0].checked = true;
		}
		if (document.f_data.i_loan_apply_fl.value == "N"){
			document.f_data.loan_apply_fl[1].checked = true;
		}
	}

//	�I�����C�����[�����
	if (document.f_data.i_payment_method.value == "���[��"){
		if (document.f_data.i_loan_apply_fl.value == "Y"){
			if (document.f_data.i_loan_company.value == "�Z���g����"){
				document.f_data.loan_company[0].checked = true;
			}
			if (document.f_data.i_loan_company.value == "�W���b�N�X"){
				document.f_data.loan_company[1].checked = true;
			}
		}
	}

// �͐�ꗗ
	for (var i=0; i<document.f_data.ship_address_no.options.length; i++){
		if (document.f_data.ship_address_no.options[i].value == document.f_data.i_ship_address_no.value){
			document.f_data.ship_address_no.options[i].selected = true;
			break;
		}
	}

// �s���{��
	for (var i=0; i<document.f_data.ship_prefecture.options.length; i++){
		if (document.f_data.ship_prefecture.options[i].value == document.f_data.i_ship_prefecture.value)		{
			document.f_data.ship_prefecture.options[i].selected = true;
			break;
		}
	}

// �[�i�����t
	if (document.f_data.i_ship_invoice_fl.value == "Y"){
		document.f_data.ship_invoice_fl[0].checked = true;
	}
	if (document.f_data.i_ship_invoice_fl.value == "N"){
		document.f_data.ship_invoice_fl[1].checked = true;
	}

//	�^�����
	for (var i=0; i<document.f_data.freight_forwarder.options.length; i++){
		if (document.f_data.freight_forwarder.options[i].value == document.f_data.i_freight_forwarder.value){
			document.f_data.freight_forwarder.options[i].selected = true;
			break;
		}
	}

//	�z�B��
	freight_forwarder_onChange();			

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

//	���Ԏw��
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
	if (document.f_data.i_toriyose_fl.value == "Y"){
		if (document.f_data.ikkatsu_fl.length >= 2){
			if (document.f_data.i_ikkatsu_fl.value == "Y"){
				document.f_data.ikkatsu_fl[0].checked = true;
			}
			if (document.f_data.i_ikkatsu_fl.value == "N"){
				document.f_data.ikkatsu_fl[1].checked = true;
			}
		}
	}

// �̎���
	if (document.f_data.receipt_fl.type != "hidden"){
		if (document.f_data.i_receipt_fl.value == "N"){
			document.f_data.receipt_fl[0].checked = true;
		}
		if (document.f_data.i_receipt_fl.value == "Y"){
			document.f_data.receipt_fl[1].checked = true;
		}
	}

}

//=====================================================================
//	�^����Ђ�ύX���ꂽ��A���ԑюw��h���b�v�_�E�����^����Ђɍ��킹�ĕύX
//=====================================================================
function freight_forwarder_onChange(){

	for (var i=0; i<document.f_data.freight_forwarder.options.length; i++){
		if (document.f_data.freight_forwarder.options[i].selected == true){
			if (document.f_data.freight_forwarder.options[i].text == "����}��"){
				document.f_data.delivery_tm.options.length = 6;
				document.f_data.delivery_tm.options[0].value = "";
				document.f_data.delivery_tm.options[1].value = "�ߑO��";
				document.f_data.delivery_tm.options[2].value = "12������14���܂�";
				document.f_data.delivery_tm.options[3].value = "14������16���܂�";
				document.f_data.delivery_tm.options[4].value = "16������18���܂�";
				document.f_data.delivery_tm.options[5].value = "18������21���܂�";
				document.f_data.delivery_tm.options[0].text = "";
				document.f_data.delivery_tm.options[1].text = "�ߑO��";
				document.f_data.delivery_tm.options[2].text = "12������14���܂�";
				document.f_data.delivery_tm.options[3].text = "14������16���܂�";
				document.f_data.delivery_tm.options[4].text = "16������18���܂�";
				document.f_data.delivery_tm.options[5].text = "18������21���܂�";
			}
			if (document.f_data.freight_forwarder.options[i].text == "���}�g�^�A"){
				document.f_data.delivery_tm.options.length = 7;
				document.f_data.delivery_tm.options[0].value = "";
				document.f_data.delivery_tm.options[1].value = "�ߑO��";
				document.f_data.delivery_tm.options[2].value = "12������14��";
				document.f_data.delivery_tm.options[3].value = "14������16��";
				document.f_data.delivery_tm.options[4].value = "16������18��";
				document.f_data.delivery_tm.options[5].value = "18������20��";
				document.f_data.delivery_tm.options[6].value = "20������21��";
				document.f_data.delivery_tm.options[0].text = "";
				document.f_data.delivery_tm.options[1].text = "�ߑO��";
				document.f_data.delivery_tm.options[2].text = "12������14��";
				document.f_data.delivery_tm.options[3].text = "14������16��";
				document.f_data.delivery_tm.options[4].text = "16������18��";
				document.f_data.delivery_tm.options[5].text = "18������20��";
				document.f_data.delivery_tm.options[6].text = "20������21��";
			}
		}
	}
	document.f_data.delivery_tm.options[0].selected = true;
}

</script>

</head>

<body background="../Navi/Images/back_ground.gif" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!--#include file="../Navi/NaviTop.inc"-->

<table width="940" height="26" border="0" cellpadding="0" cellspacing="0">
  <tr>

<!--#include file="../Navi/NaviLeft.inc"-->

    <td width="798" align="left" valign="top" bgcolor="#ffffff">

<!------------ �y�[�W���C�������̋L�q START ------------>

<!-- �G���[���b�Z�[�W -->
<% if msg <> "" then %>

<table width="99%" border="1" cellspacing="0" cellpadding="3" bordercolor="#999999" bordercolorlight="#999999" bordercolordark="#ffffff">
  <tr align="center" valign="top">
    <td align="left" bgcolor="#D2FFFF">
      <font color = "#ff0000">
      <%=wMSG%>
      </font>

	<% if msg = "CardError1" OR msg = "CardError2" then %>
      <b>���悭����J�[�h�G���[�ɂ���</b><br>

      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td bgcolor="#000000">
            <table width="100%" border="0" cellpadding="2" cellspacing="1" class="ctTable">
              <tr>
                <td bgcolor="#CCCCCC" class="honbun"><p class="bold">�G���[�R�[�h</p></td>
                <td bgcolor="#CCCCCC" class="honbun"><p class="bold">���R�E�Ώ����@</p></td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S0010G12</td>
                <td bgcolor="#FFFFFF" class="honbun">���̃J�[�h�͂����p�ł��܂���ł����B<br>�����p�ł��Ȃ��ڍח��R�Ɋւ��܂��ẮA�J�[�h��Ђւ��₢���킹���������B</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S0010G65</td>
                <td bgcolor="#FFFFFF" class="honbun">���͂��ꂽ�J�[�h�ԍ��Ɍ�肪����\��������܂��B<br>�ēx���͂���邩�A�J�[�h��Ђւ��₢���킹���������B</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S0010G83</td>
                <td bgcolor="#FFFFFF" class="honbun">���͂��ꂽ�L�������Ɍ�肪����\��������܂��B<br>�ēx���͂���邩�A�J�[�h��Ђւ��₢���킹���������B</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S102000C</td>
                <td bgcolor="#FFFFFF" class="honbun">3D�Z�L���A�F�ؒ��ɃL�����Z�������ꂽ���A���͂��ꂽ�p�X���[�h���F�؂ł��܂���ł����B</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S20210A2</td>
                <td bgcolor="#FFFFFF" class="honbun">�J�[�h�ԍ�������͂��ꂽ�\��������܂��B<br>�ēx���͂���邩�A�J�[�h��Ђւ��₢���킹���������B</td>
              </tr>
              <tr>
                <td bgcolor="#FFFFFF" class="honbun">S2022017</td>
                <td bgcolor="#FFFFFF" class="honbun">�������̍ہA��莞�Ԃ��o�߂����ׁA�^�C���A�E�g���ꂽ�\�����l�����܂��B<br>���萔�ł����A�͂��߂����蒼���Ă��������B</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
	<% end if %>

    </td>
  </tr>
</table>

<% end if %>

      <form method="post" name="f_data">
      <table width="790" border="0" cellspacing="0" cellpadding="2">
        <tr class="honbun">
          <td width="5" height="5"></td>
          <td></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top"><b><%=customer_nm%> �l</b></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
          <font color="#666666">
          �E�m�F���[���͎����I�ɑ��M����܂��B�����_��͏��i�̔����������Đ����ƂȂ�܂��B<br>
          �E�g�ѓd�b�Ȃǂ̎�M����������A�h���X�ł��o�^���ꂽ�ꍇ�A��������񂪎�M�ł��Ȃ��ꍇ���������܂��B���炩���߂������������B<br>
          �E���������i�ɂ��Ă̖⍇���́A���[���₨�d�b�ɂď����Ă���܂��̂ł������O�ɂ��m�F�����܂��l���肢���܂��B<br>
          �E�J�[�g�ɓ��ꂽ���i�ȊO������]�̍ۂ́A�\�߃��[���₨�d�b�ɂĂ��⍇���������B
          </font>
          </td>
        </tr>

<!---- �������e -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">�������e�̊m�F</span></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top">

<!--�������i�ꗗ-->
<%=wOrderProductHTML%>

          </td>
        </tr>

<!---- �ڋq��� -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">���q�l���̊m�F</span></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="770" border="0" cellspacing="1" cellpadding="0">
              <tr class="honbun">
                <td align="left" valign="top"><b>�����O</b></td>
                <td align="left" valign="top"><%=customer_nm%></td>
              </tr class="honbun">
              <tr class="honbun">
                <td align="left" valign="top"><b>�t���K�i</b></td>
                <td align="left" valign="top"><%=furigana%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>��</b></td>
                <td align="left" valign="top"><%=zip%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>�Z��</b></td>
                <td align="left" valign="top"><%=prefecture%><%=address%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>�d�b�ԍ�</b></td>
                <td align="left" valign="top"><%=telephone%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>FAX�ԍ�</b></td>
                <td align="left" valign="top"><%=fax%></td>
              </tr>
              <tr class="honbun">
                <td align="left" valign="top"><b>e-mail</b></td>
                <td align="left" valign="top"><%=customer_email%></td>
              </tr>
            </table>
          </td>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top" colspan="2"><a href="../member/member.asp?called_from=order"><img src="images/MemberUpdate.gif" width="120" height="19" border="0" alt='���q�l���̕ύX'></a></td>
        </tr>

<!---- ���x�������@ -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">���x�����@�̑I��</span></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="payment_method" value="��s�U��"><b>��s�U��</b>�@�U���l���`
            <input type="text" name="furikomi_nm" size=45  maxlength=60 value="<%=furikomi_nm%>"><br>
            <img src="images/blank.gif" alt="" width="85" height="5" align="left"><font color='#666666'>�i�U�����`�����q�l�̂����O�ƈقȂ�ꍇ�݂̂��L���������B)</font>
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="payment_method" value="�R���r�j�x��"><b>�l�b�g�o���L���O�E�䂤����E�R���r�j����</b>
<table class="honbun">
  <tr>
    <td width=30></td>
    <td><font color='#666666'>�E���[�\���A�t�@�~���[�}�[�g�A�T�[�N��K�A�T���N�X�A�Z�C�R�[�}�[�g�A�䂤�����s�A�l�b�g�o���L���O �ł��x�����������܂��</font></td>
  </tr>
  <tr>
    <td width=30></td>
    <td><font color='#666666'>�E��������A���z���ύX�ƂȂ邲�����̕ύX�͏��鎖���o���܂���B�݌ɂ̖������i�Ȃǂ��������̍ۂ́A���O�ɂ��⍇�����������B<br>
�EE-MAIL�A�h���X���g�т̏ꍇ�́A�K�v�������m�F�ł��Ȃ��ꍇ������ׁA�p�\�R������̂����p���������߂��܂��B</font></td>
  </tr>
  <tr>
    <td width=30></td>
    <td><font color='#ff0000'>�E��قǁA�����Ϗ������ē��v���܂��̂ŁA����������m�F��ɂ��U���ݒ����܂��l���肢�v���܂��B</font></td>
  </tr>
</table>
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="payment_method" value="�����"><b>�������</b>�@<font color='#666666'>��������ł̂��w���̏ꍇ�A���i�̔����͈ꊇ�o�ׂƂȂ�܂��B�܂��A���x�����͌����݂̂̎�t�ƂȂ�܂��B</font>
          </td>
        </tr>

        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top"><!-- <disabled="disabled"> -->
            <input type="radio" name="payment_method" value="�N���W�b�g�J�[�h" onClick="card_onClick();"><b>�N���W�b�g�J�[�h</b><br>
			<font color="#FF0000">���݃N���W�b�g�J�[�h�̗��p���~�����Ē����Ă���܂��B<br>
			�����p�̂��q�l�ɂ́A�����f�����|���������܂����A�������������������܂��l���肢�\���グ�܂��B</font>
			<br>
          </td>
        </tr>

		<tr class='honbun'>
		  <td width=30></td>
		  <td colspan=2><font color='#666666'>�E�N���W�b�g�J�[�h�ł̂��w���̏ꍇ�A�ꊇ�����݂̂̂��戵�ƂȂ�܂��B<br>
		  �E���{�l���`�̃J�[�h�݂̂����p�����܂��B<br>
		  �E�J�[�h��Ђ̓o�^���e�ƍ��񂲓o�^�f�[�^�ɑ��Ⴊ�������ꍇ�A�����������邱�Ƃ��o���Ȃ��ꍇ������܂��B<br>
		  �E�N���W�b�g�J�[�h�͒������Ɍ��ς���܂��B�݌ɂ������A�[����������ꍇ�A���i�����͂�����O�ɑ�����������Ƃ���鎖������܂��B<br>�@���炩���߂��������������B</font></td>
		</tr>

        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="payment_method" value="���[��" onClick="loan_onClick();"><b>���[��</b>�@
            <input type="radio" name="loan_downpayment_fl" value="N">���������@/�@
            <input type="radio" name="loan_downpayment_fl" value="Y">��������@ ����
            <input type="text" name="loan_downpayment_am" size=10 maxlength=6 value="<%=loan_downpayment_am%>">�~
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <img src="images/blank.gif" alt="" width="20" height="20" align="left">
            <input type="radio" name="loan_apply_fl" value="Y">�I�����C���Ń��[����\�����ށB

<!--        <input type="hidden" name="loan_company" value="�Z���g����">  -->

            <img src="images/blank.gif" alt="" width="40" height="20" align="left">

            <input type="radio" name="loan_company" value="�W���b�N�X" checked>�W���b�N�X�@
            <input type="radio" name="loan_company" value="�Z���g����">�Z�f�B�i�@<br>
            <font color='#ff0000'>�I�����C�����[���̏ꍇ����\����̂��������e�̕ύX�����邱�Ƃ��ł��܂���B<br>
            ���������e�Ƥ�I�����C�����[���\���t�H�[���̓��e�����m�F�̏ゲ�������������B</font>
          </td>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <img src="images/blank.gif" alt="" width="20" height="20" align="left">
            <input type="radio" name="loan_apply_fl" value="N">�I�����C�����g�p���Ȃ��B(���[���񐔂܂��͌��z���w��肢�܂��B) 
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <img src="images/blank.gif" alt="" width="40" height="20" align="left">
            <input type="radio" name="loan_term_payment" value="T">��]���[����
            <select name="loan_term" onChange="loan_term_onChange();">
              <option value="0">
              <option value="1">1
              <option value="2">2
              <option value="3">3
              <option value="6">6
              <option value="10">10
              <option value="12">12
              <option value="15">15
              <option value="18">18
              <option value="20">20
              <option value="24">24
              <option value="30">30
              <option value="36">36
              <option value="42">42
              <option value="48">48
              <option value="54">54
              <option value="60">60
            </select>�@/�@
            <input type="radio" name="loan_term_payment" value="P">���z�x�����z
            <input type="text" name="loan_am" size=10 maxlength=6 value="<%=loan_am%>">�~
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <img src="images/blank.gif" alt="" width="40" height="20" align="left"><font color="#666666">�E���[����Ђɂ�育��]�̂��x�����񐔂��w��ł��Ȃ��ꍇ���������܂��B </font>
          </td>
        </tr>

<!---- ���̑��w�� -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">���̑��w��</span></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            ���o�^�Z���ȊO�ւ̂��͂��A�z�B���w��A�z�B���ԑюw��A�̎��؂̔��s�����ʂȂ��w�肪����ꍇ�́A�ȉ��̊�]�̍��ڂ���͂���[����]�{�^�������ĉ������B
          </td>
        </tr>
<% if wNoData = false then %>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <a href="JavaScript:order_onClick('N');"><img src="images/Order.gif" width="120" height="19" border="0" alt='������ʂ֐i��'></a>
          </td>
        </tr>
<% end if %>

<!---- �z����w��-------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">���͂���̕ύX</span></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">���͐���ꗗ�̒�����I������B�ꗗ�ɖ����ꍇ�͉��̗��֓��͂��ĉ������B </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top"><%=wShipAddressHTML%></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="600" border="0" cellspacing="2" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>�����O</b></td>
                <td><input type="text" name="ship_name" size=30 maxlength=60 value="<%=ship_name%>"></td>
              </tr>
              <tr class="honbun">
                <td width="80"><b>�Z��</b></td>
                <td>��<input type="text" name="ship_zip" size="10" maxlength="8" value="<%=ship_zip%>">�i���p�j<a href="JavaScript:address_search_onClick();"><img src="images/AddressSearch.gif" width="120" height="19" border="0" alt='�Z������'></a>&nbsp;�X�֔ԍ�����͂��ă{�^���������ĉ������</td>
              </tr>
              <tr class="honbun">
                <td width="80"></td>
                <td>
                  <select name="ship_prefecture">
                    <option value="" SELECTED>�s���{��
                    <option value="�k�C��">�k�C��
                    <option value="�X��">�X��
                    <option value="�H�c��">�H�c��
                    <option value="��茧">��茧
                    <option value="�{�錧">�{�錧
                    <option value="�R�`��">�R�`��
                    <option value="������">������
                    <option value="�Ȗ،�">�Ȗ،�
                    <option value="�V����">�V����
                    <option value="�Q�n��">�Q�n��
                    <option value="��ʌ�">��ʌ�
                    <option value="��錧">��錧
                    <option value="��t��">��t��
                    <option value="�����s">�����s
                    <option value="�_�ސ쌧">�_�ސ쌧
                    <option value="�R����">�R����
                    <option value="���쌧">���쌧
                    <option value="�򕌌�">�򕌌�
                    <option value="�x�R��">�x�R��
                    <option value="�ΐ쌧">�ΐ쌧
                    <option value="�É���">�É���
                    <option value="���m��">���m��
                    <option value="�O�d��">�O�d��
                    <option value="�ޗǌ�">�ޗǌ�
                    <option value="�a�̎R��">�a�̎R��
                    <option value="���䌧">���䌧
                    <option value="���ꌧ">���ꌧ
                    <option value="���s�{">���s�{
                    <option value="���{">���{
                    <option value="���Ɍ�">���Ɍ�
                    <option value="���R��">���R��
                    <option value="���挧">���挧
                    <option value="������">������
                    <option value="�L����">�L����
                    <option value="�R����">�R����
                    <option value="���쌧">���쌧
                    <option value="������">������
                    <option value="���Q��">���Q��
                    <option value="���m��">���m��
                    <option value="������">������
                    <option value="���ꌧ">���ꌧ
                    <option value="�啪��">�啪��
                    <option value="�F�{��">�F�{��
                    <option value="�{�茧">�{�茧
                    <option value="���茧">���茧
                    <option value="��������">��������
                    <option value="���ꌧ">���ꌧ
                  </select>
                  <input type="text" name="ship_address" size="60" maxlength="80" value="<%=ship_address%>"><br>��Ж��A�r�����A�����ԍ��A�����l���A���͖Y�ꂸ���L���������B
                </td>
              </tr>
              <tr class="honbun">
                <td width="80"><b>�d�b�ԍ�</b></td>
                <td><input type="text" name="ship_telephone" size="30" maxlength="20" value="<%=ship_telephone%>">�i���p�����j</td>
              </tr>
              <tr class="honbun">
                <td width="80"></td>
                <td align="left" valign="top">
                  <input type="radio" name="ship_invoice_fl" value="Y" checked>���͐�ɔ[�i���𑗕t���ėǂ��@�@�@
                  <input type="radio" name="ship_invoice_fl" value="N">���͐�ɔ[�i���𑗕t���Ȃ�
                </td>
              </tr>
            </table>          
          </td>
        </tr>

<!---- �z�����@�w�� -------------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">�^����ЁE�z�������̎w��</span></td>
        </tr>

        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
  <% if payment_method <> "�����" then %>
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>�o�׎w��</b></td>
                <td>
                  <input type="radio" name="ikkatsu_fl" value="Y" >���i���S�đ����Ă���ꊇ�o�ׂ���@/
                  <input type="radio" name="ikkatsu_fl" value="N" checked>�݌ɏ��i�݂̂��ɏo�ׂ���
                </td>
              </tr>
            </table>
  <% else %>
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>�o�׎w��</b></td>
                <td>
                  <input type="hidden" name="ikkatsu_fl" value="Y">���i���S�đ����Ă���ꊇ�o�ׂƂȂ�܂��
                </td>
              </tr>
            </table>
  <% end if %>
          </td>
        </tr>

        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="80" class="honbun"><b>�^�����</b></td>
                <td>
                  <span class="honbun">
                  <select name="freight_forwarder" onChange="freight_forwarder_onChange();">
                    <option value="1" SELECTED>����}��
                    <option value="2">���}�g�^�A
                  </select>
                  �z�����@�ɂ��Ă̏���</span><a href="<%=g_HTTP%>guide/kaimono.asp#haisou" target='_new' class='link'>�������</a><span class="honbun">������������� </span>
                </td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>�z�B��]��</b></td>
                <td>
                  <select name="delivery_mm">
                    <option value="" SELECTED>
                    <option value="01">1
                    <option value="02">2
                    <option value="03">3
                    <option value="04">4
                    <option value="05">5
                    <option value="06">6
                    <option value="07">7
                    <option value="08">8
                    <option value="09">9
                    <option value="10">10
                    <option value="11">11
                    <option value="12">12
                  </select>��
                  <select name="delivery_dd">
                    <option value="" SELECTED>
                    <option value="01">1
                    <option value="02">2
                    <option value="03">3
                    <option value="04">4
                    <option value="05">5
                    <option value="06">6
                    <option value="07">7
                    <option value="08">8
                    <option value="09">9
                    <option value="10">10
                    <option value="11">11
                    <option value="12">12
                    <option value="13">13
                    <option value="14">14
                    <option value="15">15
                    <option value="16">16
                    <option value="17">17
                    <option value="18">18
                    <option value="19">19
                    <option value="20">20
                    <option value="21">21
                    <option value="22">22
                    <option value="23">23
                    <option value="24">24
                    <option value="25">25
                    <option value="26">26
                    <option value="27">27
                    <option value="28">28
                    <option value="29">29
                    <option value="30">30
                    <option value="31">31
                  </select>
                </td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>��]���ԑ�</b></td>
                <td>
                  <select name="delivery_tm">
                    <option value="" SELECTED>
                  </select>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="700" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>�c�Ə��~��</b></td>
                <td>
                  <input type="checkbox" name="eigyousho_dome_fl" value="Y">�^����Љc�Ə��~�߁@<font color="#666666">(�c�Ə��~�߂̏ꍇ�A���͐�Z���̒S���c�Ə��ւ̗��ߒu���ƂȂ�܂��B)</font>
                </td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <font color="#666666">�E�V��A��ʏ󋵂Ȃ�тɔz�B�Ǝ҂̓s���ɂ�育��]�ɓY���Ȃ��ꍇ���������܂��B�\�߂��������������B<br>
          �E����}�ւ̏ꍇ�A���ԑюw��͕����Ȃ�тɂ��͂��悪�s�s���ɂ��Z�܂��̌l��̏ꍇ�Ɍ���\�ł��B<br>
          �E���͐悪���u�n�̏ꍇ�ͤ�����I�Ɉꊇ�o�ׂƂȂ�܂��B<br>
          �E�ꕔ���戵���ł��Ȃ��n�悪�������܂��B�ڍׂ͒S���c�Ƃɂ��⍇���������B</font>
          </td>
        </tr>

<!---- ���̑� ------------->
        <tr>
          <td>&nbsp;</td>
          <td align="left" valign="top"><span class="midashi">�̎���</span></td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <input type="radio" name="receipt_fl" value="N" checked><b>�s�v</b>�@
            <input type="radio" name="receipt_fl" value="Y"><b>�K�v</b>
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="500" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>�̎��؈���</b></td>
                <td><input type="text" name="receipt_nm" size=30 maxlength=60 value="<%=receipt_nm%>">�l</td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <table width="500" border="0" cellspacing="0" cellpadding="0">
              <tr class="honbun">
                <td width="80"><b>�A������</b></td>
                <td><input type="text" name="receipt_memo" size=30 maxlength=50 value="<%=receipt_memo%>"></td>
              </tr>
            </table>          
          </td>
        </tr>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <font color="#ff0000">�E���x�����@���ȉ��̏ꍇ�A�T�E���h�n�E�X�̗̎��؂͔��s�v���܂���B<br>
            1.�@�������<br>
            2.�@���[��<br>
            3.�@�R���r�j/�X�֋ǎx��<br>
            �E�����A�A�����́A�w�藓�ɓ��͒��������e�̂܂܍쐬�v���܂��B
          </td>
        </tr>
        <input type="hidden" name="receipt_nm_org" value="<%=receipt_nm_org%>">
        <input type="hidden" name="receipt_memo_org" value="<%=receipt_memo_org%>">

<!---- ���� -------------------->
<% if wNoData = false then %>
        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="left" valign="top">
            <a href="JavaScript:order_onClick('N');"><img src="images/Order.gif" width="120" height="19" border="0" alt='������ʂ֐i��'></a>
          </td>
        </tr>
<% end if %>

        <tr class="honbun">
          <td>&nbsp;</td>
          <td align="center" valign="top">&nbsp;</td>
        </tr>
      </table>

      <input type="hidden" name="estimate_fl" value="">
      <input type="hidden" name="customer_email" value="<%=customer_email%>">
      <input type="hidden" name="telephone" value="<%=telephone%>">
      <input type="hidden" name="i_payment_method" value="<%=payment_method%>">
      <input type="hidden" name="i_loan_downpayment_fl" value="<%=loan_downpayment_fl%>">
      <input type="hidden" name="i_loan_term" value="<%=loan_term%>">
      <input type="hidden" name="i_loan_apply_fl" value="<%=loan_apply_fl%>">
      <input type="hidden" name="i_loan_company" value="<%=loan_company%>">
      <input type="hidden" name="i_ship_address_no" value="<%=ship_address_no%>">
      <input type="hidden" name="i_ship_prefecture" value="<%=ship_prefecture%>">
      <input type="hidden" name="i_ship_invoice_fl" value="<%=ship_invoice_fl%>">
      <input type="hidden" name="i_freight_forwarder" value="<%=freight_forwarder%>">
      <input type="hidden" name="i_delivery_mm" value="<%=delivery_mm%>">
      <input type="hidden" name="i_delivery_dd" value="<%=delivery_dd%>">
      <input type="hidden" name="i_delivery_tm" value="<%=delivery_tm%>">
      <input type="hidden" name="i_eigyousho_dome_fl" value="<%=eigyousho_dome_fl%>">
      <input type="hidden" name="i_ikkatsu_fl" value="<%=ikkatsu_fl%>">
      <input type="hidden" name="i_receipt_fl" value="<%=receipt_fl%>">
      <input type="hidden" name="i_tokuchuu_fl" value="<%=i_tokuchuu_fl%>">
      <input type="hidden" name="i_toriyose_fl" value="<%=i_toriyose_fl%>">
      <input type="hidden" name="i_daibiki_fuka_fl" value="<%=i_daibiki_fuka_fl%>">
      </form>

<!------------ �y�[�W���C�������̋L�q END ------------>

    </td>
  </tr>
</table>

<!--#include file="../Navi/NaviBottom.inc"-->

<!--#include file="../Navi/NaviScript.inc"-->

</body>
</html>

<script language="JavaScript">

	preset_values();

</script>
