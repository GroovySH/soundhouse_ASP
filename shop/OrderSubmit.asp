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
'	�I�[�_�[�o�^�E���M����
'
'------------------------------------------------------------------------
'	
'		�͐悪�V�K�Ɏw�肳�ꂽ�Ƃ��͌ڋq���ɒǉ�
'		�����̏ꍇ�@����=����*�����|��
'		���󒍏���Web�󒍏��փR�s�[���A���󒍏����폜�B
'		�J�[�h�I�[�_�[�̏ꍇ�͗^�M�m�F���L���s���B
'		�I�[�_�[��t���[���̑��M�B�i�ڋq & �V���b�v)
'
'------------------------------------------------------------------------
'�X�V����
'2004/12/20 �������[�����͕ύX
'2004/12/20 �J�[�h�L��������/���Ȃ��f�[�^�̑΍�
'2004/12/21 Thanks�y�[�W�ւ̖߂��HTTPS�ɕύX
'2004/12/27 �����m�F��(OrderConfirm)��Window���珤�i�ǉ�����đ��M�{�^���������ꂽ���̑Ώ�
'2005/04/05 �J�[�h�����󒍃f�[�^������o���悤�ɕύX
'2005/05/13 OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"���Z�b�g
'2005/06/20 �I���R���[���ǉ�
'2005/08/24 �݌ɏ��\�������ʊ֐��g�p�ɕύX���A�󒍐���<�����\���ʂ̏ꍇ�͌ڋq�ɂ��݌ɐ���\��
'2005/08/31 �����ԐM���[������P���𔲂�
'2005/09/28 ���[���T�u�W�F�N�g�Ɏx�����@��ǉ�
'2005/09/29 ���[���T�u�W�F�N�g�ɃR���^�N�g�Ǘ��p����ǉ�
'2005/10/07 SH���ă��[�����M����SH�Ƃ���
'2006/06/22 Thanks.asp�Ăяo���p�����[�^��URLEncode�ǉ�
'2006/07/05 BlueGate�I�[�\������Ăяo�����́A�󒍔ԍ����n�����B
'2006/10/24 eContext����̌Ăяo�����́A�U�荞�ݕ[URL�A���ϑI��pURL���n�����(�R���r�j�x��)
'2006/11/27 �󒍋��z�ɃR���r�j�x�����萔�����݂ɕύX
'2006/11/30 �R���r�j�x�������[�����͕ύX
'2007/01/11 �R���r�j�x�������AeContext�x�����@URL�EeContext�U���[URL��Web�󒍂ɓo�^
'2007/01/12 �u�R���r�j�x���v���u�R���r�j/�X�֋ǎx���v�ɕ\���ύX
'2007/01/30 ���[���w�b�_���͕ύX
'2007/02/28 �󒍌`�Ԃ�ǉ�
'2007/04/20 �F�K�i�ʍ݌ɂ̈����\���ʂ��X�V�A���i�̊��������Z�b�g
'2007/05/09 ������A�N���W�b�g�̂Ƃ��́A���z�����[���ɕ\��
'2008/04/14 ���x�[�g�@�\�ǉ��A�J�[�h�����o�������폜�i�ڋq�ߕs�����X�V���}�C�i�X�̓G���[�F�s���ڋq�΍�j
'2008/04/21 �������ɓ͂�����Z�b�g
'2008/05/07 �����ҏ���Web�󒍂ɃZ�b�g
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2008/09/16 �R���r�j�x���\�����R���r�j�x��", "�l�b�g�o���L���O�E�䂤����E�R���r�j�����ɕύX
'2008/12/12 Email�A�ڋq�ɂ͍݌ɐ��ʔ�ʒm�ɁBShop�̓Z�b�g�i�ȊO�݌ɐ��ʒʒm + �����C��
'2009/04/30 �G���[����error.asp�ֈړ�
'2009/06/17 �p�ԂŐF�K�i������Ƃ��A�ŏ��̐F�̈����\�݌�=0�Ŋ��������Z�b�g���������C��
'2009/12/07 an�u�݌ɋH���v���u�݌ɋ͏��v�ɕύX
'2009/12/17 hn ���R�����h�p�ύX�i���i�w�����O�o�́j���R�����g�A�E�g
'2010/03/04 an ���R�����h�p�ύX�i���i�w�����O�o�́j�L����
'2010/05/07 an �[���\�肪�uXX/XX���\��v�̏ꍇ�́A"��������܂�"�����Ȃ��悤�ɏC��
'2010/08/11 an �����������[�����M����ShopBCC��BCC����悤�ɏC��
'2010/12/20 hn ���ς���̎���������P���̏ꍇ�͌�����󒍍ϐ��ʂ��X�V����
'2011/01/28 GV(ay) �͐�o�^�����̍폜
'2011/04/14 hn SessionID�֘A�ύX
'2011/06/01 if-web �������M���[���̉^����Е\���������폜
'2011/06/29 an #867 ����A���}�g�̏ꍇ�͎��Ԏw����e�Ђɉ����ēǂݑւ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/09/09 an #1123 �����ԐM���[���C�����󒍐����݌ɐ������Ȃ��Ƃ��́u�ꕔ�݌ɂ�����܂���v�ƕ\��
'2012/01/10 an ���R�����h���i�w�����O�o�͒�~
'2012/01/23 hn �󒍖��ׂɃf�[�^���Ȃ����̓G���[�Ƃ���
'2012/08/15 nt �Z�b�g�i���̔z�M���[�����e�s���i�݌ɏ󋵁j���C��
'2012/09/25 nt �̎��؈���E�A��������ύX
'2013/07/30 GV #1618 �A�t�B���G�C�g�d�����M�Ή�
'
'========================================================================

On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

Dim userID
Dim userName
Dim msg

Dim customer_email
Dim customer_no

Dim OrderNo

Dim eConF
Dim eConK

Dim w_order_no
Dim w_body_hd
Dim w_body_dt1
Dim w_body_dt2
Dim w_body_tl

Dim w_comp_ryakushou
Dim w_order_estimate
Dim w_todokesaki_no
Dim w_payment_method
Dim w_loan_company
Dim w_product_am
Dim w_holiday_fl
Dim wKabusokuAM

Dim wSalesTaxRate
Dim wPrice
Dim wProdTotalAm

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS_order_header
Dim RS_order_detail
Dim RS_web_order_header
Dim RS_web_order_detail
Dim RS_web_customer
Dim RS_cntl
Dim RS_customer
Dim RS_company
Dim RS_prod
Dim RS_set
Dim RS_calender

Dim wSQL
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

'=======================================================================

'---- UserID ���o��
userID = Session("userID")
userName = Session("userName")

'---- �Z�b�V�����؂�`�F�b�N
if userID = ""then
	Response.Redirect g_HTTP
end if

Session("msg") = ""
w_msg = ""

OrderNo = ReplaceInput(Request("OrderNo"))
eConf = ReplaceInput(Trim(Request("eConf")))
eConK = ReplaceInput(Trim(Request("eConK")))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	Connection.RollbackTrans
	wErrDesc = "OrderSubmit.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- �G���[�������Ƃ��͂��肪�Ƃ��y�[�W�A�G���[������Β������e���̓y�[�W��
if w_msg = "" then
	Session(g_cookie_name) = ""		'�������iCookie���N���A
	Session("OrderAtOnce") = "1"	'2013/07/30 GV #1618 add
	Response.Redirect "Thanks.asp?order_no=" & w_order_no & "&product_am=" & w_product_am & "&order_estimate=" & Server.URLEncode(w_order_estimate) & "&payment_method=" & Server.URLEncode(w_payment_method) & "&loan_company=" & Server.URLEncode(w_loan_company)
else
	Session("msg") = "<font color='#ff0000'>" & w_msg & "</font>"
	Response.Redirect "OrderInfoEnter.asp"
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
Dim i

'---- �g�����U�N�V�����J�n
Connection.BeginTrans

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- ��Џ��̎��o��
call get_company()

'---- �󒍏��̓o�^
call insert_web_order_header()
if w_msg <> "" then
	Connection.RollbackTrans
	exit function			'if error exit
end if

call insert_web_order_detail()
if w_msg <> "" then
	Connection.RollbackTrans
	exit function			'if error exit
end if

'---- ���󒍏��폜
call delete_web_order()

'---- �g�����U�N�V�����I��
Connection.CommitTrans

'---- ���[�����M
call send_order_mail()

End Function

'========================================================================
'
'	Function	Web�󒍂̓o�^
'
'========================================================================
'
Function insert_web_order_header()
Dim i

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'---- ���󒍂̎��o��
wSQL = ""
wSQL = wSQL & "SELECT a.*"
wSQL = wSQL & "  FROM ���� a"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	  
Set RS_order_header = Server.CreateObject("ADODB.Recordset")
RS_order_header.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_order_header.EOF = true then
	w_msg = "NoData"
	exit function
end if

'2011/01/28 GV Mod Start
''---- �ڋq�͐���o�^
'if (Trim(RS_order_header("�͐�Z���A��")) = 0) then
'	w_todokesaki_no = insert_todokesaki()
'else
'	w_todokesaki_no = Trim(RS_order_header("�͐�Z���A��"))
'end if

w_todokesaki_no = Trim(RS_order_header("�͐�Z���A��"))
'2011/01/28 GV Mod End

'---- insert ��
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM Web��"
wSQL = wSQL & " WHERE 1 = 2"
 
Set RS_web_order_header = Server.CreateObject("ADODB.Recordset")
RS_web_order_header.Open wSQL, Connection, adOpenStatic, adLockOptimistic

'---- ��/���ρ@�Z�[�u
if RS_order_header("���σt���O") = "Y" then
	w_order_estimate = "������"
else
	w_order_estimate = "������"
end if

'---- �󒍔ԍ����o��(BlueGate�Ăяo�����͎󒍔ԍ��͍̔Ԃ���Ă���)
if OrderNo = "" then
	w_order_no = CLng(get_cntl_no("����","�ԍ�","Web��"))
else
	if isNumeric(OrderNo) = false then
		w_msg = w_msg & "�󒍔ԍ��G���["
		exit function
	end if
	w_order_no = CLng(OrderNo)
end if

'---- �󒍍쐬
RS_web_order_header.AddNew

RS_web_order_header("�󒍔ԍ�") = w_order_no

For i=0 to RS_web_order_header.Fields.Count - 1
	if RS_order_header(i).Name <> "SessionID" then
		if isnull(RS_order_header(RS_order_header(i).Name)) = false then
			RS_web_order_header(RS_order_header(i).Name) = RS_order_header(RS_order_header(i).Name)
		end if
	end if
Next

'---- ����A���}�g�Ŏ��Ԏw�肠��̏ꍇ�͎��Ԏw����e�^����ЂɑΉ��������Ԃɓǂݑւ�   2011/06/29 an add s
if Trim(RS_order_header("���Ԏw��")) <> "" then
	if  RS_order_header("�^����ЃR�[�h") = "1" OR RS_order_header("�^����ЃR�[�h") = "2" then

		call getCntlMst("��","���Ԏw��ǂݑւ�",RS_order_header("���Ԏw��"), vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)

		'---- ����
		if RS_order_header("�^����ЃR�[�h") = "1" then
			RS_web_order_header("���Ԏw��") = vItemChar1
		end if
		'---- ���}�g
		if RS_order_header("�^����ЃR�[�h") = "2" then
			RS_web_order_header("���Ԏw��") = vItemChar2
		end if
	end if
end if                                                                                '2011/06/29 an add e
 
'---- �͐���̃Z�b�g
RS_web_order_header("�͐�Z���A��") = w_todokesaki_no
CAll SetTodokesaki(w_todokesaki_no)

'---- �����ҏ��̃Z�b�g
CAll SetChuumonsha()

'---- �R���r�j�x��URL���o�^�@'2007/01/11
if RS_web_order_header("�x�����@") = "�R���r�j�x��" then
	RS_web_order_header("eContext�x�����@URL") = eConK
	RS_web_order_header("eContext�U���[URL") = eConF
end if

RS_web_order_header("�󒍌`��") = "�C���^�[�l�b�g"

RS_web_order_header("�ŏI�X�V��") = now()
RS_web_order_header("���͓�") = now()

RS_web_order_header.update

'---- �ڋq�U�����`�l�o�^
if (Trim(RS_web_order_header("�U�����`�l")) <> "") then
	call update_furikomimeiginin()
end if

'---- �x�����@�A�� �Z�[�u
w_payment_method = RS_web_order_header("�x�����@")
w_product_am = RS_web_order_header("���i���v���z")
if w_payment_method = "���[��" then
	w_loan_company = RS_web_order_header("���[�����")
end if

wKabusokuAm = RS_web_order_header("�ߕs�����E���z")

if Trim(RS_web_order_header("���x�[�g�g�p�t���O")) = "Y" then
		call updateKabusokuAm()
end if

'---- ���x�[�g�g�p��0�~�ɂȂ����Ƃ�
if RS_web_order_header("�󒍍��v���z") = 0 then
	RS_web_order_header("�x�����@") = "����"
	RS_web_order_header.update
end if

'---- ���[���w�b�_�A�g���[���̕ҏW
call edit_mail_ht()

RS_web_order_header.close

End function

'========================================================================
'
'	Function	Web�󒍖��ׂ̓o�^
'
'========================================================================
'
Function insert_web_order_detail()
Dim i
Dim vTotalAm

'---- ���󒍖��ׂ̎��o��
wSQL = ""
wSQL = wSQL & "SELECT a.*"
wSQL = wSQL & "     , b.�Z�b�g���i�t���O"
wSQL = wSQL & "     , b.���[�J�[�������敪"
wSQL = wSQL & "     , b.�󏭐���"
wSQL = wSQL & "     , b.�戵���~��"
wSQL = wSQL & "     , b.�p�ԓ�"
wSQL = wSQL & "     , b.���i��"
wSQL = wSQL & "     , c.�����\����"
wSQL = wSQL & "     , c.�����\���ח\���"
wSQL = wSQL & "  FROM ���󒍖��� a WITH (NOLOCK)"
wSQL = wSQL & "     , Web���i b WITH (NOLOCK)"
wSQL = wSQL & "     , Web�F�K�i�ʍ݌� c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND c.���i�R�[�h = a.���i�R�[�h"
wSQL = wSQL & "   AND c.�F = a.�F"
wSQL = wSQL & "   AND c.�K�i = a.�K�i"
wSQL = wSQL & "   AND SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
wSQL = wSQL & " ORDER BY �󒍖��הԍ�"
	  
Set RS_order_detail = Server.CreateObject("ADODB.Recordset")
RS_order_detail.Open wSQL, Connection, adOpenStatic

'---- �󒍖��ׂ��Ȃ���΃G���[	2012/01/23 hn add
if RS_order_detail.EOF = true then
	w_msg = w_msg & "������������܂���B"
	exit function
end if

'---- insert �󒍖���
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM Web�󒍖���"
wSQL = wSQL & " WHERE 1 = 2"
 
Set RS_web_order_detail = Server.CreateObject("ADODB.Recordset")
RS_web_order_detail.Open wSQL, Connection, adOpenStatic, adLockOptimistic

vTotalAm = 0
'---- �󒍖��׍쐬
Do while RS_order_detail.EOF = false

	RS_web_order_detail.AddNew

	RS_web_order_detail("�󒍔ԍ�") = w_order_no

	For i=0 to RS_order_detail.Fields.Count - 1 - 8		'-1 -���󒍖��׈ȊO�̍��ڐ�(��LSQL)
		if RS_order_detail(i).Name <> "SessionID" then
			RS_web_order_detail(RS_order_detail(i).Name) = RS_order_detail(RS_order_detail(i).Name)
		end if
	Next

	RS_web_order_detail.update

	'---- �F�K�i�ʍ݌Ɂ@�����\���ʁ@�X�V	'2007/04/20
	call updateInventory()
	if w_msg <> "" then
		exit function
	end if

	'---- B�i�A�p�ԕi�̊������ǂ����̃`�F�b�N	'2009/06/17
	if RS_order_detail("B�i�t���O") = "Y"  OR isNull(RS_order_detail("�p�ԓ�")) = false then
		call updateKanbaibi()		'���������Z�b�g
	end if

	'---- ������P�����i�Ȃ珤�i�}�X�^�X�V
'	if (RS_order_detail("������P���t���O") = "Y") AND (w_order_estimate = "������") AND ((w_payment_method = "�N���W�b�g�J�[�h") OR (w_payment_method = "�����")) then  '2010/12/20 hn del
	if RS_order_detail("������P���t���O") = "Y" then	'2010/12/20 hn add
		call updateProduct()
	end if

	'---- ���[�����׍s�ҏW
	call edit_mail_dt()

	'---- �󒍍��v���z�v�Z
	vTotalAm = vTotalAm + RS_order_detail("�󒍋��z")

	'----- ���R�����h���i�w�����O�o�^   2009/12/17 add hn 2010/03/04 an �L���� 2012/01/10 an ��~
	'call AddRecommendPurchaseLog(RS_order_detail("���[�J�[�R�[�h"), RS_order_detail("���i�R�[�h"))

	RS_order_detail.MoveNext
Loop

'---- �󒍍��v���z�`�F�b�N
if vTotalAm <> w_product_am then
	w_msg = w_msg & "����i���v���z�s��v� �������e���ēx���m�F�肢�܂��B"
end if

RS_web_order_detail.close
RS_order_detail.close

End function

'2011/01/28 GV Del Start
''========================================================================
''
''	Function	�͐���̓o�^
''
''========================================================================
''
'Function insert_todokesaki()
'Dim i
'Dim v_Max_no
'
''---- ����Z�������邩�ǂ����`�F�b�N
'wSQL = ""
'wSQL = wSQL & "SELECT �Z���A��"
'wSQL = wSQL & "  FROM Web�ڋq�Z�� WITH (NOLOCK)"
'wSQL = wSQL & " WHERE �ڋq�ԍ� = " & userID
'wSQL = wSQL & "   AND �Z������ = '" & Replace(RS_order_header("�͐於�O"),"'","''") & "'"
'wSQL = wSQL & "   AND �ڋq�X�֔ԍ� = '" & RS_order_header("�͐�X�֔ԍ�") & "'"
'wSQL = wSQL & "   AND �ڋq�s���{�� = '" & RS_order_header("�͐�s���{��") & "'"
'wSQL = wSQL & "   AND �ڋq�Z�� = '" & Replace(RS_order_header("�͐�Z��"),"'","''") & "'"
'
'Set RS_customer = Server.CreateObject("ADODB.Recordset")
'RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic
'
'if RS_customer.EOF = false then				'����Z������
'	insert_todokesaki = RS_customer("�Z���A��")
'	Exit Function
'end if
'
''---- MAX�Z���A�Ԃ̎��o��
'wSQL = ""
'wSQL = wSQL & "SELECT MAX(�Z���A��) AS MAX�Z���A��"
'wSQL = wSQL & "  FROM Web�ڋq�Z�� WITH (NOLOCK)"
'wSQL = wSQL & " WHERE �ڋq�ԍ� = " & userID
'	  
'Set RS_customer = Server.CreateObject("ADODB.Recordset")
'RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic
'
'v_max_no = RS_customer("MAX�Z���A��") + 1
'
''---- insert �ڋq�Z��
'wSQL = ""
'wSQL = wSQL & "SELECT *"
'wSQL = wSQL & "  FROM Web�ڋq�Z��"
'wSQL = wSQL & " WHERE 1 = 2"
' 
'Set RS_customer = Server.CreateObject("ADODB.Recordset")
'RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic
'
'RS_customer.AddNew
'
'RS_customer("�ڋq�ԍ�") = UserID
'RS_customer("�Z���A��") = v_Max_no
'RS_customer("�Z���敪") = "�͐�"
'RS_customer("�Z������") = RS_order_header("�͐於�O")
'RS_customer("�ڋq�X�֔ԍ�") = RS_order_header("�͐�X�֔ԍ�")
'RS_customer("�ڋq�s���{��") = RS_order_header("�͐�s���{��")
'RS_customer("�ڋq�Z��") = RS_order_header("�͐�Z��")
'RS_customer("�Ζ���t���O") = "N"
'RS_customer("�[�i�����t�t���O") = RS_order_header("�͐�[�i�����t�t���O")
'RS_customer("�K��͐�t���O") = "N"
'RS_customer("�ŏI�X�V��") = Now()
'RS_customer("�ŏI�X�V�҃R�[�h") = "Internet"
'
'RS_customer.update
'
''---- insert �ڋq�Z���d�b�ԍ�
'wSQL = ""
'wSQL = wSQL & "SELECT *"
'wSQL = wSQL & "  FROM Web�ڋq�Z���d�b�ԍ�"
'wSQL = wSQL & " WHERE 1 = 2"
' 
'Set RS_customer = Server.CreateObject("ADODB.Recordset")
'RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic
'
'RS_customer.AddNew
'
'RS_customer("�ڋq�ԍ�") = UserID
'RS_customer("�Z���A��") = v_Max_no
'RS_customer("�d�b�A��") = 1
'RS_customer("�d�b�敪") = "�d�b"
'RS_customer("�ڋq�d�b�ԍ�") = RS_order_header("�͐�d�b�ԍ�")
'RS_customer("�����p�ڋq�d�b�ԍ�") = cf_numeric_only(RS_order_header("�͐�d�b�ԍ�"))
'RS_customer("�ŏI�X�V��") = Now()
'RS_customer("�ŏI�X�V�҃R�[�h") = "Internet"
'
'RS_customer.update
'RS_customer.close
'
'insert_todokesaki = v_Max_no
'
'End function
'2011/01/28 GV Del End

'========================================================================
'
'	Function	�͐���̃Z�b�g(Web�󒍂�)
'
'========================================================================
'
Function SetTodokesaki(pNo)
Dim RSv

'---- 
wSQL = ""
wSQL = wSQL & "SELECT a.*"
wSQL = wSQL & "     , b.�ڋq�d�b�ԍ�"
wSQL = wSQL & "  FROM Web�ڋq�Z�� a WITH (NOLOCK)"
wSQL = wSQL & "     , Web�ڋq�Z���d�b�ԍ� b WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.�ڋq�ԍ� = a.�ڋq�ԍ�"
wSQL = wSQL & "   AND b.�Z���A�� = a.�Z���A��"
wSQL = wSQL & "   AND b.�d�b�A�� = 1"
wSQL = wSQL & "   AND a.�ڋq�ԍ� = " & userID
wSQL = wSQL & "   AND a.�Z���A�� = " & pNo

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = false then	
	RS_web_order_header("�͐於�O") = RSv("�Z������")
	RS_web_order_header("�͐�X�֔ԍ�") = RSv("�ڋq�X�֔ԍ�")
	RS_web_order_header("�͐�s���{��") = RSv("�ڋq�s���{��")
	RS_web_order_header("�͐�Z��") = RSv("�ڋq�Z��") 
	RS_web_order_header("�͐�d�b�ԍ�") = RSv("�ڋq�d�b�ԍ�")
else
	RS_web_order_header("�͐於�O") = ""
	RS_web_order_header("�͐�X�֔ԍ�") = ""
	RS_web_order_header("�͐�s���{��") = ""
	RS_web_order_header("�͐�Z��") = ""
	RS_web_order_header("�͐�d�b�ԍ�") = ""
end if

RSv.close

End function

'========================================================================
'
'	Function	�����ҏ��̃Z�b�g(Web�󒍂�)
'
'========================================================================
'
Function SetChuumonsha()
Dim RSv

'---- 
wSQL = ""
wSQL = wSQL & "SELECT a.�ڋq��"
wSQL = wSQL & "     , b.�ڋq�X�֔ԍ�"
wSQL = wSQL & "     , b.�ڋq�s���{��"
wSQL = wSQL & "     , b.�ڋq�Z��"
wSQL = wSQL & "     , c.�ڋq�d�b�ԍ�"
wSQL = wSQL & "  FROM Web�ڋq a WITH (NOLOCK)"
wSQL = wSQL & "     , Web�ڋq�Z�� b WITH (NOLOCK)"
wSQL = wSQL & "     , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.�ڋq�ԍ� = a.�ڋq�ԍ�"
wSQL = wSQL & "   AND c.�ڋq�ԍ� = b.�ڋq�ԍ�"
wSQL = wSQL & "   AND c.�Z���A�� = b.�Z���A��"
wSQL = wSQL & "   AND a.�ڋq�ԍ� = " & userID
wSQL = wSQL & "   AND b.�Z���A�� = 1"
wSQL = wSQL & "   AND c.�d�b�A�� = 1"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = false then	
	RS_web_order_header("�����Җ��O") = RSv("�ڋq��")
	RS_web_order_header("�����җX�֔ԍ�") = RSv("�ڋq�X�֔ԍ�")
	RS_web_order_header("�����ғs���{��") = RSv("�ڋq�s���{��")
	RS_web_order_header("�����ҏZ��") = RSv("�ڋq�Z��") 
	RS_web_order_header("�����ғd�b�ԍ�") = RSv("�ڋq�d�b�ԍ�")
else
	RS_web_order_header("�����Җ��O") = ""
	RS_web_order_header("�����җX�֔ԍ�") = ""
	RS_web_order_header("�����ғs���{��") = ""
	RS_web_order_header("�����ҏZ��") = ""
	RS_web_order_header("�����ғd�b�ԍ�") = ""
end if

RSv.close

End function

'========================================================================
'
'	Function	�U�����`�l�̍X�V
'
'========================================================================
'
Function update_furikomimeiginin()
Dim i

'---- �U�����`�l�̍X�V
wSQL = ""
wSQL = wSQL & "SELECT �U�����`�l"
wSQL = wSQL & "       , �ŏI�X�V��"
wSQL = wSQL & "       , �ŏI�X�V�҃R�[�h"
wSQL = wSQL & "    FROM Web�ڋq"
wSQL = wSQL & "   WHERE �ڋq�ԍ� = " & RS_order_header("�ڋq�ԍ�")

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_customer.EOF = false then	
	RS_customer("�U�����`�l") = RS_order_header("�U�����`�l")
	RS_customer("�ŏI�X�V��") = Now()
	RS_customer("�ŏI�X�V�҃R�[�h") = "Internet"
end if

RS_customer.update
RS_customer.close

End function

'========================================================================
'
'	Function	�����ߕs�����z�̍X�V
'
'========================================================================
'
Function updateKabusokuAm()
Dim i
Dim vCustKabusoku

'---- �����ߕs�����z�̍X�V
wSQL = ""
wSQL = wSQL & "SELECT �����ߕs�����z"
wSQL = wSQL & "       , �ŏI�X�V��"
wSQL = wSQL & "       , �ŏI�X�V�҃R�[�h"
wSQL = wSQL & "       , �ŏI�X�V������"
wSQL = wSQL & "    FROM Web�ڋq"
wSQL = wSQL & "   WHERE �ڋq�ԍ� = " & RS_order_header("�ڋq�ԍ�")

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_customer.EOF = false then	
	vCustKabusoku = RS_customer("�����ߕs�����z") - wKabusokuAm
	if vCustKabusoku < 0 then
		w_msg = w_msg & "���x�[�g���z���s�����Ă��܂��B �������e���ēx���m�F�肢�܂��B<br>"
	else
		RS_customer("�����ߕs�����z") = RS_customer("�����ߕs�����z") - wKabusokuAm
		RS_customer("�ŏI�X�V��") = Now()
		RS_customer("�ŏI�X�V�҃R�[�h") = "Internet"
		RS_customer("�ŏI�X�V������") = "OrderSubmit.asp"
	end if
else
	w_msg = w_msg & "�ڋq��񂪂���܂���B<br>"
end if

RS_customer.update
RS_customer.close

End function

'========================================================================
'
'	Function	���󒍂̍폜
'
'========================================================================
'
Function delete_web_order()

'---- ���󒍂̍폜
RS_order_header.delete
RS_order_header.close

End function

'========================================================================
'
'	Function	���R�����h���i�w�����O	2009/12/17
'
'========================================================================
'
Function AddRecommendPurchaseLog(pMakerCd, pProductCd)

Dim RSv

'---- 
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM ���R�����h���i�w�����O"
wSQL = wSQL & " WHERE 1 = 2"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

'---- ���R�����h���i�w�����O�o�^
RSv.AddNew

RSv("���R�����h���[�U�[ID") = gSessionID				'2011/04/14 hn mod
RSv("���[�J�[�R�[�h") = pMakerCd
RSv("���i�R�[�h") = pProductCd
RSv("���[�U�[�G�[�W�F���g") = Request.ServerVariables("HTTP_USER_AGENT")
RSv("�A�N�Z�X��") = Now()

RSv.Update
RSv.close

End function

'========================================================================
'
'	Function	�R���g���[���}�X�^����ԍ��̔�
'
'		parm: sub_sustem_cd, item_cd, item_sub_cd
'		return:	�ԍ�
'
'========================================================================
'
Function get_cntl_no(p_sub_system_cd, p_item_cd, p_item_sub_cd)

'---- �R���g���[���}�X�^���o��
wSQL = ""
wSQL = wSQL & "SELECT item_num1"
wSQL = wSQL & "  FROM �R���g���[���}�X�^"
wSQL = wSQL & " WHERE sub_system_cd = '" & p_sub_system_cd & "'"
wSQL = wSQL & "   AND item_cd = '" & p_item_cd & "'"
wSQL = wSQL & "   AND item_sub_cd = '" & p_item_sub_cd & "'"
	  
'@@@@@@response.write(wSQL)

Set RS_cntl = Server.CreateObject("ADODB.Recordset")
RS_cntl.Open wSQL, Connection, adOpenStatic, adLockOptimistic

RS_cntl("item_num1") = Clng(RS_cntl("item_num1")) + 1
get_cntl_no = RS_cntl("item_num1")

RS_cntl.update
RS_cntl.close

End function

'========================================================================
'
'	Function	�F�K�i�ʍ݌ɂ̈����\���ʂ��X�V	2007/04/20
'
'========================================================================
'
Function updateInventory()

Dim RSv

wSQL = ""
wSQL = wSQL & "SELECT �����\����"
wSQL = wSQL & "     , B�i�����\����"
wSQL = wSQL & "  FROM Web�F�K�i�ʍ݌�"
wSQL = wSQL & " WHERE ���[�J�[�R�[�h = '" & RS_order_detail("���[�J�[�R�[�h")  & "'"
wSQL = wSQL & "   AND ���i�R�[�h = '" & RS_order_detail("���i�R�[�h")  & "'"
wSQL = wSQL & "   AND �F = '" & RS_order_detail("�F")  & "'"
wSQL = wSQL & "   AND �K�i = '" & RS_order_detail("�K�i")  & "'"
	  
'@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_order_detail("B�i�t���O") <> "Y" then
	if isNull(RS_order_detail("�p�ԓ�")) = false AND RSv("�����\����") < RS_order_detail("�󒍐���") then
		w_msg = w_msg & RS_order_detail("���i��") & "�́A�݌ɂ�" & RSv("�����\����") & "��������܂���B�@���ʂ�ύX���Ă��������������B<br>"
		exit function
	else
		RSv("�����\����") = RSv("�����\����") - RS_order_detail("�󒍐���")
	end if

else
	if RSv("B�i�����\����") >= RS_order_detail("�󒍐���") then
		RSv("B�i�����\����") = RSv("B�i�����\����") - RS_order_detail("�󒍐���")
	else
		w_msg = w_msg & RS_order_detail("���i��") & "�́A�݌ɂ�" & RSv("B�i�����\����") & "��������܂���B�@���ʂ�ύX���Ă��������������B<br>"
		exit function
	end if
end if

RSv.update
RSv.close

End function

'========================================================================
'
'	Function	���i�̊��������Z�b�g	2007/04/20
'
'========================================================================
'
Function updateKanbaibi()

Dim RSv

wSQL = ""
wSQL = wSQL & "SELECT SUM(�����\����) AS �����\����"
wSQL = wSQL & "     , SUM(B�i�����\����) AS B�i�����\����"
wSQL = wSQL & "  FROM Web�F�K�i�ʍ݌�"
wSQL = wSQL & " WHERE ���[�J�[�R�[�h = '" & RS_order_detail("���[�J�[�R�[�h")  & "'"
wSQL = wSQL & "   AND ���i�R�[�h = '" & RS_order_detail("���i�R�[�h")  & "'"
	  
'@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if (RS_order_detail("B�i�t���O") = "Y" and RSv("B�i�����\����") <= 0) OR (isNull(RS_order_detail("�p�ԓ�")) = false AND RSv("�����\����") <= 0) then

	RSv.close
	
	wSQL = ""
	wSQL = wSQL & "SELECT ������"
	wSQL = wSQL & "  FROM Web���i"
	wSQL = wSQL & " WHERE ���[�J�[�R�[�h = '" & RS_order_detail("���[�J�[�R�[�h")  & "'"
	wSQL = wSQL & "   AND ���i�R�[�h = '" & RS_order_detail("���i�R�[�h")  & "'"
		  
	'@@@@@@response.write(wSQL)

	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	RSv("������") = Now()

	RSv.update

end if

RSv.close

End function

'========================================================================
'
'	Function	���i�}�X�^�̌�����󒍍ϐ��ʂ��X�V
'
'========================================================================
'
Function updateProduct()

wSQL = ""
wSQL = wSQL & "SELECT ������󒍍ϐ���"
wSQL = wSQL & "     , �����萔��"
wSQL = wSQL & "     , �������i�t���O"
wSQL = wSQL & "  FROM Web���i"
wSQL = wSQL & " WHERE ���[�J�[�R�[�h = '" & RS_order_detail("���[�J�[�R�[�h")  & "'"
wSQL = wSQL & "   AND ���i�R�[�h = '" & RS_order_detail("���i�R�[�h")  & "'"
	  
'@@@@@@response.write(wSQL)

Set RS_prod = Server.CreateObject("ADODB.Recordset")
RS_prod.Open wSQL, Connection, adOpenStatic, adLockOptimistic

RS_prod("������󒍍ϐ���") = RS_prod("������󒍍ϐ���") + RS_order_detail("�󒍐���")
if RS_prod("������󒍍ϐ���") >= RS_prod("�����萔��") then
	RS_prod("�������i�t���O") = ""
end if

RS_prod.update
RS_prod.close

End function

'========================================================================
'
'	Function	���[���w�b�_�A�g���[���̕ҏW
'
'========================================================================
'
Function edit_mail_ht()
Dim i
Dim v_temp
Dim vPaymentMethod

'---- �ڋq�f�[�^���o��
call get_customer()

'---- ���[���w�b�_
w_body_hd = "���@��t�����F" & FormatDateTime(RS_web_order_header("���͓�"), 0) & "�@" & w_order_estimate & "�@" & w_order_no

w_body_hd = w_body_hd & vbNewLine & vbNewLine
w_body_hd = w_body_hd & "�|�|�|�|�|�@���q�l�@�|�|�|�|�|" & vbNewLine
w_body_hd = w_body_hd & "���O�@�@�@�F�@" & RS_customer("�ڋq��") & vbNewLine
w_body_hd = w_body_hd & "�ӂ肪�� �F�@" & RS_customer("�ڋq�t���K�i") & vbNewLine
w_body_hd = w_body_hd & "�Z���@�@�@�F�@��" & RS_customer("�ڋq�X�֔ԍ�") & "�@" & RS_customer("�ڋq�s���{��") & RS_customer("�ڋq�Z��") & vbNewLine
w_body_hd = w_body_hd & "�d�b�ԍ��F�@" & RS_customer("�ڋq�d�b�ԍ�") & vbNewLine
w_body_hd = w_body_hd & "�e�`�w�@�@�@�F�@" & RS_customer("FAX") & vbNewLine
w_body_hd = w_body_hd & "�d���[���@�F�@" & RS_web_order_header("�ڋqE_mail") & vbNewLine
w_body_hd = w_body_hd & "�ڋq�ԍ��F�@" & RS_web_order_header("�ڋq�ԍ�") & vbNewLine & vbNewLine
		
w_body_hd = w_body_hd & "�|�|�|�|�|�@���͂���@�|�|�|�|" & vbNewLine
if RS_web_order_header("�͐�敪") = "D" then

	if Trim(RS_web_order_header("�͐�Z���A��")) <> "0" then
		call get_todokesaki(Trim(RS_web_order_header("�͐�Z���A��")))

		w_body_hd = w_body_hd & "���O�@�@�@�F�@" & RS_customer("�Z������") & vbNewLine
		w_body_hd = w_body_hd & "�Z���@�@�@�F�@��" & RS_customer("�ڋq�X�֔ԍ�") & "�@" & RS_customer("�ڋq�s���{��") & RS_customer("�ڋq�Z��") & vbNewLine
		w_body_hd = w_body_hd & "�d�b�ԍ��F�@" & RS_customer("�ڋq�d�b�ԍ�") & vbNewLine
	else
		w_body_hd = w_body_hd & "���O�@�@�@�F�@" & RS_web_order_header("�͐於�O") & vbNewLine
		w_body_hd = w_body_hd & "�Z���@�@�@�F�@��" & RS_web_order_header("�͐�X�֔ԍ�") & "�@" & RS_web_order_header("�͐�s���{��") & RS_web_order_header("�͐�Z��") & vbNewLine
		w_body_hd = w_body_hd & "�d�b�ԍ��F�@" & RS_web_order_header("�͐�d�b�ԍ�") & vbNewLine
	end if

	if RS_web_order_header("�͐�[�i�����t�t���O") = "Y" then
		w_body_hd = w_body_hd & "�[�i�� �@�F�@���t���ėǂ�" & vbNewLine & vbNewLine
	else
		w_body_hd = w_body_hd & "�[�i�� �@�F�@���t���Ȃ�" & vbNewLine & vbNewLine
	end if
else
	w_body_hd = w_body_hd & "�����@����" & vbNewLine & vbNewLine
end if

w_body_hd = w_body_hd & "�|�|�|�|�|�@�z���w��@�|�|�|�|" & vbNewLine

'2011/06/01 if-web del start
'if RS_web_order_header("�^����ЃR�[�h") = "1" then
'	w_body_hd = w_body_hd & "�^����Ё@�@�@�F�@����}��" & vbNewLine
'end if
'if RS_web_order_header("�^����ЃR�[�h") = "2" then
'	w_body_hd = w_body_hd & "�^����Ё@�@�@�F�@���}�g�^�A" & vbNewLine
'end if
'if RS_web_order_header("�^����ЃR�[�h") = "3" then
'	w_body_hd = w_body_hd & "�^����Ё@�@�@�F�@���R�ʉ^" & vbNewLine
'end if
'2011/06/01 if-web del end

if Trim(RS_web_order_header("�w��[��")) <> "" then
	w_body_hd = w_body_hd & "�z�����w�� �@�F�@" & RS_web_order_header("�w��[��") & vbNewLine
end if

if Trim(RS_web_order_header("���Ԏw��")) <> "" then
	w_body_hd = w_body_hd & "�z�����Ԏw��F�@" & RS_web_order_header("���Ԏw��") & vbNewLine
end if

if RS_web_order_header("�c�Ə��~�߃t���O") = "Y" then
	w_body_hd = w_body_hd & "�c�Ə��~��" & vbNewLine
end if

if RS_web_order_header("�ꊇ�o�׃t���O") = "Y" then
	w_body_hd = w_body_hd & "���i���S�đ����Ă���o�ׂ��s��" & vbNewLine & vbNewLine
end if
if RS_web_order_header("�ꊇ�o�׃t���O") = "N" then
	w_body_hd = w_body_hd & "�݌ɂ̂��鏤�i����o�ׂ��s��" & vbNewLine & vbNewLine
end if

w_body_hd = w_body_hd & "���l�@�F" & vbNewLine & RS_web_order_header("���ϔ��l") & vbNewLine

w_body_hd = w_body_hd & vbNewLine

'---- �g���[��
'---- ���z�֘A 2007/05/09
if w_payment_method = "�N���W�b�g�J�[�h" OR w_payment_method = "�����" then
	w_body_tl = w_body_tl & "���i���v�i�ō��݁j�F�@���i���v���z���m��" & vbNewLine
	wPrice = Fix(RS_web_order_header("����") * (100 + wSalesTaxRate) / 100)
	w_body_tl = w_body_tl & "�����i�ō��݁j�F�@" & FormatCurrency(wPrice,0) & vbNewLine

		if w_payment_method = "�����" then
			wPrice = Fix(RS_web_order_header("����萔��") * (100 + wSalesTaxRate) / 100)
			w_body_tl = w_body_tl & "����萔���i�ō��݁j�F�@" & FormatCurrency(wPrice,0) & vbNewLine
		end if
	
	if RS_web_order_header("���x�[�g�g�p�t���O") = "Y" then
		w_body_tl = w_body_tl & "�N���W�b�g/�ߕs�����F�@" & FormatCurrency(RS_web_order_header("�ߕs�����E���z") * (-1) ,0) & vbNewLine
	end if

	w_body_tl = w_body_tl & "���v���z�i�ō��݁j�F�@" & FormatCurrency(RS_web_order_header("�󒍍��v���z"),0) & vbNewLine & vbNewLine

end if

'---- �x�����@
vPaymentMethod = w_payment_method

if RS_web_order_header("�󒍍��v���z") = 0 then
	w_body_tl = w_body_tl	& "�x�����@�@�F�@���x�����s�v" & vbNewLine 
else
	w_body_tl = w_body_tl	& "�x�����@�@�F�@" & Replace(vPaymentMethod, "�R���r�j�x��", "�l�b�g�o���L���O�E�䂤����E�R���r�j����") & vbNewLine 
end if

if RS_web_order_header("�x�����@") = "��s�U��" then
	if RS_web_order_header("�U�����`�l") <> "" then
		w_body_tl = w_body_tl & "�U�����`�l�F�@" & RS_web_order_header("�U�����`�l") & vbNewLine
	end if
end if

if RS_web_order_header("�x�����@") = "�����" then
	w_body_tl = w_body_tl & vbNewLine
end if

if RS_web_order_header("�x�����@") = "���[��" then
	if RS_web_order_header("���[����������t���O") = "Y" then
		w_body_tl = w_body_tl & "��������@�@�@�F" & FormatCurrency(RS_web_order_header("���[������"),0) & vbNewLine
	else
		w_body_tl = w_body_tl & "�����Ȃ�" & vbNewLine
	end if

	if Trim(RS_web_order_header("�I�����C�����[���\���t���O")) <> "Y" then
		if Trim(RS_web_order_header("��]���[����")) <> "" then
			w_body_tl = w_body_tl & "��]���[���񐔁F�@" & RS_web_order_header("��]���[����") & vbNewLine
		end if
		if Trim(RS_web_order_header("���[�����z")) <> "" then
			v_temp = RS_web_order_header("���[�����z")
			w_body_tl = w_body_tl & "���[�����z�@�F�@" & FormatCurrency(Ccur(v_temp)) & vbNewLine
		end if
	else
		w_body_tl = w_body_tl & "�I�����C�����[���\��" & vbNewLine
		if RS_web_order_header("���[�����") = "�Z���g����" then
			w_body_tl = w_body_tl & "�i�Z�f�B�i���p�j" & vbNewLine
		end if
		if RS_web_order_header("���[�����") = "�I���R" then
			w_body_tl = w_body_tl & "�i�I���R���p�j" & vbNewLine
		end if
	end if
	w_body_tl = w_body_tl & vbNewLine
end if

'---- ���x�[�g���b�Z�[�W
if RS_web_order_header("���x�[�g�g�p�t���O") = "Y" then
	w_body_tl = w_body_tl & vbNewLine & "�N���W�b�g/�ߕs�����́A���̂������E���ς�݂̂ɏ[������܂��B" & vbNewLine & "�L�����Z�����Ă����p�ɂȂ�Ȃ��ꍇ�͕��Љc�ƈ��܂ł��A�����������B" & vbNewLine & vbNewLine
end if

'---- �̎���
if RS_web_order_header("�̎������s�t���O") = "Y" then
	w_body_tl = w_body_tl & "�̎����K�v"
	if RS_web_order_header("�̎�������") <> "" then
		'2012/09/25 nt mod
		w_body_tl = w_body_tl & "�@�@�̎��������F" & RS_web_order_header("�̎�������") & " �l"
		w_body_tl = w_body_tl & "�@�@�A�������F" & RS_web_order_header("�̎����A������")
		'w_body_tl = w_body_tl & "�@�@�̎��؈���F" & RS_web_order_header("�̎�������") & " �l"
		'w_body_tl = w_body_tl & "�@�@�̎��ؒA�������F" & RS_web_order_header("�̎����A������")
	end if
end if

customer_email = RS_web_order_header("�ڋqE_mail")	'�ڋq���[���A�h���X�Z�[�u
customer_no = RS_web_order_header("�ڋq�ԍ�")	'�ڋq�ԍ��Z�[�u

RS_customer.close

End function

'========================================================================
'
'	Function	���[�����׍s�ҏW
'
'========================================================================
'
Function edit_mail_dt()

Dim v_body_dt
Dim v_inv1
Dim v_inv2
Dim v_product_nm
Dim vInventoryCd
Dim vProdTermFl

v_product_nm = RS_web_order_detail("���i��")
if Trim(RS_web_order_detail("�F")) <> "" then
	v_product_nm = v_product_nm & "/" & RS_web_order_detail("�F")
end if
if Trim(RS_web_order_detail("�K�i")) <> "" then
	v_product_nm = v_product_nm & "/" & RS_web_order_detail("�K�i")
end if

if RS_web_order_detail("���i��") <> RS_web_order_detail("���i�R�[�h") then
		v_product_nm = v_product_nm & " (" & RS_web_order_detail("���i�R�[�h") & ")"
end if

if RS_web_order_detail("B�i�t���O") = "Y" then
		v_product_nm = v_product_nm & " (B�i�j"
end if

v_body_dt = ""
v_body_dt = v_body_dt & "���[�J�[	�F�@" & RS_web_order_detail("���[�J�[��") & vbNewLine
v_body_dt = v_body_dt & "���i�� �@�F�@" & v_product_nm & vbNewLine

wPrice = calcPrice(RS_web_order_detail("�󒍒P��"), wSalesTaxRate)
wProdTotalAm = wProdTotalAm + (wPrice * RS_web_order_detail("�󒍐���"))

w_html = w_html & "    <td align='right' width='100'>"

'---- �P���A���ʁA���z 2007/05/09
if w_payment_method = "�N���W�b�g�J�[�h" OR w_payment_method = "�����" then
	v_body_dt = v_body_dt & "�P��(�ō�)�F�@" & FormatCurrency(wPrice,0) & vbNewLine
	v_body_dt = v_body_dt & "���� �@�@�F�@" & RS_web_order_detail("�󒍐���") & vbNewLine
	v_body_dt = v_body_dt & "���z(�ō�)�F�@" & FormatCurrency(wPrice * RS_web_order_detail("�󒍐���"),0) & vbNewLine
else
	v_body_dt = v_body_dt & "���� �@�@�F�@" & RS_web_order_detail("�󒍐���") & vbNewLine
end if

'---- �p�ԃ`�F�b�N
if  (isNull(RS_order_detail("�戵���~��")) = true AND isNull(RS_order_detail("�p�ԓ�")) = true) _
 OR (isNull(RS_order_detail("�p�ԓ�")) = false AND RS_order_detail("�����\����") > 0) then
	vProdTermFl = "N"
else
	vProdTermFl = "Y"
end if

'---- �݌ɏ��Z�b�g�i�ڋq�p�j
vInventoryCd = GetInventoryStatus(RS_web_order_detail("���[�J�[�R�[�h"),RS_web_order_detail("���i�R�[�h"),RS_web_order_detail("�F"),RS_web_order_detail("�K�i"),RS_order_detail("�����\����"),RS_order_detail("�󏭐���"),RS_order_detail("�Z�b�g���i�t���O"),RS_order_detail("���[�J�[�������敪"),RS_order_detail("�����\���ח\���"),vProdTermFl)
'---- �݌ɂ�������΁A�u�݌ɁF�����v�ƒǋL�B�݌ɂ���ꍇ�͋L�ڂ��Ă��Ȃ��B
if vInventoryCd <> "�݌ɂ���" AND vInventoryCd <> "�݌ɋ͏�" then
	v_body_dt = v_body_dt & "�݌� �@�@�F�@����" & vbNewLine
end if
'---- �[���\���\�L
v_body_dt = v_body_dt & "�[���\�� �F�@"
	
'---- w_body_dt1�i�Г��j, w_body_dt2�i�ڋq�j
if vInventoryCd <> "�⍇��" AND vInventoryCd <> "���ʒ���" AND vInventoryCd <> "�����" AND vInventoryCd <> "�戵���~" then
	'---- "��X����"�̂悤�ɂ����悻�̔[�����L�ڂ����ꍇ�͕������C���B
	if vInventoryCd <> "�݌ɂ���" AND vInventoryCd <> "�݌ɋ͏�" AND Right(Trim(vInventoryCd), 2) <> "�\��" then  '2010/05/07 an changed
		w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & "��������܂��B" & vbNewLine & vbNewLine       '2011/09/09 an mod s
		w_body_dt2 = w_body_dt2 & v_body_dt & vInventoryCd & "��������܂��B" & vbNewLine & vbNewLine
	'---- "�݌ɂ���"��"�݌ɋ͏�"��"X/X���\��"
	else
		'---- "�݌ɂ���"�ł��󒍐����݌ɐ������Ȃ��Ƃ��́u�ꕔ�݌ɂ�����܂���v�ƕ\��
		'if vInventoryCd = "�݌ɂ���" AND ( RS_web_order_detail("�󒍐���") > RS_order_detail("�����\����")) then	'2012/08/15 nt mod
		'---- �Z�b�g�i�̏ꍇ�A�����\���ʂ����m�łȂ����߁A���b�Z�[�W���ʂ͏��O�ivInventoryCd�𐳂Ƃ���j
		if vInventoryCd = "�݌ɂ���" AND ( RS_web_order_detail("�󒍐���") > RS_order_detail("�����\����") AND (RS_order_detail("�Z�b�g���i�t���O") <> "Y")) then
			w_body_dt1 = w_body_dt1 & v_body_dt & "�ꕔ�݌ɂ�����܂���"
			w_body_dt2 = w_body_dt2 & v_body_dt & "�ꕔ�݌ɂ�����܂���" & vbNewLine & vbNewLine
		else
			'---- ��L�ȊO�͏]���ʂ�\�L
			w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd
			w_body_dt2 = w_body_dt2 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine
		end if
		
		'---- �Г��͈����\���ʂ��L��
		if RS_order_detail("�Z�b�g���i�t���O") <> "Y" then
			w_body_dt1 = w_body_dt1 & "(" & RS_order_detail("�����\����") & "��)" & vbNewLine & vbNewLine
		else
			w_body_dt1 = w_body_dt1 & vbNewLine & vbNewLine
		end if  
	end if
else
	'---- "�⍇��","���ʒ���","�����","�戵���~"�Ȃ�]���ʂ�\�L
	w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine
	w_body_dt2 = w_body_dt2 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine                             '2011/09/09 an mod e
end if

'if vInventoryCd <> "�⍇��" AND vInventoryCd <> "���ʒ���" AND vInventoryCd <> "�����" AND vInventoryCd <> "�戵���~" then  '2011/09/09 an del s
	''---- "��X����"�̂悤�ɂ����悻�̔[�����L�ڂ����ꍇ�͕������C���B
	'if vInventoryCd <> "�݌ɂ���" AND vInventoryCd <> "�݌ɋ͏�" AND Right(Trim(vInventoryCd), 2) <> "�\��" then   '2010/05/07 an changed
		'w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & "��������܂��B" & vbNewLine & vbNewLine
	'else
		''---- "�݌ɂ���"��"�݌Ɋ�"��"X/X���\��"�Ȃ�]���ʂ�\�L
		'if RS_order_detail("�Z�b�g���i�t���O") <> "Y" then
			'w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & "(" & RS_order_detail("�����\����") & "��)" & vbNewLine & vbNewLine
		'else
			'w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine
		'end if
	'end if	
'else
	''---- "�⍇��","���ʒ���","�����","�戵���~"�Ȃ�]���ʂ�\�L
	'w_body_dt1 = w_body_dt1 & v_body_dt & vInventoryCd & vbNewLine & vbNewLine
'end if            '2011/09/09 an del e

End function

'========================================================================
'
'	Function	���[�����M
'
'========================================================================
'
Function send_order_mail()

Dim v_from_mail
Dim v_BCC_mail    '2010/08/11 an add
Dim OBJ_NewMail
Dim v_subject

'Set OBJ_NewMail = CreateObject("CDONTS.NewMail")
Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

'---- �R���g���[���}�X�^���烁�[���A�h���X���o��
if w_order_estimate = "������" then
	call getCntlMst("����","���M��Email","Web���ϒʒm", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
else
	if w_payment_method = "�N���W�b�g�J�[�h" then
		call getCntlMst("����","���M��Email","Web�J�[�h�󒍒ʒm", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
	else
		call getCntlMst("����","���M��Email","Web�󒍒ʒm", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
	end if
end if

v_from_mail = wItemChar1

'---- �ڋq�ւ̎����������[����BCC���郁�[���A�h���X��肾��
call getCntlMst("����","���M��Email","ShopBCC", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)  '2010/08/11 an add
v_BCC_mail = wItemChar1  '2010/08/11 an add

w_body_dt1 = "�|�|�|�|�|�@���i���ׁ@�|�|�|�|" & vbNewLine & w_body_dt1
w_body_dt2 = "�|�|�|�|�|�@���i���ׁ@�|�|�|�|" & vbNewLine & w_body_dt2

'---- �g���[���ҏW���ɖ��m�肾�������i���v���z��u������
w_body_tl = Replace(w_body_tl, "���i���v���z���m��", FormatCurrency(wProdTotalAm,0))

'---- ���[�����M�@�V���b�v
OBJ_NewMail.from = v_from_mail
OBJ_NewMail.to = v_from_mail
OBJ_NewMail.subject = w_order_estimate & w_order_no & "/" & w_payment_method & " ["  & customer_no & "/Web-Emax/Web��-" & w_payment_method & "]"
OBJ_NewMail.TextBody = w_body_hd & w_body_dt1 & w_body_tl
OBJ_NewMail.MimeFormatted = False

OBJ_NewMail.Send

Set OBJ_NewMail = Nothing

'---- ���[�����M�@�ڋq
'---- �R���g���[���}�X�^����w�b�_�A�g���[���t������o��
call CheckHoliday()		'	�x���p�ǉ��w�b�_���K�v���ǂ����m�F

call getCntlMst("Web","Email","�w�b�_", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

if w_Holiday_fl = "Y" then
	w_body_hd = wItemChar1 & vbNewLine & wItemChar2 & vbNewLine & vbNewLine & w_body_hd
else
	w_body_hd = wItemChar1 & vbNewLine & vbNewLine & w_body_hd
end if

call getCntlMst("Web","Email","�g���[��", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
w_body_tl = w_body_tl & vbNewLine & vbNewLine & wItemChar1

if w_order_estimate = "������" then
	v_subject = "�T�E���h�n�E�X�@�����ώ�t�m�F���[���i�����z�M�j" & w_order_no
else
	v_subject = "�T�E���h�n�E�X�@��������t�m�F���[���i�����z�M�j" & w_order_no
end if

Set OBJ_NewMail = Server.CreateObject("CDO.Message") 

OBJ_NewMail.from = v_from_mail
OBJ_NewMail.to = customer_email
OBJ_NewMail.bcc = v_BCC_mail   '2010/08/11 an add
OBJ_NewMail.subject = v_subject
OBJ_NewMail.TextBody = w_body_hd & w_body_dt2 & w_body_tl
OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"

OBJ_NewMail.Send

Set OBJ_NewMail = Nothing

End function

'========================================================================
'
'	Function	�ڋq���̎��o��
'
'========================================================================
'
Function get_customer()

'---- �ڋq�����o��
wSQL = ""
wSQL = wSQL & "SELECT a.�ڋq��"
wSQL = wSQL & "       , a.�ڋq�t���K�i"
wSQL = wSQL & "       , a.�ڋqE_mail1"
wSQL = wSQL & "       , b.�ڋq�X�֔ԍ�"
wSQL = wSQL & "       , b.�ڋq�s���{��"
wSQL = wSQL & "       , b.�ڋq�Z��"
wSQL = wSQL & "       , c.�ڋq�d�b�ԍ�"
wSQL = wSQL & "       , d.�ڋq�d�b�ԍ� AS FAX"
wSQL = wSQL & "  FROM Web�ڋq a WITH (NOLOCK)"
wSQL = wSQL & "     , Web�ڋq�Z�� b WITH (NOLOCK) LEFT JOIN Web�ڋq�Z���d�b�ԍ� d WITH (NOLOCK)"
wSQL = wSQL & "                                          ON d.�ڋq�ԍ� = b.�ڋq�ԍ�"
wSQL = wSQL & "                                         AND d.�Z���A�� = b.�Z���A��"
wSQL = wSQL & "                                         AND d.�d�b�敪 = 'FAX'" 
wSQL = wSQL & "     , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.�ڋq�ԍ� = " & userID
wSQL = wSQL & "   AND b.�ڋq�ԍ� = a.�ڋq�ԍ�"
wSQL = wSQL & "   AND b.�Z���A�� = 1"
wSQL = wSQL & "   AND c.�ڋq�ԍ� = a.�ڋq�ԍ�"
wSQL = wSQL & "   AND c.�Z���A�� = 1"
wSQL = wSQL & "   AND c.�d�b�A�� = 1"
	  
Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic

End function

'========================================================================
'
'	Function	�ڋq�͐���̎��o��
'
'========================================================================
'
Function get_todokesaki(p_ship_address_no)

'---- �ڋq�͐�����o��
wSQL = ""
wSQL = wSQL & "SELECT b.�Z���A��"
wSQL = wSQL & "       , b.�Z������"
wSQL = wSQL & "       , b.�ڋq�X�֔ԍ�"
wSQL = wSQL & "       , b.�ڋq�s���{��"
wSQL = wSQL & "       , b.�ڋq�Z��"
wSQL = wSQL & "       , c.�ڋq�d�b�ԍ�"
wSQL = wSQL & "  FROM Web�ڋq�Z�� b WITH (NOLOCK)"
wSQL = wSQL & "     , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
wSQL = wSQL & " WHERE b.�ڋq�ԍ� = " & userID
wSQL = wSQL & "   AND b.�Z���A�� = " & Clng(p_ship_address_no)
wSQL = wSQL & "   AND c.�ڋq�ԍ� = b.�ڋq�ԍ�"
wSQL = wSQL & "   AND c.�Z���A�� = b.�Z���A��"
wSQL = wSQL & "   AND c.�d�b�敪 = '�d�b'"
	  
'@@@@@@response.write(wSQL)

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open wSQL, Connection, adOpenStatic, adLockOptimistic

End function

'========================================================================
'
'	Function	�{�x�X�����o��
'
'		return:	w_comp_ryakushou
'
'========================================================================
'
Function get_company()

'---- �{�x�X���o��
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM �{�x�X WITH (NOLOCK)"
wSQL = wSQL & " WHERE �{�x�X�R�[�h = '1'"
	  
Set RS_company = Server.CreateObject("ADODB.Recordset")
RS_company.Open wSQL, Connection, adOpenStatic, adLockOptimistic

w_comp_ryakushou = RS_company("�{�x�X����")

RS_company.close

End function

'========================================================================
'
'	Function	�x���p�w�b�_���K�v���ǂ����`�F�b�N
'
'		return:	w_holiday_fl
'
'========================================================================
'
Function CheckHoliday()

Dim v_time

w_holiday_fl = ""

'---- �J�����_�[�����o��(����)
wSQL = ""
wSQL = wSQL & "SELECT �x���t���O"
wSQL = wSQL & "  FROM �J�����_�[ WITH (NOLOCK)"
wSQL = wSQL & " WHERE �N���� = '" & cf_FormatDate(DateAdd("d", 1, Date()), "YYYY/MM/DD") & "'"
	  
Set RS_calender = Server.CreateObject("ADODB.Recordset")
RS_calender.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS_calender.EOF = false OR DatePart("w", DateAdd("d", 1, Date())) = vbSunday then		'�������x��
	'---- �J�����_�[�����o��(����)
	wSQL = ""
	wSQL = wSQL  & "SELECT �x���t���O"
	wSQL = wSQL  & "  FROM �J�����_�[ WITH (NOLOCK)"
	wSQL = wSQL  & " WHERE �N���� = '" & cf_FormatDate(Date(), "YYYY/MM/DD") & "'"
		  
	Set RS_calender = Server.CreateObject("ADODB.Recordset")
	RS_calender.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	if RS_calender.EOF = false OR DatePart("w", Date()) = vbSunday then		'�������x��
			w_holiday_fl = "Y"
	else
		if DatePart("w", Date()) = vbSaturday then	'�y�j��
			if cf_FormatTime(Now(), "HH:MM") > "17:00" then
				w_holiday_fl = "Y"
			end if
		else				'����
			if cf_FormatTime(Now(), "HH:MM") > "19:00" then
				w_holiday_fl = "Y"
			end if
		end if
	end if
end if

RS_calender.close

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing    '2011/08/01 an add

End function

%>
