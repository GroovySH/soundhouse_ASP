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
'	���������e�̊m�F�y�[�W
'
'2012/06/15 ok �f�U�C���ύX�̂��ߋ��ł����ɐV�K�쐬
'2013/10/21 GV # ��^���i�̕\��
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

'---- Session���
Dim wUserID
Dim wUserName
Dim wMsg
Dim Skey

'---- �󂯓n����������ϐ�

'---- �ڋq���
Dim wCustomerNm
Dim wCustomerZip
Dim wCustomerPref
Dim wCustomerAddress
Dim wCustomerTel
Dim wCustomerKabusokuAm

'---- ���󒍏��
Dim wShipNm
Dim wShipZip
Dim wShipPrefecture
Dim wShipAddress
Dim wShipTel
Dim wFreightForwarder
Dim wDeliveryDt
Dim wDeliveryTm
Dim wEigyoushoDome
Dim wIkkatsu
Dim wPaymentMethod
Dim wFurikomiNm
Dim wRitouFl
Dim wRebateFl


'2013/10/21 GV # add start
'---- ��^���i
Dim wLargeItemHtml
Dim wLargeItemFl
Dim wNonLargeItemFl
wLargeItemHtml = ""
wLargeItemFl = "N"
wNonLargeItemFl = "N"
'2013/10/21 GV # add end

'---- ���z
Dim wPrdctAmTotal
Dim wPrdctAmTotalNoTax
Dim wShippingNoTax
Dim wCodAm
Dim wTax
Dim wOrderAmTotal
Dim wSokoCnt
Dim wSalesTaxRate
Dim wKoguchi
Dim wErrDesc   '2011/08/01 an add
Dim wTotal_NoDaibikiFee  '2012/03/03 an add

'---- DB
Dim Connection

'---- HTML
Dim wProductHtml
Dim wHaisouHtml
Dim wPaymentHtml
Dim wReceiptHtml

'=======================================================================
'	�󂯓n�������o��
'=======================================================================
'---- Session�ϐ�
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session.Contents("msg")

'---- �󂯓n�������o��
Session("msg") = ""

'---- �Z�b�V�����؂�`�F�b�N
If wUserID = "" Then
	Response.Redirect g_HTTP
End If

'=======================================================================
'	Execute main
'=======================================================================
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderConfirm.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

If wMsg <> "" Then
	Session("msg") = wMsg
	Server.Transfer "OrderInfoEnter.asp"
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
'	Function	main proc
'
'========================================================================
Function main()

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'---- �Z�L�����e�B�[�L�[�Z�b�g
Skey = SetSecureKey()

'---- ����ŗ���o��
Call getCntlMst("����","����ŗ�","1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wSalesTaxRate = Clng(vItemNum1)

Call getCustomer()				'�ڋq���̎��o��
Call getOrder()					'���󒍏��̎��o���A�X�V

End Function

'========================================================================
'
'	Function	�ڋq���̎��o��
'
'========================================================================
Function getCustomer()

Dim vSQL
Dim RSv

'---- �ڋq�����o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    a.�ڋq��"
vSQL = vSQL & "  , a.�����ߕs�����z"
vSQL = vSQL & "  , b.�ڋq�X�֔ԍ�"
vSQL = vSQL & "  , b.�ڋq�s���{��"
vSQL = vSQL & "  , b.�ڋq�Z��"
vSQL = vSQL & "  , c.�ڋq�d�b�ԍ�"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web�ڋq a WITH (NOLOCK)"
vSQL = vSQL & "  , Web�ڋq�Z�� b WITH (NOLOCK)"
vSQL = vSQL & "  , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        a.�ڋq�ԍ� = " & wUserID
vSQL = vSQL & "    AND b.�ڋq�ԍ� = a.�ڋq�ԍ�"
vSQL = vSQL & "    AND b.�Z���A�� = 1"
vSQL = vSQL & "    AND c.�ڋq�ԍ� = a.�ڋq�ԍ�"
vSQL = vSQL & "    AND c.�Z���A�� = 1"
vSQL = vSQL & "    AND c.�d�b�A�� = 1"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = True Then
	wMsg = wMsg & "�ڋq��񂪂���܂���B<BR />"
Else
	wCustomerNm = RSv("�ڋq��")
	wCustomerZip = RSv("�ڋq�X�֔ԍ�")
	wCustomerPref = RSv("�ڋq�s���{��")
	wCustomerAddress = RSv("�ڋq�Z��")
	wCustomerTel = RSv("�ڋq�d�b�ԍ�")
	wCustomerKabusokuAm = RSv("�����ߕs�����z")
End If

RSv.Close

End Function

'========================================================================
'
'	Function	�󒍏��̎��o�� �X�V�i����)
'
'========================================================================
Function getOrder()

Dim vSQL
Dim RSv
Dim vLoanDownPayment
Dim vPaymentMethodDisp	'2011/03/22
Dim wLargeItemHTMLBuff	'2013/10/21 GV # add

'---- ���󒍏����o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ����"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'" 

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = True Then
	Exit Function
End If

'---- �͐�
wShipNm = RSv("�͐於�O")
wShipZip = RSv("�͐�X�֔ԍ�")
wShipPrefecture = RSv("�͐�s���{��")
wShipAddress = RSv("�͐�Z��")
wShipTel = RSv("�͐�d�b�ԍ�")

'---- �z�����
Select Case Trim(RSv("�^����ЃR�[�h"))
	Case "1"
		wFreightForwarder = "����}��"
	Case "2"
		wFreightForwarder = "���}�g�^�A"
	Case "3"
		wFreightForwarder = "���R�ʉ^"
	Case "4"
		wFreightForwarder = "���Дz��"
	Case "5"                                            '2011/06/29 an add
		wFreightForwarder = "���Z�^�A"
	Case Else
		wMsg = wMsg & "�z����񂪂���܂���B<BR />"	'2011/04/11 hn add
End Select

wDeliveryDt= RSv("�w��[��")
wDeliveryTm = RSv("���Ԏw��")

If Trim(RSv("�c�Ə��~�߃t���O")) = "Y" Then
	wEigyoushoDome = "�^����Љc�Ə��~��"
Else
	wEigyoushoDome = ""
End If

Select Case Trim(RSv("�ꊇ�o�׃t���O"))
	Case "Y"
		wIkkatsu = "���i���S�đ����Ă���ꊇ�o�ׂ������܂��"
	Case "N"
		wIkkatsu = "�݌ɂ̂��鏤�i����o�ׂ������܂��"
	Case Else
		wIkkatsu = ""
End Select

'�����`�F�b�N
Call check_ritou(wShipZip)
If wRitouFl = "Y" Then
	wIkkatsu = "���͐悪���u�n�̂��߁A���i���S�đ����Ă���ꊇ�o�ׂ������܂��"
End If

'---- �x�����
wPaymentMethod = RSv("�x�����@")
wFurikomiNm = RSv("�U�����`�l")
wRebateFl = RSv("���x�[�g�g�p�t���O")
vLoanDownPayment = ""

vPaymentMethodDisp = wPaymentMethod
if vPaymentMethodDisp = "�R���r�j�x��" then
	vPaymentMethodDisp = "�l�b�g�o���L���O�E�䂤����E�R���r�j����"
end if

if wPaymentMethod = "��s�U��" OR wPaymentMethod = "�����" OR wPaymentMethod = "���[��" OR wPaymentMethod = "�R���r�j�x��" then
else
	wMsg = wMsg & "�x�����@������܂���B<BR />"
end if

'---- �󒍖��׏��\�� + �����v�Z
Call display_order_detail() '�󒍖��׏��\��+�����v�Z

'---- HTML�o��
'------ �z���w��
wHaisouHtml = ""
wHaisouHtml = wHaisouHtml & "    <tr>" & vbNewLine
wHaisouHtml = wHaisouHtml & "      <td colspan='3' class='preview'>" & vbNewLine
wHaisouHtml = wHaisouHtml & "        <dl class='delivery'>" & vbNewLine
wHaisouHtml = wHaisouHtml & "          <dt>�z���w��</dt>" & vbNewLine
wHaisouHtml = wHaisouHtml & "          <dd>��" & wShipZip & "<br>" & wShipPrefecture & wShipAddress & "<br>Tel. " & wShipTel & "<br>" & wShipNm & " �l</dd>" & vbNewLine
wHaisouHtml = wHaisouHtml & "        </dl>" & vbNewLine
wHaisouHtml = wHaisouHtml & "      </td>" & vbNewLine
wHaisouHtml = wHaisouHtml & "      <td colspan='3' class='preview'>" & vbNewLine

If IsNull(wDeliveryDt) = False Then
	wHaisouHtml = wHaisouHtml & "        <dd>�z�B���w��@" & wDeliveryDt & "</dd>" & vbNewLine
End If
If wDeliveryTm <> "" Then
	wHaisouHtml = wHaisouHtml & "        <dd>���ԑюw��@" &  wDeliveryTm & "</dd>" & vbNewLine
End If
If wEigyoushoDome <> "" Then
	wHaisouHtml = wHaisouHtml & "        <dd>" & wEigyoushoDome & "</dd>" & vbNewLine
End If
If wIkkatsu <> "" Then
	wHaisouHtml = wHaisouHtml & "        <dd>" & wIkkatsu & "</dd>" & vbNewLine
End If
wHaisouHtml = wHaisouHtml & "      </td>" & vbNewLine
wHaisouHtml = wHaisouHtml & "    </tr>" & vbNewLine

'------ �x�����@
If wPaymentMethod = "���[��" Then
	If RSv("���[����������t���O") = "Y" Then
		vLoanDownPayment = "(��������)"
	Else
		vLoanDownPayment = "(�����Ȃ�)"
	End If
End If

wPaymentHtml = ""
wPaymentHtml = wPaymentHtml & "    <tr>" & vbNewLine
wPaymentHtml = wPaymentHtml & "      <td colspan='6' class='preview'>" & vbNewLine
wPaymentHtml = wPaymentHtml & "        <dl class='delivery'>" & vbNewLine
wPaymentHtml = wPaymentHtml & "          <dt>���x�������@</dt>" & vbNewLine

If wRebateFl = "Y" And wOrderAmTotal = 0 Then
	wPaymentHtml = wPaymentHtml & "          <dd>���x�����s�v</dd>" & vbNewLine
Else
	wPaymentHtml = wPaymentHtml & "          <dd>" & vPaymentMethodDisp & vLoanDownPayment & "</dd>" & vbNewLine
End If

Select Case wPaymentMethod
	Case "��s�U��"
		wPaymentHtml = wPaymentHtml & "          <dd>�U�����`�@" & wFurikomiNm & "</dd>" & vbNewLine
	Case "���[��"
		If RSv("���[����������t���O") = "Y" Then
			wPaymentHtml = wPaymentHtml & "          <dd>���[�������@�@�@" & FormatNumber(Ccur(RSv("���[������")), 0) & "</dd>" & vbNewLine
		End If
		If RSv("��]���[����") <> 0 Then
			wPaymentHtml = wPaymentHtml & "          <dd>��]���[���񐔁@" & RSv("��]���[����") & "</dd>" & vbNewLine
		End If
		If RSv("���[�����z") <> "0" Then
			wPaymentHtml = wPaymentHtml & "          <dd>���z�x�����z�@�@" & FormatNumber(Ccur(RSv("���[�����z")), 0) & "</dd>" & vbNewLine
		End If
		If RSv("�I�����C�����[���\���t���O") = "Y" Then
			wPaymentHtml = wPaymentHtml & "          <dd>�I�����C���Ń��[���̐\�����݂��s���B�i" & RSv("���[�����") & "�j</dd>" & vbNewLine
		End If
End Select

wPaymentHtml = wPaymentHtml & "        </dl>" & vbNewLine
wPaymentHtml = wPaymentHtml & "      </td>" & vbNewLine
wPaymentHtml = wPaymentHtml & "    </tr>" & vbNewLine

'------ �̎���
wReceiptHtml = ""
If RSv("�̎������s�t���O") = "Y" Then
	wReceiptHtml = wReceiptHtml & "    <tr>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "      <td colspan='6' class='preview'>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "        <dl class='delivery'>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "          <dt>�̎���</dt>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "          <dd>�̎�������F" & RSv("�̎�������") & " �l�@�@�̎����A�������F" & RSv("�̎����A������") & "</dd>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "        </dl>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "      </td>" & vbNewLine
	wReceiptHtml = wReceiptHtml & "    </tr>" & vbNewLine
End If

'2013/10/21 GV # add start
'------ ��^���i
If wLargeItemFl = "Y" Then
	wLargeItemHtml = "    <tr>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "      <td colspan='6' class='preview'>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "        <dl class='delivery'>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "          <dt>��^���i�ɂ���</dt>" & vbNewLine

	'��^���i�ł͂Ȃ����i�ƍ��݂��Ă���ꍇ
	If wNonLargeItemFl = "Y" Then
		Call getCntlMst("��^�ݕ�","�x�����b�Z�[�W","2", wLargeItemHTMLBuff, null, null, null, null, null)
	Else
		'�P�i
		Call getCntlMst("��^�ݕ�","�x�����b�Z�[�W","1", wLargeItemHTMLBuff, null, null, null, null, null)
	End If

	wLargeItemHtml = wLargeItemHtml & "<dd style='color:red;'>" & wLargeItemHTMLBuff & "</dd>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "        </dl>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "      </td>" & vbNewLine
	wLargeItemHtml = wLargeItemHtml & "    </tr>" & vbNewLine
End If
'2013/10/21 GV # add end

'---- ���󒍏��X�V
RSv("�o�בq�ɐ�") = wSokoCnt
RSv("���i���v���z") = wPrdctAmTotalNoTax
RSv("����") = wShippingNoTax
RSv("����萔��") = wCodAm
RSv("�R���r�j�x���萔��") = 0
RSv("�O�ō��v���z") = wTax
RSv("�󒍍��v���z") = wOrderAmTotal
RSv("�����t���O") = wRitouFl

'---- ���x�[�g���z
If wRebateFl = "Y" Then
	RSv("�ߕs�����E���z") = wCustomerKabusokuAm
Else
	RSv("�ߕs�����E���z") = 0
End If

RSv("����ŗ�") = wSalesTaxRate
RSv("�ŏI�X�V��") = Now()

RSv.Update
RSv.Close

'---- �������20���~�ȏ�̒����̓G���[
If wPaymentMethod = "�����" And wOrderAmTotal > 200000 Then
	wMsg = wMsg & "������̏ꍇ�A1��̂�������20���~�����x�ƂȂ�܂��B���������e���͂��x�����@��ύX���ĉ������<br />"
End If

'---- �R���r�j�x������30���~�ȏ�̒����̓G���[
If wPaymentMethod = "�R���r�j�x��" And wOrderAmTotal > 300000 Then
	wMsg = wMsg & "�l�b�g�o���L���O�E�䂤����E�R���r�j�����̏ꍇ�A1��̂�������30���~�����x�ƂȂ�܂��B���������e���͂��x�����@��ύX���ĉ������<br />"		'2011/03/22 hn mod
End If

End Function

'========================================================================
'
'	Function	�����`�F�b�N
'
'		parm:		�z����X�֔ԍ�
'		return:	�����Ȃ�@wRitouFl = Y
'						�����ȊO�@wRitouFl = N
'				�������̗����Ȃ�@wRitouRitouFl = Y
'						�������̗����ȊO�@wRitouRitouFl = N
'
'========================================================================
Function check_ritou(pZip)

Dim vZip
Dim vSQL
Dim RSv

vZip = Replace(pZip, "-", "")

If vZip = "" Then
	wRitouFl  = "N"
	Exit Function
End If

'---- �����`�F�b�N
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    �X�֔ԍ�"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ���� WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    �X�֔ԍ� = '" & vZip & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = True Then
	wRitouFl = "N"
Else
	wRitouFl = "Y"
End If

RSv.Close

End Function

'========================================================================
'
'	Function	����萔���v�Z
'
'		�E(���i���z+����)*����łɉ������萔�������R���g���[���}�X�^������o��
'		�E����萔���𖳗��ɂ��鏤�i���z���v(vItemNum1)�����o��
'
'		parm:	(���i���z+����)*�����
'		return: ����萔��
'
'========================================================================
Function calc_cod_am(p_total_am)

Dim i
Dim vTotalAm()
Dim vCodAm()
Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2

'---- '����萔��
Call getCntlMst("��","����萔��","1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
Call cf_unstring(vItemChar1, vTotalAm, ",")
Call cf_unstring(vItemChar2, vCodAm, ",")

wTotal_NoDaibikiFee = vItemNum1

For i = 0 to UBound(vTotalAm) - 1
	If CCur(vTotalAm(i)) > CCur(p_total_am) Then
		Exit For
	End If
Next

calc_cod_am = vCodAm(i)

End function

'========================================================================
'
'	Function	�󒍖��ד��e�\��
'
'========================================================================
Function display_order_detail()

Dim vProductNm
Dim vPrice
Dim vBeforeRebateAm
Dim vSQL
Dim RSv
Dim vProdTermFl
Dim vInventoryCd
Dim vInventoryImage
Dim strLargeItem	'2013/10/21 GV # add

'---- ���󒍖��׏����o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    a.*"
vSQL = vSQL & "  , b.�{�x�X�R�[�h"
vSQL = vSQL & "  , b.�����敪"
vSQL = vSQL & "  , b.���菤�i��"
vSQL = vSQL & "  , b.�d�ʏ��i����"
vSQL = vSQL & "  , b.�}�X�^�[�J�[�g����"
vSQL = vSQL & "  , b.�q�Ɏw��Ȃ��t���O"
vSQL = vSQL & "  , b.ASK���i�t���O"
vSQL = vSQL & "  , b.�戵���~��"
vSQL = vSQL & "  , b.�p�ԓ�"
vSQL = vSQL & "  , b.������"
vSQL = vSQL & "  , b.�󏭐���"
vSQL = vSQL & "  , b.�Z�b�g���i�t���O"
vSQL = vSQL & "  , b.���[�J�[�������敪"
vSQL = vSQL & "  , b.Web�[����\���t���O"
vSQL = vSQL & "  , b.���ח\�薢��t���O"
vSQL = vSQL & "  , b.B�i�t���O"
vSQL = vSQL & "  , b.�����萔��"
vSQL = vSQL & "  , b.������󒍍ϐ���"
vSQL = vSQL & "  , b.��A�֎~�t���O "					'2013/10/21 GV # add
vSQL = vSQL & "  , b.����s�t���O "					'2013/10/21 GV # add
vSQL = vSQL & "  , c.�����\����"
vSQL = vSQL & "  , c.�����\���ח\���"
vSQL = vSQL & "  , c.B�i�����\����"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ���󒍖��� a WITH (NOLOCK)"
vSQL = vSQL & "  , Web���i b WITH (NOLOCK)"
vSQL = vSQL & "  , Web�F�K�i�ʍ݌� c WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        a.SessionID = '" & gSessionID & "'"
vSQL = vSQL & "    AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "    AND b.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "    AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "    AND c.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "    AND c.�F = a.�F"
vSQL = vSQL & "    AND c.�K�i = a.�K�i"
vSQL = vSQL & " ORDER BY"
vSQL = vSQL & "     a.�󒍖��הԍ�"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

wProductHtml = ""

'---- ����HTML�쐬
If RSv.EOF = True Then
	'-- �f�[�^���Ȃ�
	wProductHtml = wProductHtml & "    <tr>" & vbNewLine
	wProductHtml = wProductHtml & "      <td align=""center""><b>�J�[�g�ɏ��i������܂���B</b></td>" & vbNewLine
	wProductHtml = wProductHtml & "    </tr>" & vbNewLine
Else
	'-- �f�[�^������
	'----- ���o��
	wProductHtml = wProductHtml & "    <tr>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='maker'>���[�J�[</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='name'>���i��</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='stock'>�݌�</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='price'>�P��</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='number'>����</th>" & vbNewLine
	wProductHtml = wProductHtml & "      <th class='amount'>���z(�ō�)</th>" & vbNewLine
	wProductHtml = wProductHtml & "    </tr>"

	wPrdctAmTotal = 0
	wPrdctAmTotalNoTax = 0

	'---- ���i���ו�
	Do Until RSv.EOF = True

		'---- 2013.10.21 GV # add start
		'---- ��^���i�̕\��
		strLargeItem = ""
		If (((IsNull(RSv("��A�֎~�t���O")) = False) And (RSv("��A�֎~�t���O") = "Y")) And _
			((IsNull(RSv("����s�t���O")) = False) And (RSv("����s�t���O") = "Y")) And _
			(RSv("�����敪") = "�d�ʏ��i")) Then
			strLargeItem = strLargeItem & "<br><span style='color:red;'>��^���i</span>"
			wLargeItemFl = "Y"
		Else
			wNonLargeItemFl = "Y"
		End If
		'---- 2013.10.21 GV # add start

		vProductNm = RSv("���i��")
		If Trim(RSv("�F")) <> "" Then
			vProductNm = vProductNm & "/" & RSv("�F")
		End If
		If Trim(RSv("�K�i")) <> "" Then
			vProductNm = vProductNm & "/" & RSv("�K�i")
		End If

		vPrice = calcPrice(RSv("�󒍒P��"), wSalesTaxRate)
		wPrdctAmTotal = wPrdctAmTotal + (vPrice * RSv("�󒍐���"))
		wPrdctAmTotalNoTax = wPrdctAmTotalNoTax + (Fix(RSv("�󒍒P��")) * RSv("�󒍐���"))

		wProductHtml = wProductHtml & "    <tr>" & vbNewLine
		wProductHtml = wProductHtml & "      <td>" & RSv("���[�J�[��") & "</td>" & vbNewLine
'---- 2013/10/21 GV # mod start
'		wProductHtml = wProductHtml & "      <td><a href='" & g_HTTP & "shop/ProductDetail.asp?Item=" & RSv("���[�J�[�R�[�h") & "^" & Server.URLEncode(RSv("���i�R�[�h")) & "^" & RSv("�F") & "^" & RSv("�K�i") & "' alt=''>" & vProductNm & "</a></td>" & vbNewLine
		wProductHtml = wProductHtml & "      <td><a href='" & g_HTTP & "shop/ProductDetail.asp?Item=" & RSv("���[�J�[�R�[�h") & "^" & Server.URLEncode(RSv("���i�R�[�h")) & "^" & RSv("�F") & "^" & RSv("�K�i") & "' alt=''>" & vProductNm & "</a>" & strLargeItem & "</td>" & vbNewLine
'---- 2013/10/21 GV # mod end
		
		'------------- �݌�
		vProdTermFl = "N"
		If IsNull(RSv("�戵���~��")) = False Then	'�戵���~
			vProdTermFl = "Y"
		End If
		If IsNull(RSv("�p�ԓ�")) = False And RSv("�����\����") <= 0 Then	'�p�Ԃō݌ɖ���
			vProdTermFl = "Y"
		End If
		If IsNull(RSv("������")) = False Then		'�������i
			vProdTermFl = "Y"
		End If

		'---- �݌ɏ�
		vInventoryCd = GetInventoryStatus(RSv("���[�J�[�R�[�h"), RSv("���i�R�[�h"), RSv("�F"), RSv("�K�i"), RSv("�����\����"), RSv("�󏭐���"), RSv("�Z�b�g���i�t���O"), RSv("���[�J�[�������敪"), RSv("�����\���ח\���"), vProdTermFl)

		'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
		Call GetInventoryStatus2(RSv("�����\����"), RSv("Web�[����\���t���O"), RSv("���ח\�薢��t���O"), RSv("�p�ԓ�"), RSv("B�i�t���O"), RSv("B�i�����\����"), RSv("�����萔��"), RSv("������󒍍ϐ���"), vProdTermFl, vInventoryCd, vInventoryImage)

		'----- �݌ɏ󋵕\��
		If IsNull(RSv("�戵���~��")) = False Or _
		   IsNull(RSv("������")) = False Or _
		   (RSv("B�i�t���O") = "Y" And RSv("B�i�����\����") <= 0) Or _
		   (IsNull(RSv("�p�ԓ�")) = False And RSv("�����\����") <= 0) Then
			wProductHtml = wProductHtml & "      <td><span class='stock'>&nbsp</span></td>" & vbNewLine
		Else
			'---- �������łȂ��ꍇ�̂݁A�݌ɏ󋵂�\��
			wProductHtml = wProductHtml & "      <td><span class='stock'><img src='images/" & vInventoryImage & "' alt=''>" & vInventoryCd & "</span></td>" & vbNewLine
		End If

		wProductHtml = wProductHtml & "      <td>" & FormatNumber(vPrice, 0) & "�~</td>"& vbNewLine
		wProductHtml = wProductHtml & "      <td>" & RSv("�󒍐���") & "</td>" & vbNewLine
		wProductHtml = wProductHtml & "      <td>" & FormatNumber(vPrice * RSv("�󒍐���"), 0) & "�~</td>" & vbNewLine
		wProductHtml = wProductHtml & "    </tr>" & vbNewLine

		RSv.MoveNext
	Loop

	'---- ���i���v���z
	wProductHtml = wProductHtml & "    <tr>" & vbNewLine
	wProductHtml = wProductHtml & "      <td colspan='6'>" & vbNewLine
	wProductHtml = wProductHtml & "        <dl class='total'>" & vbNewLine
	wProductHtml = wProductHtml & "          <dt>���i���v�i�ō��j</dt><dd>" & FormatNumber(wPrdctAmTotal, 0) & "�~</dd>" & vbNewLine

	'---- �����v�Z
	Call fCalcShipping(gSessionID, "�ʏ�", wShippingNoTax, wFreightForwarder, wSokoCnt, wKoguchi)
	vPrice = Fix(wShippingNoTax * (100 + wSalesTaxRate) / 100)

	If wRitouFl = "Y" Then
		wProductHtml = wProductHtml & "          <dt>�����i�ō��j�i���u�n�j</dt><dd>" & FormatNumber(vPrice, 0) & "�~</dd>" & vbNewLine
	Else
		wProductHtml = wProductHtml & "          <dt>�����i�ō��j</dt><dd>" & FormatNumber(vPrice, 0) & "�~</dd>" & vbNewLine
	End If

	'---- ����萔���v�Z
	wCodAm = 0
	If wPaymentMethod = "�����" THen
		wCodAm = calc_cod_am((wPrdctAmTotal + wShippingNoTax) * (wSalesTaxRate + 100) / 100) * wSokoCnt

		if CCur(wPrdctAmTotal) >= CCur(wTotal_NoDaibikiFee) then
			wCodAm = 0
		end if
	End If

	'---- ���x�[�g���z�v�Z
	'---- �萔���͖������ă`�F�b�N���A�x�������z��0�~�ɂȂ�����x�����@�@���Ȃ��Ɂi�\���̂݁j
	wOrderAmTotal = wPrdctAmTotal + ((wShippingNoTax + wCodAm) * (wSalesTaxRate + 100) / 100)
	vBeforeRebateAm = wOrderAmTotal

	If wRebateFl = "Y" Then
		' �����ߕs�����z �� ����萔���Ȃ��̎󒍋��z�i���i���z�{�����j
		If wCustomerKabusokuAm >= (wOrderAmTotal - (wCodAm * (wSalesTaxRate + 100) / 100)) Then
			wCustomerKabusokuAm = (wOrderAmTotal - (wCodAm * (wSalesTaxRate + 100) / 100))
			vBeforeRebateAm = wCustomerKabusokuAm
			wOrderAmTotal = 0
			wCodAm = 0
		Else
			wOrderAmTotal = wOrderAmTotal - wCustomerKabusokuAm
		End If
	End If

	'---- ����萔��
	If wPaymentMethod = "�����" Then
		vPrice = Fix(wCodAm * (100 + wSalesTaxRate) / 100)
		If wCodAm = 0 Then
			wProductHtml = wProductHtml & "          <dt>����萔���i�ō��j</dt><dd>" & "����</dd>" & vbNewLine
		Else
			wProductHtml = wProductHtml & "          <dt>����萔���i�ō��j</dt><dd>" & FormatNumber(vPrice, 0) & "�~</dd>" & vbNewLine
		End If
	End If

	'------------- �w�����v
	wProductHtml = wProductHtml & "          <dt>���w�����v���z�i�ō��j</dt><dd>" & FormatNumber(vBeforeRebateAm,0) & "�~</dd>" & vbNewLine

	'------------- �����
	wTax = vBeforeRebateAm - (wPrdctAmTotalNoTax + wShippingNoTax + wCodAm)
	wProductHtml = wProductHtml & "          <dt class='normalweight'>�������</dt><dd>" & FormatNumber(wTax, 0) & "�~</dd>" & vbNewLine

	'---- ���x�[�g
	If wRebateFl = "Y" Then
		vPrice = wCustomerKabusokuAm * -1
		wProductHtml = wProductHtml & "          <dt class='credit'>�N���W�b�g�^�ߕs����</dt><dd>" &  FormatNumber(vPrice,0) & "�~</dd>" & vbNewLine

		'------------- �x�����v
		wProductHtml = wProductHtml & "          <dt>���x�������v���z�i�ō��j</dt><dd>" & FormatNumber(wOrderAmTotal, 0) & "�~</dd>" & vbNewLine

		wProductHtml = wProductHtml & "        </dl>" & vbNewLine

		'------------- ���x�[�g�g�p���b�Z�[�W
		wProductHtml = wProductHtml & "        <div class='contact'>��L�N���W�b�g/�ߕs�����́A���̂������E�����ς�݂̂ɏ[������܂��B<br>�L�����Z�����Ă����p�ɂȂ�Ȃ��ꍇ�͕��Љc�ƈ��܂ł��A�����������B</div>" & vbNewLine
	Else
		wProductHtml = wProductHtml & "        </dl>" & vbNewLine
	End If

	wProductHtml = wProductHtml & "      </td>" & vbNewLine
	wProductHtml = wProductHtml & "    </tr>" & vbNewLine

End If

RSv.Close

End function
'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>���������e�̊m�F�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
<script type="text/javascript">
//=====================================================================
//	Next onClick
//=====================================================================
function next_onClick(){

	if (document.f_data.payment_method.value == "�N���W�b�g�J�[�h"){
		document.f_data.action = "OrderCardEnter.asp";
	}else{
		document.f_data.action = "OrderProcessing.asp";
	}
	document.f_data.submit();
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
      <li class="now">���������e�̊m�F</li>
    </ul>
  </div></div></div>

  <h1 class="title">���������e�̊m�F</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.�V���b�s���O�J�[�g" width="170" height="50"></li>
    <li><img src="images/step02.gif" alt="2.���͂���A���x�����@�̑I��" width="170" height="50"></li>
    <li><img src="images/step03_now.gif" alt="3.���������e�̊m�F" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.����������" width="170" height="50"></li>
  </ol>

  <h2 class="cart_title">�J�[�g���e</h2>
  <table id="cart" class="confirm">
            <!---- �������i�ꗗ start ---->
<% = wProductHtml %>
            <!---- �������i�ꗗ end ---->
            <!-- �z���w�� start -->
<% = wHaisouHtml %>
            <!-- �z���w�� end -->
            <!-- �x�����@ start -->
<% = wPaymentHtml %>
            <!-- �x�����@ end -->
            <!-- �̎��� start -->
<% = wReceiptHtml %>
            <!-- �̎��� end -->
            <!-- ��^���i start -->
<% = wLargeItemHtml %>
            <!-- ��^���i end -->
  </table>

  <div id="btn_box">
    <ul class="btn">
      <li><a href="OrderInfoEnter.asp"><img src="images/btn_fix.png" alt="���e��ύX����" class="opover"></a></li>
      <li class="last"><a href="JavaScript:next_onClick();"><img src="images/btn_send.png" alt="���@�M" class="opover"></a></li>
    </ul>
  </div>

  <p class="caution">�����M�{�^����2�x�����Ȃ��悤�ɂ��肢���܂��B</p>
  <ul class="info left">
    <li><a href="../guide/change.asp">���������i�̃L�����Z���E�ԕi�ɂ���</a></li>
    <li><a href="../guide/nouki.asp">���i�̔[���ɂ��Ă͂�����</a></li>
  </ul>

  <form method="post" name="f_data" action="">
    <input type="hidden" name="OrderTotalAm" value="<% = wOrderAmTotal %>">
    <input type="hidden" name="payment_method" value="<% = wPaymentMethod %>">
    <input type="hidden" name="Skey" value="<% = Skey %>">
  </form>

<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>