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
'	�R���r�j�x�����N�G�X�g���� (eContext)
'
'------------------------------------------------------------------------
'	�X�V����
'2008/04/28 ���x�[�g�Ή��̂��ߍ��v���z�̂݃Z�b�g
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2009/04/30 �G���[����error.asp�ֈړ�
'2010/09/27 hn eContext����̖߂�l�̃`�F�b�N����
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim CustomerTel
Dim CustomerName
Dim CustomerEmail
Dim Shipping
Dim OrderTotal
Dim SalesTax
Dim OrderDate
Dim ItemTotal
Dim OrderNo
Dim wSalesTaxRate
Dim wPrice

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim OBJeContext
Dim eContextRtn
Dim eContextNo
Dim eConF
Dim eConK

Dim wSQL
Dim wHTML
Dim wMSG
Dim wNextURL
Dim wErrDesc   '2011/08/01 an add

'=======================================================================

userID = Session("UserID")

Session("msg") = ""
wMSG = ""

'---- execute main process
call ConnectDB()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderEcontext.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- �G���[�������Ƃ��͒����o�^�����y�[�W�A�G���[������Ίm�F�y�[�W��
if wMSG = "" then
''response.write("NO=" & eContextNo & "F=" & eConF & "  k=" & eConK)
	Response.Redirect "OrderSubmit.asp?OrderNo=" & OrderNo & "&eConF=" & eConF & "&eConK=" & eConK
else
	Session("msg") = wMSG
	Response.Redirect "OrderInfoEnter.asp"
end if

'========================================================================
'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function ConnectDB()

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

Dim vTemp

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- ���󒍎��o��
wSQL = ""
wSQL = wSQL & "SELECT a.���i���v���z"
wSQL = wSQL & "     , a.����"
wSQL = wSQL & "     , a.����萔��"
wSQL = wSQL & "     , a.�R���r�j�x���萔��"
wSQL = wSQL & "     , a.�O�ō��v���z"
wSQL = wSQL & "     , a.�󒍍��v���z"
wSQL = wSQL & "     , a.�ڋq�d�b�ԍ�"
wSQL = wSQL & "     , a.�ڋqE_mail"
wSQL = wSQL & "     , a.eContext��t�ԍ�"
wSQL = wSQL & "     , b.�ڋq��"
wSQL = wSQL & "  FROM ���� a"
wSQL = wSQL & "     , Web�ڋq b"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
wSQL = wSQL & "   AND b.�ڋq�ԍ� = a.�ڋq�ԍ�"
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = true then
	wMSG = "<font color='#ff0000'>NoData</font>"
	exit function
end if

CustomerTel = cf_numeric_only(RS("�ڋq�d�b�ԍ�"))
CustomerName = RS("�ڋq��")
CustomerEmail = RS("�ڋqE_mail")
'Shipping = (RS("����") + RS("����萔��") + RS("�R���r�j�x���萔��")) * (100 + wSalesTaxRate) / 100
OrderTotal = RS("�󒍍��v���z")
'SalesTax = RS("�O�ō��v���z")
OrderDate = cf_FormatDate(Now(), "YYYY/MM/DD") & " " & cf_FormatTime(Now(), "HH:MM:SS")
'ItemTotal = RS("�󒍍��v���z") - Shipping
ItemTotal = OrderTotal

OrderNo = GetOrderNo()              '�����ԍ�

'---- eContext ���N�G�X�g
Set OBJeContext = CreateObject("FormPost.Https")

vTemp = OBJeContext.init()

vTemp = OBJeContext.set("shopID", g_eContext_ID)
vTemp = OBJeContext.set("orderID", OrderNo)
vTemp = OBJeContext.set("sessionID", "1")
vTemp = OBJeContext.set("telNo", CustomerTel)
vTemp = OBJeContext.set("kanjiName1_1", CustomerName)
vTemp = OBJeContext.set("kanjiName1_2", "�@")
vTemp = OBJeContext.set("email", CustomerEmail)
vTemp = OBJeContext.set("paymentFlg", "0")
vTemp = OBJeContext.set("shippmentFlg", "2")
'vTemp = OBJeContext.set("commission", Shipping)
vTemp = OBJeContext.set("commission", 0)
vTemp = OBJeContext.set("ordAmount", OrderTotal)
vTemp = OBJeContext.set("ordAmountbfTax", "0")
'vTemp = OBJeContext.set("ordAmountTax", SalesTax)
vTemp = OBJeContext.set("ordAmountTax", 0)
vTemp = OBJeContext.set("ordItemNo", "1")
vTemp = OBJeContext.set("orderDate", OrderDate)
vTemp = OBJeContext.set("siteInfo", "�̎������l���L�q")

vTemp = OBJeContext.set("itemName1", "�������ꎮ(�ō���)")
vTemp = OBJeContext.set("unitPrice1", ItemTotal)
vTemp = OBJeContext.set("ordUnit1", "1")
vTemp = OBJeContext.set("unitChar1", "��")
vTemp = OBJeContext.set("dtlAmount1", ItemTotal)
vTemp = OBJeContext.set("goodsCode1", "0")

'---- ���N�G�X�g
vTemp = OBJeContext.send3(g_eContext_URL)

'---- �߂�l�̃`�F�b�N
call checkError(vTemp)

'---- �󒍏���eContext��t�ԍ����Z�b�g
if wMSG = "" then
	RS("eContext��t�ԍ�") = eContextNo
	RS.update
end if


'----
vTemp = OBJeContext.finally()
Set OBJeContext = Nothing

RS.close

End Function

'========================================================================
'
'	Function	�󒍔ԍ����o��
'
'========================================================================
'
Function GetOrderNo()

Dim vRS_Cntl

'---- �R���g���[���}�X�^���o��
wSQL = ""
wSQL = wSQL & "SELECT item_num1"
wSQL = wSQL & "  FROM �R���g���[���}�X�^"
wSQL = wSQL & " WHERE sub_system_cd = '����'"
wSQL = wSQL & "   AND item_cd = '�ԍ�'"
wSQL = wSQL & "   AND item_sub_cd = 'Web��'"
	  
Set vRS_Cntl = Server.CreateObject("ADODB.Recordset")
vRS_Cntl.Open wSQL, Connection, adOpenStatic, adLockOptimistic

vRS_Cntl("item_num1") = Clng(vRS_Cntl("item_num1")) + 1
GetOrderNo = vRS_Cntl("item_num1")

vRS_Cntl.update
vRS_Cntl.close

End function

'========================================================================
'
'	Function	�J�[�h�G���[�`�F�b�N
'
'========================================================================
'
Function checkError(pRtn)

Dim vTemp

eContextRtn = Split(pRtn,chr(10))			'�߂�l���s�P�ʂɕ���
vTemp = Split(eContextRtn(0)," ")			'line1��' '�ŕ���

'---- �߂�R�[�h�`�F�b�N�@�@����
'2010/09/27 hn mod s
If (vTemp(0) = "1") Then
	eContextNo = Replace(eContextRtn(1), chr(13), "")		'��t�ԍ�
	eConF = Replace(eContextRtn(2), chr(13), "")				'�U�荞�ݕ[URL
	eConK = Replace(eContextRtn(7), chr(13), "")				'���ϑI��pURL
end if

if (vTemp(0) = "1") AND eContextNo <> "" AND eConF <> "" AND eConK <> "" then
	wMSG = ""
else
	'---- �G���[
	wMSG = "<font color='#ff0000'>" _
				& "�\���󂲂����܂��񂪤�������ɃG���[���������܂����<br>" _
				& "Code: " & eContextRtn(0) & "<br>" _
				& "������x�����������������A���̂��x�����@��I�����������<br>�悭���邲�����<a href='" & G_HTTP & "information/t_qanda.htm#card'>������</a>" _
				& "</font>"
end if
'2010/09/27 hn mod e

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
