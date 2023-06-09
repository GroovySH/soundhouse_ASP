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
'	�V���b�v �I���R �N���b�N�I�� �C���^�[�t�F�[�X
'
'�X�V����
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/07/13 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Const kamei_no	= "06218226"	'�����X�ԍ�
'Const kyaku_syu	= "004"	'�戵�_��ԍ�
Const kyaku_syu	= "005"	'�戵�_��ԍ�	'2005/07/05 na mod
Const buten_cd	= "519"	'���X�R�[�h
Const OricoURL	= "https://www2.orico.co.jp/webcredit/sp/top.asp"	'�{��
'@@@Const OricoURL	= "https://www2.orico.co.jp/webcredit/sp/simulation.asp"	'�e�X�g

Dim userID
Dim userName
Dim msg

Dim order_no
Dim order_estimate

Dim CustomerName
Dim CustomerEmail
Dim Zip
Dim Prefecture
Dim Address
Dim Telephone
Dim ProductAm
Dim ShippingAm
Dim DownPaymentAm
Dim OrderTotalAm
Dim Continue
Dim SalesTaxRate

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_error_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.Expires = -1			' Do not cache

'---- UserID ���o��
userID = Session("userID")
userName = Session("userName")

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OricoApply.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

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

'---- ��M�f�[�^�[�̎��o��
order_no = Clng(ReplaceInput(Request("order_no")))
order_estimate = ReplaceInput(Request("order_estimate"))

if IsNumeric(order_no) = false then
	Response.redirect g_HTTP
end if

'---- �󒍏����o��
w_sql = ""
w_sql = w_sql & "SELECT a.�ڋq��"
w_sql = w_sql & "     , a.�ڋqE_mail1"
w_sql = w_sql & "     , b.�ڋq�X�֔ԍ�"
w_sql = w_sql & "     , b.�ڋq�s���{��"
w_sql = w_sql & "     , b.�ڋq�Z��"
w_sql = w_sql & "     , c.�ڋq�d�b�ԍ�"
w_sql = w_sql & "     , d.���i���v���z"
w_sql = w_sql & "     , d.����"
w_sql = w_sql & "     , d.�󒍍��v���z"
w_sql = w_sql & "     , d.���[������"
w_sql = w_sql & "     , d.����ŗ�"
w_sql = w_sql & "  FROM Web�ڋq a WITH (NOLOCK)"
w_sql = w_sql & "     , Web�ڋq�Z�� b WITH (NOLOCK)"
w_sql = w_sql & "     , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
w_sql = w_sql & "     , Web�� d WITH (NOLOCK)"
w_sql = w_sql & " WHERE d.�󒍔ԍ� = " & order_no
w_sql = w_sql & "   AND a.�ڋq�ԍ� = d.�ڋq�ԍ�"
w_sql = w_sql & "   AND b.�ڋq�ԍ� = a.�ڋq�ԍ�"
w_sql = w_sql & "   AND b.�Z���A�� = 1"
w_sql = w_sql & "   AND c.�ڋq�ԍ� = a.�ڋq�ԍ�"
w_sql = w_sql & "   AND c.�Z���A�� = 1"
w_sql = w_sql & "   AND c.�d�b�A�� = 1"
	  
'@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

CustomerName = RS("�ڋq��")
CustomerEmail = RS("�ڋqE_mail1")
Zip = RS("�ڋq�X�֔ԍ�")
Prefecture = RS("�ڋq�s���{��")
Address = RS("�ڋq�Z��")
Telephone = RS("�ڋq�d�b�ԍ�")
SalesTaxRate = Ccur(RS("����ŗ�"))
ShippingAm = Fix(RS("����") * (100 + SalesTaxRate) / 100)
DownPaymentAm = RS("���[������")
OrderTotalAm = RS("�󒍍��v���z")
ProductAm = OrderTotalAm - ShippingAm		'����Œ����̂��ߏ��i���v���z�͎g�p���Ȃ�

RS.close

if order_estimate = "������" then
	Continue = "1"
else
	Continue = "0"
end if

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

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>�I���R�ďo���b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
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
      <li class="now">�I���R�ďo��</li>
    </ul>
  </div></div></div>

  <h1 class="title">�I���R�ďo��</h1>

  <p>���΂炭���҂����������B</p>

  <form name="f_cf" method="post" action="<%=OricoURL%>">
    <input type="hidden" name="kamei_no" value="<%=kamei_no%>">
    <input type="hidden" name="kyaku_syu" value="<%=kyaku_syu%>">
    <input type="hidden" name="buten_cd" value="<%=buten_cd%>">
    <input type="hidden" name="back_url" value="http://www.soundhouse.co.jp">
    <input type="hidden" name="pr_num" value="<%=order_no%>">
    <input type="hidden" name="brand_mei1" value="�����@��">
    <input type="hidden" name="brand_suu1" value="1">
    <input type="hidden" name="brand_kin1" value="<%=ProductAm%>">
    <input type="hidden" name="brand_gokei" value="<%=ProductAm%>">
    <input type="hidden" name="soryo_gokei" value="<%=ShippingAm%>">
    <input type="hidden" name="loan_kin" value="<%=OrderTotalAm%>">
    <input type="hidden" name="head_kin" value="<%=DownPaymentAm%>">
    <input type="hidden" name="h_name" value="<%=CustomerName%>">
    <input type="hidden" name="h_yubin" value="<%=Zip%>">
    <input type="hidden" name="h_addr1" value="<%=Prefecture & Address%>">
    <input type="hidden" name="h_addr2" value="">
    <input type="hidden" name="h_telno" value="<%=Telephone%>">
  </form>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	document.f_cf.submit();		//Orico�y�[�W�փW�����v
</script>
</body>
</html>