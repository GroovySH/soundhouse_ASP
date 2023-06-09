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
'	�V���b�v CF�ǂ��ƃN���W�b�g �C���^�[�t�F�[�X
'
'�X�V����
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT��)
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/07/13 if-web ���j���[�A�����C�A�E�g����
'2013/10/01 GV # �Z�f�B�i�V�X�e���ڍs�Ή�
'
'========================================================================

On Error Resume Next

Const store	= "160437002000000"	'CF store code 111111111111111 (for test)

Dim userID
Dim userName
Dim msg

Dim order_no
Dim customer_nm
Dim furigana
Dim customer_email
Dim zip
Dim prefecture
Dim address
Dim telephone
Dim loan_am
Dim continue
Dim order_estimate

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add
Dim sedyna_url '2013/10/01 GV # add

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
	wErrDesc = "CfApply.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
order_no = ReplaceInput(Request("order_no"))
order_estimate = ReplaceInput(Request("order_estimate"))

if IsNumeric(order_no) = false then
	Response.redirect g_HTTP
end if

'---- �󒍏����o��
w_sql = ""
w_sql = w_sql & "SELECT a.�ڋq��"
w_sql = w_sql & "       , a.�ڋq�t���K�i"
w_sql = w_sql & "       , a.�ڋqE_mail1"
w_sql = w_sql & "       , b.�ڋq�X�֔ԍ�"
w_sql = w_sql & "       , b.�ڋq�s���{��"
w_sql = w_sql & "       , b.�ڋq�Z��"
w_sql = w_sql & "       , c.�ڋq�d�b�ԍ�"
w_sql = w_sql & "       , d.�󒍍��v���z"
w_sql = w_sql & "       , d.���[������"
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

customer_nm = RS("�ڋq��")
furigana = RS("�ڋq�t���K�i")
customer_email = RS("�ڋqE_mail1")
zip = RS("�ڋq�X�֔ԍ�")
prefecture = RS("�ڋq�s���{��")
address = RS("�ڋq�Z��")
telephone = RS("�ڋq�d�b�ԍ�")
loan_am = RS("�󒍍��v���z") - RS("���[������") 

RS.close

if order_estimate = "������" then
	continue = "1"
	'2013/10/01 GV # add start
	'������t
	sedyna_url = "https://c-web.cedyna.co.jp/customer/action/ssAA01/WAA0101Action/RWAA010101"
	'2013/10/01 GV # add end
else
	continue = "0"
	'2013/10/01 GV # add start
	'�V�~�����[�g�P��
	sedyna_url = "https://c-web.cedyna.co.jp/customer/action/ssAA01/WAA0106Action/RWAA010601"
	'2013/10/01 GV # add end
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
<title>�Z�f�B�i�ďo���b�T�E���h�n�E�X</title>
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
      <li class="now">�Z�f�B�i�ďo����</li>
    </ul>
  </div></div></div>

  <h1 class="title">�Z�f�B�i�ďo����</h1>

  <p>���΂炭���҂����������B</p>
<%
' 2013/10/01 GV # mod 
'  <form name="f_cf" method="post" action="https://cf.ufit.ne.jp/dotcredit/simulate/simulate.asp">
%>
  <form name="f_cf" method="post" action="<%=sedyna_url%>">
    <input type="hidden" name="store" value="<%=store%>">
    <input type="hidden" name="amount" value="<%=loan_am%>">
    <input type="hidden" name="continue" value="<%=continue%>">
    <input type="hidden" name="labor" value="0">
    <input type="hidden" name="item1" value="�����@��">
    <input type="hidden" name="item1count" value="1">
    <input type="hidden" name="item1amount" value="<%=loan_am%>">
    <input type="hidden" name="tranno" value="<%=order_no%>">
    <input type="hidden" name="namekn" value="<%=furigana%>">
    <input type="hidden" name="namekj" value="<%=customer_nm%>">
    <input type="hidden" name="zip" value="<%=zip%>">
    <input type="hidden" name="address" value="<%=prefecture%><%=address%>">
    <input type="hidden" name="tel" value="<%=telephone%>">
    <input type="hidden" name="mail" value="<%=customer_email%>">
    <input type="hidden" name="bonusdeal" value="100">
    <input type="hidden" name="twobonusdeal" value="200">
    <input type="hidden" name="result" value="1">
  </form>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	document.f_cf.submit();		//CF�y�[�W�փW�����v
</script>
</body>
</html>