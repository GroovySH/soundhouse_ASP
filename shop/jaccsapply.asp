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
'	JACCS �C���^�[�t�F�[�X
'
'�X�V����
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/07/13 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Dim userID
Dim userName
Dim msg

Dim order_no
Dim customer_nm
Dim customer_email

Dim ShippingAm
Dim SalesTaxAm
Dim OrderTotalAm
Dim ProductName(30)
Dim ProductPrice(30)
Dim ProductQt(30)
Dim ProductCnt

Dim wProductHTML

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
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
	wErrDesc = "jaccsapply.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

if IsNumeric(order_no) = false then
	Response.redirect g_HTTP
end if

'---- �󒍏����o��
w_sql = ""
w_sql = w_sql & "SELECT a.�ڋq��"
w_sql = w_sql & "     , a.�ڋqE_mail1"
w_sql = w_sql & "     , d.����"
w_sql = w_sql & "     , d.�O�ō��v���z"
w_sql = w_sql & "     , d.�󒍍��v���z"
w_sql = w_sql & "     , e.���[�J�[��"
w_sql = w_sql & "     , e.���i��"
w_sql = w_sql & "     , e.�F"
w_sql = w_sql & "     , e.�K�i"
w_sql = w_sql & "     , e.�󒍐���"
w_sql = w_sql & "     , e.�󒍒P��"
w_sql = w_sql & "  FROM Web�ڋq a WITH (NOLOCK)"
w_sql = w_sql & "     , Web�� d WITH (NOLOCK)"
w_sql = w_sql & "     , Web�󒍖��� e WITH (NOLOCK)"
w_sql = w_sql & " WHERE d.�󒍔ԍ� = " & order_no
w_sql = w_sql & "   AND a.�ڋq�ԍ� = d.�ڋq�ԍ�"
w_sql = w_sql & "   AND e.�󒍔ԍ� = d.�󒍔ԍ�"
w_sql = w_sql & " ORDER BY e.�󒍖��הԍ�"
	  
'@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

customer_nm = RS("�ڋq��")
customer_email = RS("�ڋqE_mail1")
ShippingAm = RS("����")
SalesTaxAm = RS("�O�ō��v���z")
OrderTotalAm = RS("�󒍍��v���z")

wProductHTML = ""
ProductCnt = 0

Do until RS.EOF = true or ProductCnt = 30
	ProductCnt = ProductCnt + 1
	wProductHTML = wProductHTML & "<input type='hidden' name='SINA_" & Right("0" & CStr(ProductCnt), 2) & "_SINAMEI' value='" & RS("���i��") & "'>" & vbNewLine
	wProductHTML = wProductHTML & "<input type='hidden' name='SINA_" & Right("0" & CStr(ProductCnt), 2) & "_SURYOU' value='" & RS("�󒍐���") & "'>" & vbNewLine
	wProductHTML = wProductHTML & "<input type='hidden' name='SINA_" & Right("0" & CStr(ProductCnt), 2) & "_TANKA' value='" & FIX(RS("�󒍒P��")) & "'>" & vbNewLine & vbNewLine
	RS.MoveNext
Loop

RS.close

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
<title>JACCS�ďo���b�T�E���h�n�E�X</title>
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
      <li class="now">JACCS�ďo����</li>
    </ul>
  </div></div></div>

  <h1 class="title">JACCS�ďo����</h1>

  <p>���΂炭���҂����������B</p>

  <form name="f_JACCS" method="post" action="<%=g_JACCS_URL%>">
    <input type="hidden" name="KAMEITEN_BANGO" value="<%=g_JACCS_ID%>"> <!-- �����X���ʔԍ� -->
    <input type="hidden" name="NAME" value="<%=customer_nm%>"> <!-- ���� -->
    <input type="hidden" name="EMAIL" value="<%=customer_email%>"> <!-- ���[���A�h���X -->
    <input type="hidden" name="DENPYO_BANGO" value="<%=order_no%>"> <!-- �`�[�ԍ� -->

<!-- ���i���@�i���A���ʁA�P�� -->
<%=wProductHTML%>

    <input type="hidden" name="SYOHIZEI" value="<%=SalesTaxAm%>"> <!-- ����� -->
    <input type="hidden" name="SOURYOU" value="<%=ShippingAm%>"> <!-- ���� -->
    <input type="hidden" name="GOUKEIGAKU" value="<%=OrderTotalAm%>"> <!-- ���v���z�i�ō��j -->
  </form>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript">
	document.f_JACCS.submit();		//JACCS�y�[�W�փW�����v
</script>
</body>
</html>