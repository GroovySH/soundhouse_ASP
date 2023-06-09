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

<%
'========================================================================
'
'	�J�[�g�|�b�v�A�b�v�y�[�W
'�X�V����
'2005/02/21 hn ASK�\���𖳂��ɂ���
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/20 an SELECT����LAC�N�G���[�Ă�K�p
'
'========================================================================

On Error Resume Next

Dim msg

Dim Connection
Dim RS

Dim wTotalCnt
Dim wTotalAm
Dim wSalesTaxRate
Dim wPrice

Dim wListHTML

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim w_sql
Dim w_html
Dim w_error_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.Expires = -1			' Do not cache

'---- �Ăяo�����v���O��������̃��b�Z�[�W���o��

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "MiniCart.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	���i���׍쐬
'
'========================================================================
Function main()
Dim i

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- ���󒍃f�[�^SELECT

w_sql = ""
w_sql = w_sql & "SELECT a.���[�J�[�R�[�h"
w_sql = w_sql & "     , a.���i�R�[�h"
w_sql = w_sql & "     , a.���[�J�[��"
w_sql = w_sql & "     , a.���i��"
w_sql = w_sql & "     , a.�F"
w_sql = w_sql & "     , a.�K�i"
w_sql = w_sql & "     , a.�󒍐���"
w_sql = w_sql & "     , a.�󒍒P��" 
w_sql = w_sql & "     , a.�󒍋��z" 
'w_sql = w_sql & "     , b.ASK���i�t���O"            '2012/01/20 an del
w_sql = w_sql & "  FROM ���󒍖��� a WITH (NOLOCK)"  '2012/01/20 an mod
'w_sql = w_sql & "     , Web���i b"
w_sql = w_sql & " WHERE a.SessionID = '" & gSessionID & "'"       '2011/04/14 hn mod
'w_sql = w_sql & "   AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"     '2012/01/20 an del
'w_sql = w_sql & "   AND b.���i�R�[�h = a.���i�R�[�h"             '2012/01/20 an del
w_sql = w_sql & " ORDER BY a.�󒍖��הԍ�"

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

'---- ����HTML�쐬
wTotalCnt = 0
wTotalAm = 0

w_html = ""

if RS.EOF = true then
	w_html = w_html & "<br><center><span class='honbun'>�J�[�g�ɏ��i������܂���B</span></center>"
else
	w_html = w_html & "<table bgcolor='#000000' border='0' width='100%' cellpadding='0' cellspacing='1'>" & vbNewLine
	w_html = w_html & "<tr>" & vbNewLine
	w_html = w_html & "<td>" & vbNewLine
	w_html = w_html & "<table bgcolor='#ffffff' border='0' class='small' width='100%' cellpadding='0' cellspacing='2'>" & vbNewLine

	Do While RS.EOF = false
	'----- ���[�J�[�A���i��
		w_html = w_html & "  <tr>" & vbNewLine
		w_html = w_html & "    <td align='left' valign='middle' colspan='3'><font size='-1'>" & RS("���[�J�[��") & " <a href='JavaScript:Product_onClick(""Item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "^" & Trim(RS("�F")) & "^" & Trim(RS("�K�i")) & """)'>" & RS("���i��")
		if Trim(RS("�F")) <> "" then
			w_html = w_html & "/" & Trim(RS("�F"))
		end if
		if Trim(RS("�K�i")) <> "" then
			w_html = w_html & "/" & Trim(RS("�K�i"))
		end if
		w_html = w_html & "</a></font></td>" & vbNewLine
		w_html = w_html & "  </tr>" & vbNewLine

	'----- ���ʁA�P���A���z
		w_html = w_html & "  <tr>" & vbNewLine
		w_html = w_html & "    <td align='left' valign='middle' nowrap>����: " & RS("�󒍐���") & "</td>" & vbNewLine

'@@@@2005/02/21 change start
		wPrice = calcPrice(RS("�󒍒P��"), wSalesTaxRate)
		w_html = w_html & "    <td align='left' valign='middle' nowrap>�P��: " & FormatNumber(wPrice,0) & "�~</td>" & vbNewLine
		w_html = w_html & "    <td align='right' valign='middle' nowrap>���z: " & FormatNumber((wPrice * RS("�󒍐���")),0) & "�~</td>" & vbNewLine

'		if RS("ASK���i�t���O") = "Y" then
'			wPrice = 0
'			w_html = w_html & "    <td align='left' valign='middle' nowrap>�P��: ASK</td>" & vbNewLine
'			w_html = w_html & "    <td align='right' valign='middle' nowrap>���z: ASK</td>" & vbNewLine
'		else
'			wPrice = calcPrice(RS("�󒍒P��"), wSalesTaxRate)
'			w_html = w_html & "    <td align='left' valign='middle' nowrap>�P��: " & FormatNumber(wPrice,0) & "�~</td>" & vbNewLine
'			w_html = w_html & "    <td align='right' valign='middle' nowrap>���z: " & FormatNumber((wPrice * RS("�󒍐���")),0) & "�~</td>" & vbNewLine
'		end if
'@@@@ 2005/02/21 change end

		w_html = w_html & "  </tr>" & vbNewLine

		wTotalCnt = wTotalCnt + RS("�󒍐���")
		wTotalAm = wTotalAm + wPrice * RS("�󒍐���")

		RS.MoveNext

		'----- ���C��
		if RS.EOF = false then
			w_html = w_html & "  <tr bgcolor='#ffc600'>" & vbNewLine
			w_html = w_html & "    <td height='1' colspan='3'><img src='images/blank.gif' width='1' height='1'></td>"
			w_html = w_html & "  </tr>" & vbNewLine
		end if

	Loop
	w_html = w_html & "</table>" & vbNewLine
	w_html = w_html & "</td>" & vbNewLine
	w_html = w_html & "</tr>" & vbNewLine
	w_html = w_html & "</table>" & vbNewLine

end if

wListHTML = w_html

RS.Close

End Function

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

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>�J�[�g�̓��e</title>

<!--#include file="../Navi/NaviStyle.inc"-->

<script language="JavaScript">

//
// ====== �Ăяo����Window�փJ�[�g�m�F�y�[�W��\��
//
function GoToCart_onClick(){
//	parent.opener.location = "Order.asp";
	parent.location = "<%=g_HTTP%>shop/Order.asp";
}

//
// ====== �����y�[�W��\��
//
function GoToOrder_onClick(){

	parent.location = "<%=g_HTTP%>shop/LoginCheck.asp?called_from=order";
}

//
// ====== �ʏ��i�y�[�W��\��
//
function Product_onClick(pItem){
//		parent.opener.location = "ProductDetail.asp?" + pItem;
		parent.location = "<%=g_HTTP%>shop/ProductDetail.asp?" + pItem;
}

</script>

</head>

<body bgcolor="#eeeeee" leftmargin="3" topmargin="3" marginwidth="0" marginheight="0">

<table class="honbun" border="0" width="100%">
  <tr bgcolor="#ffc600">
    <td align="left" valign="middle"><b>�J�[�g�̓��e</b></td>
  </tr>
</table>

<!-- ���i���v -->
<table class="small" border="0" width="100%">
  <tr>
    <td width="100" align="left" nowrap>���v����</td>
    <td align="right" nowrap><%=wTotalCnt%>��</td>
  </tr>
  <tr>
    <td width="100" align="left" nowrap>���v���z(�ō�)</td>
    <td align="right" nowrap><%=FormatNumber(wTotalAm,0)%>�~</td>
  </tr>
</table>

<!-- ���i�ꗗ -->
<%=wListHTML%>

<!-- �����փ{�^���A�J�[�g�փ{�^�� -->
<% if wTotalCnt > 0 then%>
<table border='0' width='100%'>
  <tr>
    <td align='center'><a href='JavaScript:GoToOrder_onClick();'><img src='images/GoToOrder.gif' border='0'></a></td>
  </tr>
  <tr>
    <td align='center'><a href='JavaScript:GoToCart_onClick();'><img src='images/GoToCart.gif' border='0'></a></td>
  </tr>
</table>
<% end if%>

</body>
</html>
