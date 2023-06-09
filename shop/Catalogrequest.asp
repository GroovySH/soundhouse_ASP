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
'	�J�^���O�����y�[�W
'
'�X�V����
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim customer_nm
Dim furigana
Dim customer_email
Dim zip
Dim prefecture
Dim address
Dim telephone

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- �Ăяo�����v���O��������̃��b�Z�[�W���o��
msg = Session.contents("msg")
Session("msg") = ""

'---- UserID ���o��

userID = Session("userID")

if userID = "" then
	w_msg = "���O�C�����s���Ă��������"
else
	'---- Execute main
	call connect_db()
	call main()
	
	'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
	if Err.Description <> "" then
		wErrDesc = "Catalogrequest.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
	end if                                           '2011/08/01 an add e

	call close_db()
end if

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if w_msg <> "" then
	Response.Redirect "../shop/Login.asp?called_from=catalog"
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()
Dim i

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

'---- �ڋq�����o��
w_sql = ""
w_sql = w_sql & "SELECT a.�ڋq��"
w_sql = w_sql & "       , a.�ڋq�t���K�i"
w_sql = w_sql & "       , a.�ڋqE_mail1"
w_sql = w_sql & "       , b.�ڋq�X�֔ԍ�"
w_sql = w_sql & "       , b.�ڋq�s���{��"
w_sql = w_sql & "       , b.�ڋq�Z��"
w_sql = w_sql & "       , c.�ڋq�d�b�ԍ�"
w_sql = w_sql & "  FROM Web�ڋq a WITH (NOLOCK)"
w_sql = w_sql & "     , Web�ڋq�Z�� b WITH (NOLOCK)"
w_sql = w_sql & "     , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
w_sql = w_sql & " WHERE a.�ڋq�ԍ� = " & userID
w_sql = w_sql & "   AND b.�ڋq�ԍ� = a.�ڋq�ԍ�"
w_sql = w_sql & "   AND b.�Z���A�� = 1"
w_sql = w_sql & "   AND c.�ڋq�ԍ� = a.�ڋq�ԍ�"
w_sql = w_sql & "   AND c.�d�b�A�� = 1"
	  
'@@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

if RS.EOF = true then
	w_msg = "<p class='error'>�ڋq��񂪂���܂���B</p>"
	Session("msg") = w_msg
else
	customer_nm = RS("�ڋq��")
	furigana = RS("�ڋq�t���K�i")
	customer_email = RS("�ڋqE_mail1")
	zip = RS("�ڋq�X�֔ԍ�")
	prefecture = RS("�ڋq�s���{��")
	address = RS("�ڋq�Z��")
	telephone = RS("�ڋq�d�b�ԍ�")
end if

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
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>�J�^���O�����b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->

</head>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">�J�^���O����</li>
      </ul>
    </div></div></div>

    <h1 class="title">�J�^���O����</h1>
    <p>�e��l�C���i�����ڂ���Ă���wHOT MENU�x�J���[�J�^���O�𖳗��Ŕ������Ă���܂��B<br>���J�^���O�̓��[���ւł̔����ƂȂ邽�߁A���͂��܂�1�T�ԑO�ォ����ꍇ���������܂��B</p>

<form action="CatalogRequestStore.asp" method="post">    
<table class="form">
  <tr>
    <th>�����O</th>
    <td><%=customer_nm%><input type="hidden" name="customer_nm" value="<%=customer_nm%>"></td>
  </tr>
  <tr>
    <th>�t���K�i</th>
    <td><%=furigana%><input type="hidden" name="furigana" value="<%=furigana%>"></td>
  </tr>
  <tr>
    <th>���[���A�h���X</th>
    <td><%=customer_email%><input type="hidden" name="e_mail" value="<%=customer_email%>"></td>
  </tr>
  <tr>
    <th>�X�֔ԍ�</th>
    <td><%=zip%><input type="hidden" name="zip" value="<%=zip%>"></td>
  </tr>
  <tr>
    <th>�Z��</th>
    <td><%=prefecture%><%=address%><input type="hidden" name="address" value="<%=prefecture%> <%=address%>"></td>
  </tr>
  <tr>
    <th>�d�b�ԍ�</th>
    <td><%=telephone%><input type="hidden" name="telephone" value="<%=telephone%>"></td>
  </tr>
  <tr>
    <th>���̑��̃J�^���O</th>
    <td><input type="text" name="message" size="70" maxlength="100"><div>���̑�����]�̃J�^���O������ۂ́A�W�������A���f���A���[�J�[�������L�����������B�i100�����ȓ��j</div></td>
  </tr>
  <tr>
    <th>HOT MENU ��]</th>
    <td><input type="checkbox" id="i_HOTMENU" name="i_HOTMENU" value="Y" checked><label for="i_HOTMENU">��]����</label></td>
  </tr>
</table>
<p class="btnBox"><input type="submit" value="���M" class="opover"></p>
</form>


</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>