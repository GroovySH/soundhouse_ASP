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
'	�v���[���g����
'
'	�X�V����
'2008/05/12 ���s�R�[�h�C���W�F�N�V�����΍�ii_to�p�����[�^�폜�j
'2008/05/13 �N���X�T�C�g���N�G�X�g�t�H�W�F���[�΍� Key�p�����[�^�`�F�b�N
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/03/02 hn SetSecureKey�̈ʒu�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim customer_nm
Dim furigana
Dim e_mail
Dim zip
Dim prefecture
Dim address
Dim telephone

DIm Skey

Dim Connection
Dim RS_customer

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
	w_msg = "<p class='error'>���O�C�����s���Ă��������</p>"
else
	call connect_db()
	call main()
	
	'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
	if Err.Description <> "" then
		wErrDesc = "PresentOubo.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
	end if                                           '2011/08/01 an add e

	call close_db()
end if

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if w_msg <> "" then
	Session("msg") = w_msg
	Response.Redirect g_HTTPS & "shop/Login.asp?called_from=present"
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
'	Function	main proc
'
'========================================================================
'
Function main()

'---- �Z�L�����e�B�[�L�[�Z�b�g 
Skey = SetSecureKey()

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
w_sql = w_sql & "   AND c.�Z���A�� = 1"
w_sql = w_sql & "   AND c.�d�b�A�� = 1"

'@@@@@@response.write(w_sql)

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open w_sql, Connection, adOpenStatic

if RS_customer.EOF = true then
	w_msg = "<p class='error'>�ڋq��񂪂���܂���B</p>"
else
	customer_nm = RS_customer("�ڋq��")
	furigana = RS_customer("�ڋq�t���K�i")
	e_mail = RS_customer("�ڋqE_mail1")
	zip = RS_customer("�ڋq�X�֔ԍ�")
	prefecture = RS_customer("�ڋq�s���{��")
	address = RS_customer("�ڋq�Z��")
	telephone = RS_customer("�ڋq�d�b�ԍ�")
end if

RS_customer.close

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
<title>�v���[���g����b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->

<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">�v���[���g����</li>
      </ul>
    </div></div></div>

    <h1 class="title">�v���[���g����</h1>
    
<form name="f_data" action="PresentOuboSend.asp" method="post">
    
<table class="form">
  <tr>
    <th>�����O</th>
    <td><%=customer_nm%></td>
  </tr>
  <tr>
    <th>�t���K�i</th>
    <td><%=furigana%></td>
  </tr>
  <tr>
    <th>�X�֔ԍ�</th>
    <td><%=zip%></td>
  </tr>
  <tr>
    <th>�Z��</th>
    <td><%=prefecture%><%=address%></td>
  </tr>
  <tr>
    <th>���[���A�h���X</th>
    <td><%=e_mail%></td>
  </tr>
  <tr>
    <th>�T�E���h�n�E�X�w����</th>
    <td>
    	<input type="radio" id="0" name="purchase" value="���߂�"><label for="0">���߂�</label>
		<input type="radio" id="1_2" name="purchase" value="1�`2��"><label for="1_2">1�`2��</label>
		<input type="radio" id="3_9" name="purchase" value="3�`9��"><label for="3_9">3�`9��</label>
        <input type="radio" id="10" name="purchase" value="10��ȏ�"><label for="10">10��ȏ�</label>
    </td>
  </tr>
  <tr>
    <th>�R�����g</th>
    <td><textarea name="comment" cols="70" rows="5"></textarea></td>
  </tr>
</table>

<p class="btnBox"><input type="submit" value="���M" class="opover"></p>

<input type="hidden" name="customer_nm" value="<%=customer_nm%>">
<input type="hidden" name="furigana" value="<%=furigana%>">
<input type="hidden" name="zip" value="<%=zip%>">
<input type="hidden" name="prefecture" value="<%=prefecture%>">
<input type="hidden" name="address" value="<%=address%>">
<input type="hidden" name="telephone" value="<%=telephone%>">
<input type="hidden" name="e_mail" value="<%=e_mail%>">
<input type="hidden" name="Skey" value="<%=Skey%>">
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