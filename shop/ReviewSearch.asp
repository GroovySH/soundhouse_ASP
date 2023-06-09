<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
'�T�E���h�n�E�X
Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/Bfunctions1.asp"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/HttpsSecurity.inc"-->

<%
'========================================================================
'
'    ���i���r���[����
'�X�V����
'2011/09/06 an #816 �V�K�쐬
'2012/08/11 nt ���O�C�����̈����p����ǉ�
'
'========================================================================

On Error Resume Next
Response.buffer = true
Response.Expires = -1			' Do not cache

Dim Connection

Dim UserID		'2012/08/11 nt add
Dim Password	'2012/08/11 nt add
Dim wErrMSG
Dim wLoginFl

'========================================================================

'2012/08/11 nt add
'---- Get Cookie data
UserID = Request.Cookies("UserID")
Password = Request.Cookies("Password")

'---- Execute main
call connect_db()
call main()
call close_db()

if Err.Description <> "" then
    Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- �����O�C���̏ꍇ�̓��O�C����ʂ�
if wLoginFl <> "Y" then
	Response.Redirect "ReviewMaintLogin.asp"
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
'	Function	main
'
'========================================================================
'
Function main()

wErrMSG = ""
wLoginFl = "N"

'---- ���O�C���X�e�[�^�X�擾
wLoginFl = fGetSessionData(gSessionID, "ShAdminFl")

if wLoginFl <> "Y" then
	call fSetSessionData(gSessionID, "���b�Z�[�W", "���O�C�����Ă��������B")
	exit function
end if

'---- ReviewMaint�̃G���[���b�Z�[�W�AReviewMain2�̍폜�������b�Z�[�W�擾�E�N���A
wErrMSG = fGetSessionData(gSessionID, "���b�Z�[�W")
call fSetSessionData(gSessionID, "���b�Z�[�W", "")

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing

End function

'========================================================================
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS" />
<title>���i���r���[����</title>
<link rel="stylesheet" type="text/css" href="style/review.css" />
</head>
<body>
<div id="content">
<h1>���i���r���[����</h1>
<td>��<font color="red"><%=UserID%> </font>����Ń��O�C�����ł��B</td>
<% if wErrMSG <> "" then %>
<p class="notes"><%=wErrMSG%></p>
<% end if %>
<p>�ύX�E�폜���s�����r���[ID����͂��Ă��������B</p>
<form name="f_data" method="post" action="ReviewMaint.asp">
<ul>
<li>���r���[ID�@<input name="ReviewID" maxlength="30" size="30" autocomplete="off" /></li>
<ul>
<br />
<span style="margin:45px"><input type="submit" value="���r���[����" /></span>
<!-- 2012/08/11 nt add Start -->
<input type="hidden" name="UserID" value="<% = UserID %>">
<!-- 2012/08/11 nt add End -->
</form>
<br /><br />
<form method="post" action="ReviewMaintLoginCheck.asp?Logout=Y">
<span style="margin:45px"><input type="submit" value="���O�A�E�g" /></span>
</form>
</div>
</body>
</html>