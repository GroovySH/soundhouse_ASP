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
'	���i���r���[�����e�i���X�y�[�W
'     ���i���r���[�̕ύX/�폜���s��
'
'2011/09/06 an #816 �V�K�쐬
'2012/08/11 nt �V���b�v�R�����g�̓��̓t�H�[���\������т��̐����ǉ�
'
'========================================================================
On Error Resume Next
Response.buffer = true
Response.Expires = -1			' Do not cache

Dim ReviewID
Dim i_Mode
Dim Title
Dim Hyouka
Dim Review
Dim UserID		'2012/08/11 nt add
Dim Password	'2012/08/11 nt add
Dim Auth		'2012/08/11 nt add
Dim sCDate		'2012/08/11 nt add
Dim sComment	'2012/08/11 nt add
Dim vErrMSG		'2012/08/11 nt add

Dim wReviewDate
Dim wReviewName
Dim wMakerName
Dim wProductName

Dim Skey
Dim Connection

Dim wMSG   'ReviewMaint2����̃G���[/�������b�Z�[�W
Dim wNoData
Dim wLoginFl

'========================================================================

'---- Get GET/POST data
ReviewID = ReplaceInput(Request("ReviewID"))
i_Mode = ReplaceInput(Request("i_Mode"))          '�ȉ��A�G���[����ReviewMaint2����󂯎��
Title = ReplaceInput(Left(Request("Title"),50))
Hyouka = ReplaceInput(Request("Hyouka"))
Review = ReplaceInput(Left(Request("Review"),1000))
UserID = ReplaceInput(Request.Cookies("UserID"))			'2012/08/11 nt add
Password = ReplaceInput(Request.Cookies("Password"))		'2012/08/11 nt add
sCDate = ReplaceInput(Request("sCDate"))					'2012/08/11 nt add
sComment = ReplaceInput(Left(Request("sComment"),1000))		'2012/08/11 nt add

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

'---- ID�̎w�肪�s���A�Y�����r���[���Ȃ��ꍇ�͌�����ʂ�
if wNoData = "Y" then
	Response.Redirect "ReviewSearch.asp?UserID=" & UserID
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
Function main()

wMSG = ""
wNoData = "N"
wLoginFl = "N"

'---- �Z�L�����e�B�[�L�[�Z�b�g
Skey = SetSecureKey()

'---- ���O�C���X�e�[�^�X�擾
wLoginFl = fGetSessionData(gSessionID, "ShAdminFl")

if wLoginFl <> "Y" then
	call fSetSessionData(gSessionID, "���b�Z�[�W", "���O�C�����Ă��������B")
	exit function
end if

'2012/08/13 nt add
'---- ���O�C�����[�U�̌������擾
call getWEBMasterAuth

'---- ReviewMaint2.asp�̃G���[���b�Z�[�W�擾�E�N���A
wMSG = fGetSessionData(gSessionID, "���b�Z�[�W")
call fSetSessionData(gSessionID, "���b�Z�[�W", "")

'---- ���̓`�F�b�N
call validation()

if wNoData <> "Y" then
	call GetReview()
end if

End function

'========================================================================
'
'    Function    ���͓��e�`�F�b�N
'
'========================================================================
'
Function validation()

Dim vErrMSG

vErrMSG = ""

if ReviewID = "" then
	vErrMSG = "���r���[ID����͂��Ă��������B"
else
	if cf_checkNumeric(ReviewID) = false then
		vErrMSG = "���r���[ID���s���ł��B"
	end if
end if

if vErrMSG <> "" then
	wNoData = "Y"
	call fSetSessionData(gSessionID, "���b�Z�[�W", vErrMSG)
end if

End function

'========================================================================
'
'    Function    ���i���r���[�擾
'
'========================================================================
'
Function GetReview()

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    a.ID"
vSQL = vSQL & "  , a.���e��"
vSQL = vSQL & "  , a.�^�C�g��"
vSQL = vSQL & "  , a.�]��"
vSQL = vSQL & "  , a.���O"
vSQL = vSQL & "  , a.���r���[���e"
vSQL = vSQL & "  , b.���[�J�[��"
vSQL = vSQL & "  , c.���i��"
'2012/08/11 nt add Start
vSQL = vSQL & "  , a.�V���b�v�R�����g��"
vSQL = vSQL & "  , a.�V���b�v�R�����g�^�C�g��"
vSQL = vSQL & "  , a.�V���b�v�R�����g"
'2012/08/11 nt add End
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ���i���r���[ a WITH (NOLOCK)"
vSQL = vSQL & "  , ���[�J�[ b WITH (NOLOCK)"
vSQL = vSQL & "  , Web���i c WITH (NOLOCK)"
vSQL = vSQL & " WHERE a.ID = " & ReviewID
vSQL = vSQL & "   AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "   AND c.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"

'@@@@@response.write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

if RSv.EOF = true then
	wNoData = "Y"
	vErrMSG = "�Y���̃��r���[������܂���B ���r���[ID��" & ReviewID
	call fSetSessionData(gSessionID, "���b�Z�[�W", vErrMSG)
else
	wReviewDate = RSv("���e��")
	wReviewName = RSv("���O")
	wMakerName = RSv("���[�J�[��")
	wProductName = RSv("���i��")

	'---- ReviewMaint2����G���[�Ŗ߂������́ADB����擾���Ȃ�
	if i_Mode <> "update" then

		Title = RSv("�^�C�g��")
		Hyouka = RSv("�]��")
		Review = RSv("���r���[���e")

		'2012/08/11 nt add Start
		sCDate = RSv("�V���b�v�R�����g��")
		if (isNull(sCDate) = true) then
			'---- �V���b�v�R�����g���t���Ȃ��ꍇ�A�V�X�e�����t���Z�b�g
			sCDate = now()
		end if

		sComment = RSv("�V���b�v�R�����g")
		'2012/08/11 nt add End
	end if
end if

RSv.Close

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

'2012/08/11 nt add
'========================================================================
'
'	Function	���O�C�����[�U��WEB�Ǘ��Ҍ������擾
'
'========================================================================
Function getWEBMasterAuth()

Dim RSv
Dim vSQL

vSQL = ""
vSQL = vSQL & "SELECT ���� "
vSQL = vSQL & " FROM "
vSQL = vSQL & "    WEB�Ǘ��� a WITH (NOLOCK) "
vSQL = vSQL & " WHERE "
vSQL = vSQL & "        a.���[�UID = '" & UserID & "' "
vSQL = vSQL & "    AND a.�p�X���[�h = '" & Password & "' "
vSQL = vSQL & "    AND a.�폜�t���O = '0'"	'�폜�t���O[0]�FActive�A[1]�FNon-Active

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	wNoData = "Y"
	vErrMSG = "�Y���̃��[�U�����݂��܂���B ���[�UID��" & UserID
	call fSetSessionData(gSessionID, "���b�Z�[�W", vErrMSG)
else
	'---- WEB�Ǘ��ҏ��̗L�����擾
	Auth = RSv("����")
end if

RSv.Close

End Function

'========================================================================

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS" />
<title>���i���r���[�����e�i���X</title>
<link rel="stylesheet" type="text/css" href="style/review.css" />
<script type="text/javascript" src="jslib/review.js?20120817"></script>
</head>
<body>
<div id="content">
<h1>���i���r���[�����e�i���X</h1>
<td>��<font color="red"><%=UserID%> </font>����Ń��O�C�����ł��B</td>
<% if wMSG <> "" then %>
<p class="notes"><%=wMSG%></p>
<% end if %>
<form name="f_data" method="post">
<table>
  <tr>
    <th>���r���[ID</th>
    <td><%=ReviewID%></td>
  </tr>
  <tr>
    <th>���[�J�[</th>
    <td><%=wMakerName%></td>
  </tr>
  <tr>
    <th>���i��</th>
    <td><%=wProductName%></td>
  </tr>
  <tr>
    <th>���e��</th>
    <td><%=wReviewDate%></td>
  </tr>
  <tr>
    <th>�^�C�g��</th>
    <td><input type="text" name="Title" value="<%=Title%>" size="50" maxsize="50" <%if (Auth <> "1") And (Auth <> "2") then%>readonly="readonly" style="background-color:#DCDCDC;"<%end if%> /></td>
  </tr>
  <tr>
    <th>�]��</th>
    <td><input type="text" name="Hyouka" value="<%=Hyouka%>" <%if (Auth <> "1") And (Auth <> "2") then%>readonly="readonly" style="background-color:#DCDCDC;"<%end if%> /></td>
  </tr>
  <tr>
    <th>���e�Җ�</th>
    <td><%=wReviewName%></td>
  </tr>
  <tr>
    <th>���r���[���e</th>
    <td><textarea name="Review" rows="15" cols="60" <%if (Auth <> "1") And (Auth <> "2") then%>readonly="readonly" style="background-color:#DCDCDC;"<%end if%> ><%=Review%></textarea></td>
  </tr>
</table>
<hr>
<h2>�V���b�v�R�����g</h2>
<table>
  <tr>
    <th>�R�����g��</th>
    <td><input type="text" name="sCDate" value="<%=sCDate%>" <%if (Auth <> "1") And (Auth <> "3") then%>readonly="readonly" style="background-color:#DCDCDC;"<%end if%> /></td>
  </tr>
  <tr>
    <th>�R�����g</th>
    <td><textarea name="sComment" rows="15" cols="60" <%if (Auth <> "1") And (Auth <> "3") then%>readonly="readonly" style="background-color:#DCDCDC;" <%end if%> ><%=sComment%></textarea></td>
  </tr>
</table>
<div id="button_div">

<!-- 2012/08/11 nt mod Start -->
<input type="submit" value=" �ύX " onClick="return Update_onClick();" <%if (Auth <> "1") And (Auth <> "2") And (Auth <> "3") then%>disabled<%end if%> />
<input type="submit" value=" �폜 " onClick="return Delete_onClick();" <%if (Auth <> "1") then%>disabled<%end if%> />
<input type="submit" value=" �V���b�v�R�����g�̂ݍ폜 " onClick="return sCDelete_onClick();" <%if (Auth <> "1") And (Auth <> "3") then%>disabled<%end if%> />
<!-- 2012/08/11 nt mod End -->

<input type="submit" value=" �߂� " onClick="Return_onClick();" />
<input type="hidden" name="ReviewID" value="<%=ReviewID%>" />
<!-- 2012/08/11 nt add Start -->
<input type="hidden" name="UserID" value="<%=UserID%>" />
<!-- 2012/08/11 nt add End -->
<input type="hidden" name="i_Mode" value="" />
<input type="hidden" name="Skey" value="<%=Skey%>" />
</div>
</form>
</div>
</body>
</html>