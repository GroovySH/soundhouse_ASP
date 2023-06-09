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
'	�����EMovie�|�b�v�A�b�v�y�[�W
'       �����܂���Movie��URL��2�ȏ�(,��؂�)�̏ꍇ�̂ݕ\�������
'
'�X�V����
'2006/01/10 �����A���惊���N��http���܂܂�Ă���ꍇ�͊O�������N�Ƃ���B
'2008/05/07 EOF�`�F�b�N�ǉ�
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/20 an SELECT����LAC�N�G���[�Ă�K�p
'
'========================================================================

On Error Resume Next

Dim msg

Dim Connection
Dim RS

Dim ItemList()
Dim ItemCnt
Dim MakerCd
Dim ProductCd
Dim MakerName
Dim ProductName
Dim ImageFileName
Dim wShichoHTML
Dim wMovieHTML

Dim w_sql
Dim w_html
Dim w_error_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- �p�����[�^��荞��
ItemCnt = cf_unstring(ReplaceInput(Trim(Request("item"))), ItemList, "^")
MakerCd = ReplaceInput(ItemList(0))
ProductCd = ReplaceInput(ItemList(1))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "SoundMoviePopUp.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'	�����Movie�����N�쐬
'
'========================================================================
Function main()
Dim i

'---- ���i�f�[�^SELECT
w_sql = ""
w_sql = w_sql & "SELECT b.���[�J�[��"
w_sql = w_sql & "     , a.���i��"
w_sql = w_sql & "     , a.���i�摜�t�@�C����_��"
w_sql = w_sql & "     , a.�����t���O"
w_sql = w_sql & "     , a.����URL"
w_sql = w_sql & "     , a.����t���O"
w_sql = w_sql & "     , a.����URL"
w_sql = w_sql & "  FROM Web���i                a WITH (NOLOCK)"   '2012/01/20 an mod s
w_sql = w_sql & "         INNER JOIN  ���[�J�[ b WITH (NOLOCK)"
w_sql = w_sql & "           ON     b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'w_sql = w_sql & "     , ���[�J�[ b"
'w_sql = w_sql & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"     '2012/01/20 an mod e
w_sql = w_sql & "   AND a.���[�J�[�R�[�h = '" & MakerCd & "'"
w_sql = w_sql & "   AND a.���i�R�[�h = '" & ProductCd & "'"

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

'---- ����HTML�쐬
MakerName = RS("���[�J�[��")
ProductName = RS("���i��")
ImageFileName = RS("���i�摜�t�@�C����_��")

'----���������N
if RS("�����t���O") = "Y" AND RS("����URL") <> "" then
	ItemCnt = cf_unstring(RS("����URL"), ItemList, ",")
	for i=0 to itemCnt-1
		if i > 0 then
			wShichoHTML = wShichoHTML & " | "
		end if
		if InStr(LCase(ItemList(i)), "http://") > 0 then
			wShichoHTML = wShichoHTML & "<a href='" & ItemList(i) & "' class='link' target='SoundMovie'>" & i+1 & "</a>"
		else
			wShichoHTML = wShichoHTML & "<a href='" & g_HTTP & ItemList(i) & "' class='link' target='SoundMovie'>" & i+1 & "</a>"
		end if
	Next
end if

'----���惊���N
if RS("����t���O") = "Y" AND RS("����URL") <> "" then
	ItemCnt = cf_unstring(RS("����URL"), ItemList, ",")
	for i=0 to itemCnt-1
		if i > 0 then
			wMovieHTML = wMovieHTML & " | "
		end if
		if InStr(LCase(ItemList(i)), "http://") > 0 then
			wMovieHTML = wMovieHTML & "<a href='" & ItemList(i) & "' class='link' target='SoundMovie'>" & i+1 & "</a>"
		else
			wMovieHTML = wMovieHTML & "<a href='" & g_HTTP & ItemList(i) & "' class='link' target='SoundMovie'>" & i+1 & "</a>"
		end if
	Next
end if

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
<title>�����EMovie�I��</title>

<!--#include file="../Navi/NaviStyle.inc"-->

</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">

<table bgcolor="#FFFFFF" borderColor=#999999 cellSpacing=0 borderColorDark=#ffffff cellPadding=0 width=200 borderColorLight=#999999 border=1>
  <tr>
    <td width="195" height=39 bgColor=#eeeeee class="honbun">
      <b><%=MakerName%>&nbsp;<%=ProductName%></b>
    </td>
  </tr>
  <tr align=middle>
    <td height=100>
      <img height=99 src="prod_img/<%=ImageFileName%>" width=198 border=0>
    </td>
  </tr>

<% if wShichoHTML <> "" then %>
  <tr vAlign=top align=left>
    <td class=honbun>
      <table border="0" cellspacing="0" cellpadding="2" class="honbun">
        <tr>
          <td width="25">
            <img src='images/Shichou.gif' width='18' height='18' border='0' alt='��������'>
          </td>
          <td>
            <%=wShichoHTML%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
<% end if %>

<% if wMovieHTML <> "" then %>
  <tr align="left" valign="middle">
    <td height=25 noWrap class="honbun">
      <table border="0" cellspacing="0" cellpadding="2" class="honbun">
        <tr>
          <td width="25">
            <img src='images/Movie.jpg' width='18' height='18' border='0' alt='���������'>
          </td>
          <td>
            <%=wMovieHTML%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
<% end if %>

</table>

</body>
</html>
