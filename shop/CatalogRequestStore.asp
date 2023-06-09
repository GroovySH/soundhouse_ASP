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
'	�J�^���O����
'
'�X�V����
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT��)
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim message

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- UserID ���o��
userID = Session("userID")

'---- �Ăяo��������̃f�[�^���o��
message = ReplaceInput(Left(Request("message"), 100))

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "CatalogRequestStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

'---- �J�^���O�������o�^
w_sql = ""
w_sql = w_sql & "SELECT *"
w_sql = w_sql & "  FROM �J�^���O����"
w_sql = w_sql & " WHERE 1 = 2"
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

'---- insert �J�^���O����
RS.AddNew

RS("�J�^���O������") = now()
RS("�ڋq�ԍ�") = userID
RS("���l") = message
RS("�ŐV�t���O") = "Y"

RS.Update

RS.close

message = Replace(message, vbNewLine, "<br>")

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
    
    <p>���˗����������܂����J�^���O�𑗕t�������܂��B<br>���肪�Ƃ��������܂����B</p>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>