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
'	���i���r���[ �Q�l���o�^
'
'�X�V����
'2007/10/19 �n�b�J�[�Z�[�t�Ή�
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/07/30 if-web ���j���[�A�����C�A�E�g����
'
'========================================================================

On Error Resume Next

Dim userID
Dim msg

Dim ID
Dim Sankou
Dim Item

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- UserID ���o��
userID = Session("userID")

'---- �Ăяo��������̃f�[�^���o��
ID = ReplaceInput(Request("ID"))
Sankou = ReplaceInput(Request("Sankou"))
Item = ReplaceInput(Request("Item"))

if isNumeric(ID) = false then
	ID = 0
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ReviewSankou.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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
'	Function	main proc
'
'========================================================================
'
Function main()

'---- �Q�l���o�^
w_sql = ""
w_sql = w_sql & "SELECT �Q�l��"
w_sql = w_sql & "     , �s�Q�l��"
w_sql = w_sql & "  FROM ���i���r���["
w_sql = w_sql & " WHERE ID = " & ID 
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

if RS.EOF = false then

	'---- �Q�l��/�s�Q�l�� �X�V
	if Sankou = "Y" then
		RS("�Q�l��") = RS("�Q�l��") + 1
	else
		RS("�s�Q�l��") = RS("�s�Q�l��") + 1
	end if

	RS.Update
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
<html>
<head>
<meta charset="Shift_JIS">
<title>���i���r���[�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>

<!-- �R���e���cstart -->
<div id="globalContents">

  <p>�o�^����܂�����@�ǂ������肪�Ƃ��������܂����B</p>
  <p class="btnBox"><a href="ProductDetail.asp?Item=<%=item%>" class="opover">���i�y�[�W�֖߂�</a></p>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>