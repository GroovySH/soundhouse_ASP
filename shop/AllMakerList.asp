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
'	�S�J�e�S���[�ꗗ�y�[�W
'
'�X�V����
'2009/04/30 �G���[����error.asp�ֈړ�
'2009/08/04 �f�U�C���ύX�i�C���f�b�N�X�^�u��\���AStyle�ɂ��f�U�C���ɕύX�j
'           ���[�J�[�}�X�^�ɑ㗝�X�t���O��ǉ����A�㗝�X�t���O="Y"�Ȃ狭���\��
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/08/13 if-web ���j���[�A�����C�A�E�g����
'2012/10/25 ok ���������������A�啶���ŕ����o��ꍇ�ɑΉ�
'
'========================================================================

On Error Resume Next

Dim IndexCd

Dim Connection
Dim RS

Dim w_sql
Dim wMakerIndexHTML
Dim wMakerListHTML

Dim w_error_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- Get input data
IndexCd = ReplaceInput(Trim(Request("IndexCd")))
if IndexCd ="" then
	IndexCd = "^"
elseif cf_checkHankaku2(IndexCd) = false then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "AllMakerList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

'---- ���[�J�[ ���o��
w_sql = ""
w_sql = w_sql & "SELECT a.���[�J�[�R�[�h"
w_sql = w_sql & "     , a.���[�J�[��"
w_sql = w_sql & "     , a.�㗝�X�t���O"
w_sql = w_sql & "  FROM ���[�J�[ a WITH (NOLOCK)"
w_sql = w_sql & " WHERE a.Web���[�J�[�t���O = 'Y'"
w_sql = w_sql & "   AND a.���[�J�[�� LIKE '[" & IndexCd & "]%'"
w_sql = w_sql & " ORDER BY a.���[�J�[��"

'@@@@@@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

'----- ���[�J�[�ꗗHTML�ҏW

call CreatewMakerIndexHTML()
call CreatewMakerListHTML()

RS.Close

End Function

'========================================================================
'
'	Function	�C���f�b�N�X�쐬
'
'========================================================================

Function CreatewMakerIndexHTML()

wMakerIndexHTML = ""
wMakerIndexHTML = wMakerIndexHTML & "<ul id='all_maker_Initial_group'>" & vbNewLine

if IndexCd =  "abc0123456789" then
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=abc0123456789'><img src='images/AllMakerList/am_abc_on.jpg' alt=''></a></li>" & vbNewLine
else
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=abc0123456789'><img src='images/AllMakerList/am_abc_off.jpg' alt=''></a></li>" & vbNewLine
end if

if IndexCd = "defg" then
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=defg'><img src='images/AllMakerList/am_defg_on.jpg' alt=''></a></li>" & vbNewLine
else
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=defg'><img src='images/AllMakerList/am_defg_off.jpg' alt=''></a></li>" & vbNewLine
end if

if IndexCd = "hijk" then
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=hijk'><img src='images/AllMakerList/am_hijk_on.jpg' alt=''></a></li>" & vbNewLine
else
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=hijk'><img src='images/AllMakerList/am_hijk_off.jpg' alt=''></a></li>" & vbNewLine
end if

if IndexCd = "lmno" then
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=lmno'><img src='images/AllMakerList/am_lmno_on.jpg' alt=''></a></li>" & vbNewLine
else
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=lmno'><img src='images/AllMakerList/am_lmno_off.jpg' alt=''></a></li>" & vbNewLine
end if

if IndexCd = "pqrs" then
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=pqrs'><img src='images/AllMakerList/am_pqrs_on.jpg' alt=''></a></li>" & vbNewLine
else
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=pqrs'><img src='images/AllMakerList/am_pqrs_off.jpg' alt=''></a></li>" & vbNewLine
end if

if IndexCd = "tuvw" then
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=tuvw'><img src='images/AllMakerList/am_tuvw_on.jpg' alt=''></a></li>" & vbNewLine
else
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=tuvw'><img src='images/AllMakerList/am_tuvw_off.jpg' alt=''></a></li>" & vbNewLine
end if

if IndexCd = "xyz" then
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=xyz'><img src='images/AllMakerList/am_xyz_on.jpg' alt=''></a></li>" & vbNewLine
else
	wMakerIndexHTML = wMakerIndexHTML & "  <li><a href='AllMAkerList.asp?IndexCd=xyz'><img src='images/AllMakerList/am_xyz_off.jpg' alt=''></a></li>" & vbNewLine
end if

'if IndexCd = "^" then
'	wMakerIndexHTML = wMakerIndexHTML & "  <div class='all'><a href='AllMAkerList.asp'><img src='images/AllMakerList/am_all_on.jpg' width='90' height='28' border='0'></a></div>" & vbNewLine
'else
'	wMakerIndexHTML = wMakerIndexHTML & "  <div class='all'><a href='AllMAkerList.asp'><img src='images/AllMakerList/am_all_off.jpg' width='90' height='28' border='0'></a></div>" & vbNewLine
'end if

wMakerIndexHTML = wMakerIndexHTML &  "</ul>" & vbNewLine

End Function

'========================================================================
'
'	Function	���[�J�[�ꗗ
'
'========================================================================

Function CreatewMakerListHTML()

Dim vBreakKey
Dim vBreakNextKey

vBreakNextKey = Left(UCase(RS("���[�J�[��")),1)

wMakerListHTML = ""
wMakerListHTML = wMakerListHTML & "<div id='all_maker_list'>" & vbNewLine

Do Until RS.EOF = true
	vBreakKey = vBreakNextKey
	wMakerListHTML = wMakerListHTML & "  <p class='initial'>" & Left(RS("���[�J�[��"),1) & "</p>" & vbNewLine

	wMakerListHTML = wMakerListHTML & "  <ul>" & vbNewLine
	
	Do Until vBreakKey <> vBreakNextKey
		if RS("�㗝�X�t���O") = "Y" then
			wMakerListHTML = wMakerListHTML & "    <li class='distributor'><a href='SearchList.asp?i_type=m&amp;s_maker_cd=" & RS("���[�J�[�R�[�h") &  "'>" & Replace(RS("���[�J�[��"),"&","&amp;") & "</a></li>" & vbNewLine
		else
			wMakerListHTML = wMakerListHTML & "    <li><a href='SearchList.asp?i_type=m&amp;s_maker_cd=" & RS("���[�J�[�R�[�h") &  "'>" & Replace(RS("���[�J�[��"),"&","&amp;") & "</a></li>" & vbNewLine
		end if
		RS.MoveNext
		if RS.EOF = false then
			vBreakNextKey = Left(UCase(RS("���[�J�[��")),1)
		else
			vBreakNextKey = "@EOF"
		end if
	Loop
	
	wMakerListHTML = wMakerListHTML & "  </ul>" & vbNewLine

Loop
wMakerListHTML = wMakerListHTML & "</div>" & vbNewLine

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
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>���[�J�[�ꗗ�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" type="text/css" href="style/AllMakerList.css">
</head>
<body>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">���[�J�[�ꗗ</li>
      </ul>
    </div></div></div>

    <h1 class="title">���[�J�[�ꗗ</h1>

<div id="main_table">

<!-- �C���f�b�N�X -->
<%=wMakerIndexHTML%>

<!-- ���[�J�[�ꗗ -->
<%=wMakerListHTML%>

</div>
<p id="all_maker_Notes">���������̓I���W�i���u�����h�܂��͐��K�㗝�X�ƂȂ�܂��B</p>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>