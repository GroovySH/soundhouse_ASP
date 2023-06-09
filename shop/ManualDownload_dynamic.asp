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
'	�}�j���A���_�E�����[�h�y�[�W
'
'�X�V����
'2006/01/10 �����N��http���܂܂�Ă���ꍇ�͊O�������N�Ƃ���B
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/25 na ���X�|���X�΍�i�t�@�C���T�C�Y�폜�j
'
'========================================================================

On Error Resume Next

Dim wMakerHTML
Dim wManualHTML

Dim Connection
Dim RS
Dim FS

Dim w_sql
Dim w_html
Dim w_error_msg
Dim wErrDesc   '2011/08/01 an add

'========================================================================

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ManualDownload.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
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

Dim vBreakKey
Dim vBreakNextKey
Dim vFile
Dim vWebPath
Dim vFilePath
Dim vFileSize
Dim i
Dim j

'---- ���[�J�[ ���o��
w_sql = ""
w_sql = w_sql & "SELECT a.���[�J�[�R�[�h"
w_sql = w_sql & "     , a.���i�R�[�h"
w_sql = w_sql & "     , a.���i��"
w_sql = w_sql & "     , a.���i�}�j���A��URL"
w_sql = w_sql & "     , b.���[�J�[��"
w_sql = w_sql & "     , b.���[�J�[�z�[���y�[�WURL"
w_sql = w_sql & "  FROM Web���i a WITH (NOLOCK)"
w_sql = w_sql & "     , ���[�J�[ b WITH (NOLOCK)"
w_sql = w_sql & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
w_sql = w_sql & "   AND a.���i�}�j���A��URL != ''"
w_sql = w_sql & " ORDER BY b.���[�J�[��"
w_sql = w_sql & "     , a.���i��"

'@@@@@@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

Set FS = CreateObject("Scripting.FileSystemObject")
vWebPath = Server.MapPath("../") & "\"

'----- �}�j���A���ꗗHTML�ҏW
vBreakNextKey = RS("���[�J�[��")

wMakerHTML = ""
wMakerHTML = wMakerHTML & "<table id='maker'>" & vbNewLine
wMakerHTML = wMakerHTML & "  <tr>" & vbNewLine

wManualHTML = ""
wManualHTML = wManualHTML & "<table id='manual'>" & vbNewLine

j = 0
Do Until RS.EOF = true
	vBreakKey = vBreakNextKey

'---- ���[�J�[
	j = j + 1

	if (j-1) mod 4 = 0 and j > 1 then
		wMakerHTML = wMakerHTML & "  </tr>" & vbNewLine
		wMakerHTML = wMakerHTML & "  <tr>" & vbNewLine
		j = 1
	end if

	wMakerHTML = wMakerHTML & "    <td>"
	wMakerHTML = wMakerHTML & "      <a href='#" & RS("���[�J�[�R�[�h") & "'>" & RS("���[�J�[��") & "</a>"
	wMakerHTML = wMakerHTML & "    </td>" & vbNewLine

'---- �}�j���A��
	if i > 0 then
		wManualHTML = wManualHTML & "  <tr><td>&nbsp;</td></tr>" & vbNewLine
	end if
	wManualHTML = wManualHTML & "  <tr>" & vbNewLine
	wManualHTML = wManualHTML & "    <td colspan='2' class='makerName'><a name='" & RS("���[�J�[�R�[�h") & "'></a>" & RS("���[�J�[��") & "</td>" & vbNewLine

	if RS("���[�J�[�z�[���y�[�WURL") <> "" then
		wManualHTML = wManualHTML & "    <td class='makerLink'>���[�J�[�T�C�g�́�<a href='" & RS("���[�J�[�z�[���y�[�WURL") & "' target='_blank'>������</a></td>" & vbNewLine
	else
		wManualHTML = wManualHTML & "    <td class='makerLink'></td>" & vbNewLine
	end if

	wManualHTML = wManualHTML & "  </tr>" & vbNewLine
	wManualHTML = wManualHTML & "  <tr>" & vbNewLine

	i = 0
	Do Until vBreakKey <> vBreakNextKey
		i = i + 1

		if (i-1) mod 3 = 0 and i > 1 then
			wManualHTML = wManualHTML & "  </tr>" & vbNewLine
			wManualHTML = wManualHTML & "  <tr>" & vbNewLine
			i = 1
		end if
		
		wManualHTML = wManualHTML & "    <td>"

'2012/01/25 na �t�@�C���T�C�Y�̎�A�\�L����߂�
'		vFileSize = ""
		if InStr(LCase(RS("���i�}�j���A��URL")), "http://") = 0 then
'			'---- �t�@�C���T�C�Y���o��
'			vFilePath = Replace(vWebPath & RS("���i�}�j���A��URL"), "/", "\")
'
'			if FS.FileExists(vFilePath) = true then
'				Set vFile = FS.GetFile(vFilePath)
'				vFileSize = vFile.size
'				if vFileSize >= 1024 then
'					vFileSize = vFileSize / 1024
'					if vFileSize >= 1024 then
'						vFileSize = vFileSize / 1024
'						vFileSize = Fix(vFileSize * 100) / 100
'						vFileSize = vFileSize & "MB"
'					else
'						vFileSize = Fix(vFileSize) & "KB"
'					end if
'				else
'					vFileSize = vFileSize & "B"
'				end if
'			end if
			wManualHTML = wManualHTML & "      <a href='" & g_HTTP & RS("���i�}�j���A��URL") & "'>" & RS("���i��") & "</a>"
		else
			wManualHTML = wManualHTML & "      <a href='" & RS("���i�}�j���A��URL") & "' target='_blank'>" & RS("���i��") & "</a>"
		end if
		wManualHTML = wManualHTML & "    </td>" & vbNewLine

		RS.MoveNext
		if RS.EOF = false then
			vBreakNextKey = RS("���[�J�[��")
		else
			vBreakNextKey = "@EOF"
		end if
	Loop
	wManualHTML = wManualHTML & "  </tr>" & vbNewLine
Loop

wMakerHTML = wMakerHTML & "  </tr>" & vbNewLine
wMakerHTML = wMakerHTML & "</table>" & vbNewLine

wManualHTML = wManualHTML & "</table>" & vbNewLine

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

<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>�}�j���A���_�E�����[�h�b�T�E���h�n�E�X</title>
<meta name="robots" content="noindex,nofollow">
<meta name="robots" content="noarchive">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" type="text/css" href="style/ManualDownload.css">
</head>
<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents" class="feedback">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">�}�j���A���_�E�����[�h</li>
      </ul>
    </div></div></div>

    <h1 class="title">�}�j���A���_�E�����[�h</h1>

<!-- ���[�J�[�ꗗ -->
<%=wMakerHTML%>

<div id="reader">
  <p>PDF�t�@�C�������������������߂ɂ�Adobe Reader���K�v�ł��B<br>�������łȂ�����<a href="http://www.adobe.co.jp/products/acrobat/readstep.html" target="_blank">������</a>����_�E�����[�h���Ă��������B</p>
  <a href="http://www.adobe.co.jp/products/acrobat/readstep.html" target="_blank"><img src="images/get_adobe_reader.png" width="158" height="39" alt="Get Adobe Reader"></a>
</div>

<!-- �}�j���A���ꗗ -->
<%=wManualHTML%>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>