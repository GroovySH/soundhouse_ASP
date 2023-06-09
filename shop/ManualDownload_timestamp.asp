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

Dim wLogMsg		'2012/01/17 yo add
Dim wTimeStr	'2012/01/17 yo add


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

'2012/01/17 yo add��
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "DB�����J�n(" & wTimeStr & ")"
'2012/01/17 yo add��

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

'2012/01/17 yo add��
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "DB�����I��(" & wTimeStr & ")"
'2012/01/17 yo add��

'2012/01/17 yo add��
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "�}�j���A���ꗗ�쐬�J�n(" & wTimeStr & ")"
'2012/01/17 yo add��

Set FS = CreateObject("Scripting.FileSystemObject")
vWebPath = Server.MapPath("../") & "\"

'----- �}�j���A���ꗗHTML�ҏW
vBreakNextKey = RS("���[�J�[��")

wMakerHTML = ""
wMakerHTML = wMakerHTML & "<table bgcolor='eeeeee' width='790' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine
wMakerHTML = wMakerHTML & "  <tr>" & vbNewLine

wManualHTML = ""
wManualHTML = wManualHTML & "<table width='790' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine

j = 0
Do Until RS.EOF = true
	vBreakKey = vBreakNextKey

'---- ���[�J�[
	j = j + 1

	if (j-1) mod 4 = 0 and j > 1 then
		wMakerHTML = wMakerHTML & "  </tr>" & vbNewLine
		wMakerHTML = wMakerHTML & "  <tr bgcolor='eeeeee'>" & vbNewLine
		j = 1
	end if

	wMakerHTML = wMakerHTML & "    <td width='195'>" & vbNewLine
	wMakerHTML = wMakerHTML & "      <img src='images/Sankaku.gif' width='9' height='11'><a href='#" & RS("���[�J�[�R�[�h") & "' class='link'>" & RS("���[�J�[��") & "</a>" & vbNewLine
	wMakerHTML = wMakerHTML & "    </td>" & vbNewLine

'---- �}�j���A��
	if i > 0 then
		wManualHTML = wManualHTML & "  <tr><td>&nbsp;</td></tr>" & vbNewLine
	end if
	wManualHTML = wManualHTML & "  <tr>" & vbNewLine
	wManualHTML = wManualHTML & "    <td colspan='2' height='20' bgcolor='eeeeee' class='honbun'><b><a name='" & RS("���[�J�[�R�[�h") & "'></a>" & RS("���[�J�[��") & "</b></td>" & vbNewLine

	if RS("���[�J�[�z�[���y�[�WURL") <> "" then
		wManualHTML = wManualHTML & "    <td align='right' bgcolor='eeeeee'><font size='-1'><span class='honbun'>���[�J�[�T�C�g�́�</span><a href='" & RS("���[�J�[�z�[���y�[�WURL") & "' target='_blank' class='link'>������</a></font></td>" & vbNewLine
	else
		wManualHTML = wManualHTML & "    <td align='right' bgcolor='eeeeee'></td>" & vbNewLine
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
		
		wManualHTML = wManualHTML & "    <td width='260'>" & vbNewLine

		vFileSize = ""
		if InStr(LCase(RS("���i�}�j���A��URL")), "http://") = 0 then
			'---- �t�@�C���T�C�Y���o��
			vFilePath = Replace(vWebPath & RS("���i�}�j���A��URL"), "/", "\")

			if FS.FileExists(vFilePath) = true then
				Set vFile = FS.GetFile(vFilePath)
				vFileSize = vFile.size
				if vFileSize >= 1024 then
					vFileSize = vFileSize / 1024
					if vFileSize >= 1024 then
						vFileSize = vFileSize / 1024
						vFileSize = Fix(vFileSize * 100) / 100
						vFileSize = vFileSize & "MB"
					else
						vFileSize = Fix(vFileSize) & "KB"
					end if
				else
					vFileSize = vFileSize & "B"
				end if
			end if
			wManualHTML = wManualHTML & "      <a href='" & g_HTTP & RS("���i�}�j���A��URL") & "' class='link'>" & RS("���i��") & "</a> <span class='honbun'>(" & vFileSize & ")</span>" & vbNewLine
		else
			wManualHTML = wManualHTML & "      <a href='" & RS("���i�}�j���A��URL") & "' class='link' target='_blank'>" & RS("���i��") & "</a>" & vbNewLine
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

'2012/01/17 yo add��
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "�}�j���A���ꗗ�쐬�I��(" & wTimeStr & ")"
Call writeLog(wLogMsg)
'2012/01/17 yo add��

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

'2012/01/17 yo �b��ŏ������Ԃ������o����
Function getTimeStr(ByRef nowTimeStr)

dim nowTime, nowHour, nowMinute, nowSecond, nowMilliS

nowTime = Timer()
nowMilliS = nowTime - Fix(nowTime)
nowMilliS = Right("000" & Fix(nowMilliS * 1000), 3)
nowTime = Fix(nowTime)
nowSecond = Right("00" & (nowTime Mod 60), 2)
nowTime = Fix(nowTime / 60)
nowMinute = Right("00" & (nowTime Mod 60), 2)
nowTime = Fix(nowTime / 60)
nowHour = Right("00" & nowTime, 2)
nowTimeStr = nowHour & ":" & nowMinute & ":" & nowSecond & "." & nowMilliS

End Function

Function writeLog(inOption)

dim fso, f
dim logFileName
'---- Log File open
Set fso = CreateObject("Scripting.FileSystemObject")
logFileName = "../ErrorLog/ManualLog" & Year(Date()) & Right("0" & Month(Date()), 2) & Right("0" & Day(Date()), 2) & ".log"
logFileName = Server.MapPath(logFileName)               'Map log file

Set f = fso.OpenTextFile(logFileName, 8, true)       'Log open - Append Mode
'---- Logging
f.WriteLine(inOption )
f.Close

End function
'2012/01/17 yo �b��ŏ������Ԃ������o����

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>�T�E���h�n�E�X �}�j���A���_�E�����[�h</title>

<!-- �ǉ�SCRIPT�͂�����-->

<!--#include file="../Navi/NaviStyle.inc"-->

</head>

<body background="../Navi/Images/back_ground.gif" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!--#include file="../Navi/NaviTop.inc"-->

<!--'2012/01/17 yo add��-->
<%
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "HTML�o�͊J�n(" & wTimeStr & ")"
%>
<!--'2012/01/17 yo add��-->

<table width="940" height="26" border="0" cellpadding="0" cellspacing="0">
  <tr>

<!--#include file="../Navi/NaviLeft.inc"-->

    <td width="798" align="left" valign="top" bgcolor="#ffffff">


<!------------ �y�[�W���C�������̋L�q START ------------>

      <table border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td align="left"><b><font color="#696684">�}�j���A���_�E�����[�h</font></b></td>
        </tr>
      </table>

      <table width="798" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;</td>
          <td>

<!-- ���[�J�[�ꗗ -->
<%=wMakerHTML%>

            <table width="100%" border="0" cellspacing="0" cellpadding="4">
              <tr> 
                <td><font size="-1"><span class='honbun'>PDF�t�@�C�������������������߂ɂ�Acrobat Reader���K�v�ł��B�������łȂ�����Adobe���疳���Ŕz�z����Ă��܂��̂ŁA</span><a href="http://www.adobe.co.jp/products/acrobat/readstep.html" target="_blank" class='link'>��������</a><span class='honbun'>�C���X�g�[�����_�E�����[�h���Ă��������BAcrobat Reader�̓��[�U�o�^�����邾���Ŏ��R�Ɏg�����Ƃ��ł��܂��B</span></font></td>
                <td width="90"><font size="-1"><a href="http://www.adobe.co.jp/products/acrobat/readstep.html" target="_blank"><img src="../images/getacro.gif" width="88" height="31" border="0" alt="Get Acrobat Reader"></a></font></td>
              </tr>
              <tr>
                <td colspan="2" height="4">
                  <hr width="100%" align="center" noshade size="1">
                </td>
              </tr>
            </table>

<!-- �}�j���A���ꗗ -->
<%=wManualHTML%>

          </td>
        </tr>
      </table>

<!------------ �y�[�W���C�������̋L�q END ------------>

    </td>
  </tr>
</table>

<!--#include file="../Navi/NaviBottom.inc"-->

<!--#include file="../Navi/NaviScript.inc"-->

<!--'2012/01/17 yo add��-->
<%
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "HTML�o�͏I��(" & wTimeStr & ")"
Call writeTime(wLogMsg)
%>
<!--'2012/01/17 yo add��-->

</body>
</html>
