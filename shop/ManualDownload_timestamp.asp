<%@ LANGUAGE="VBScript" %>
<%
'ネットハウスねっとハウスネットはうす
'サウンドハウス
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->

<%
'========================================================================
'
'	マニュアルダウンロードページ
'
'更新履歴
'2006/01/10 リンクにhttpが含まれている場合は外部リンクとする。
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
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

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
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

'2012/01/17 yo add↓
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "DB検索開始(" & wTimeStr & ")"
'2012/01/17 yo add↑

'---- メーカー 取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.メーカーコード"
w_sql = w_sql & "     , a.商品コード"
w_sql = w_sql & "     , a.商品名"
w_sql = w_sql & "     , a.製品マニュアルURL"
w_sql = w_sql & "     , b.メーカー名"
w_sql = w_sql & "     , b.メーカーホームページURL"
w_sql = w_sql & "  FROM Web商品 a WITH (NOLOCK)"
w_sql = w_sql & "     , メーカー b WITH (NOLOCK)"
w_sql = w_sql & " WHERE b.メーカーコード = a.メーカーコード"
w_sql = w_sql & "   AND a.製品マニュアルURL != ''"
w_sql = w_sql & " ORDER BY b.メーカー名"
w_sql = w_sql & "     , a.商品名"

'@@@@@@@@@@response.write(w_sql)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

'2012/01/17 yo add↓
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "DB検索終了(" & wTimeStr & ")"
'2012/01/17 yo add↑

'2012/01/17 yo add↓
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "マニュアル一覧作成開始(" & wTimeStr & ")"
'2012/01/17 yo add↑

Set FS = CreateObject("Scripting.FileSystemObject")
vWebPath = Server.MapPath("../") & "\"

'----- マニュアル一覧HTML編集
vBreakNextKey = RS("メーカー名")

wMakerHTML = ""
wMakerHTML = wMakerHTML & "<table bgcolor='eeeeee' width='790' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine
wMakerHTML = wMakerHTML & "  <tr>" & vbNewLine

wManualHTML = ""
wManualHTML = wManualHTML & "<table width='790' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine

j = 0
Do Until RS.EOF = true
	vBreakKey = vBreakNextKey

'---- メーカー
	j = j + 1

	if (j-1) mod 4 = 0 and j > 1 then
		wMakerHTML = wMakerHTML & "  </tr>" & vbNewLine
		wMakerHTML = wMakerHTML & "  <tr bgcolor='eeeeee'>" & vbNewLine
		j = 1
	end if

	wMakerHTML = wMakerHTML & "    <td width='195'>" & vbNewLine
	wMakerHTML = wMakerHTML & "      <img src='images/Sankaku.gif' width='9' height='11'><a href='#" & RS("メーカーコード") & "' class='link'>" & RS("メーカー名") & "</a>" & vbNewLine
	wMakerHTML = wMakerHTML & "    </td>" & vbNewLine

'---- マニュアル
	if i > 0 then
		wManualHTML = wManualHTML & "  <tr><td>&nbsp;</td></tr>" & vbNewLine
	end if
	wManualHTML = wManualHTML & "  <tr>" & vbNewLine
	wManualHTML = wManualHTML & "    <td colspan='2' height='20' bgcolor='eeeeee' class='honbun'><b><a name='" & RS("メーカーコード") & "'></a>" & RS("メーカー名") & "</b></td>" & vbNewLine

	if RS("メーカーホームページURL") <> "" then
		wManualHTML = wManualHTML & "    <td align='right' bgcolor='eeeeee'><font size='-1'><span class='honbun'>メーカーサイトは→</span><a href='" & RS("メーカーホームページURL") & "' target='_blank' class='link'>こちら</a></font></td>" & vbNewLine
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
		if InStr(LCase(RS("製品マニュアルURL")), "http://") = 0 then
			'---- ファイルサイズ取り出し
			vFilePath = Replace(vWebPath & RS("製品マニュアルURL"), "/", "\")

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
			wManualHTML = wManualHTML & "      <a href='" & g_HTTP & RS("製品マニュアルURL") & "' class='link'>" & RS("商品名") & "</a> <span class='honbun'>(" & vFileSize & ")</span>" & vbNewLine
		else
			wManualHTML = wManualHTML & "      <a href='" & RS("製品マニュアルURL") & "' class='link' target='_blank'>" & RS("商品名") & "</a>" & vbNewLine
		end if
		wManualHTML = wManualHTML & "    </td>" & vbNewLine

		RS.MoveNext
		if RS.EOF = false then
			vBreakNextKey = RS("メーカー名")
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

'2012/01/17 yo add↓
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "マニュアル一覧作成終了(" & wTimeStr & ")"
Call writeLog(wLogMsg)
'2012/01/17 yo add↑

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

'2012/01/17 yo 暫定で処理時間を書き出す↓
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
'2012/01/17 yo 暫定で処理時間を書き出す↑

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<title>サウンドハウス マニュアルダウンロード</title>

<!-- 追加SCRIPTはここへ-->

<!--#include file="../Navi/NaviStyle.inc"-->

</head>

<body background="../Navi/Images/back_ground.gif" bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!--#include file="../Navi/NaviTop.inc"-->

<!--'2012/01/17 yo add↓-->
<%
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "HTML出力開始(" & wTimeStr & ")"
%>
<!--'2012/01/17 yo add↑-->

<table width="940" height="26" border="0" cellpadding="0" cellspacing="0">
  <tr>

<!--#include file="../Navi/NaviLeft.inc"-->

    <td width="798" align="left" valign="top" bgcolor="#ffffff">


<!------------ ページメイン部分の記述 START ------------>

      <table border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td align="left"><b><font color="#696684">マニュアルダウンロード</font></b></td>
        </tr>
      </table>

      <table width="798" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;</td>
          <td>

<!-- メーカー一覧 -->
<%=wMakerHTML%>

            <table width="100%" border="0" cellspacing="0" cellpadding="4">
              <tr> 
                <td><font size="-1"><span class='honbun'>PDFファイルをご覧いただくためにはAcrobat Readerが必要です。お持ちでない方はAdobeから無料で配布されていますので、</span><a href="http://www.adobe.co.jp/products/acrobat/readstep.html" target="_blank" class='link'>ここから</a><span class='honbun'>インストーラをダウンロードしてください。Acrobat Readerはユーザ登録をするだけで自由に使うことができます。</span></font></td>
                <td width="90"><font size="-1"><a href="http://www.adobe.co.jp/products/acrobat/readstep.html" target="_blank"><img src="../images/getacro.gif" width="88" height="31" border="0" alt="Get Acrobat Reader"></a></font></td>
              </tr>
              <tr>
                <td colspan="2" height="4">
                  <hr width="100%" align="center" noshade size="1">
                </td>
              </tr>
            </table>

<!-- マニュアル一覧 -->
<%=wManualHTML%>

          </td>
        </tr>
      </table>

<!------------ ページメイン部分の記述 END ------------>

    </td>
  </tr>
</table>

<!--#include file="../Navi/NaviBottom.inc"-->

<!--#include file="../Navi/NaviScript.inc"-->

<!--'2012/01/17 yo add↓-->
<%
Call getTimeStr(wTimeStr)
wLogMsg = wLogMsg & "HTML出力終了(" & wTimeStr & ")"
Call writeTime(wLogMsg)
%>
<!--'2012/01/17 yo add↑-->

</body>
</html>
