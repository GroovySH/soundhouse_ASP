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
'2012/01/25 na レスポンス対策（ファイルサイズ削除）
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

Set FS = CreateObject("Scripting.FileSystemObject")
vWebPath = Server.MapPath("../") & "\"

'----- マニュアル一覧HTML編集
vBreakNextKey = RS("メーカー名")

wMakerHTML = ""
wMakerHTML = wMakerHTML & "<table id='maker'>" & vbNewLine
wMakerHTML = wMakerHTML & "  <tr>" & vbNewLine

wManualHTML = ""
wManualHTML = wManualHTML & "<table id='manual'>" & vbNewLine

j = 0
Do Until RS.EOF = true
	vBreakKey = vBreakNextKey

'---- メーカー
	j = j + 1

	if (j-1) mod 4 = 0 and j > 1 then
		wMakerHTML = wMakerHTML & "  </tr>" & vbNewLine
		wMakerHTML = wMakerHTML & "  <tr>" & vbNewLine
		j = 1
	end if

	wMakerHTML = wMakerHTML & "    <td>"
	wMakerHTML = wMakerHTML & "      <a href='#" & RS("メーカーコード") & "'>" & RS("メーカー名") & "</a>"
	wMakerHTML = wMakerHTML & "    </td>" & vbNewLine

'---- マニュアル
	if i > 0 then
		wManualHTML = wManualHTML & "  <tr><td>&nbsp;</td></tr>" & vbNewLine
	end if
	wManualHTML = wManualHTML & "  <tr>" & vbNewLine
	wManualHTML = wManualHTML & "    <td colspan='2' class='makerName'><a name='" & RS("メーカーコード") & "'></a>" & RS("メーカー名") & "</td>" & vbNewLine

	if RS("メーカーホームページURL") <> "" then
		wManualHTML = wManualHTML & "    <td class='makerLink'>メーカーサイトは→<a href='" & RS("メーカーホームページURL") & "' target='_blank'>こちら</a></td>" & vbNewLine
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

'2012/01/25 na ファイルサイズ採取、表記をやめる
'		vFileSize = ""
		if InStr(LCase(RS("製品マニュアルURL")), "http://") = 0 then
'			'---- ファイルサイズ取り出し
'			vFilePath = Replace(vWebPath & RS("製品マニュアルURL"), "/", "\")
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
			wManualHTML = wManualHTML & "      <a href='" & g_HTTP & RS("製品マニュアルURL") & "'>" & RS("商品名") & "</a>"
		else
			wManualHTML = wManualHTML & "      <a href='" & RS("製品マニュアルURL") & "' target='_blank'>" & RS("商品名") & "</a>"
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
<title>マニュアルダウンロード｜サウンドハウス</title>
<meta name="robots" content="noindex,nofollow">
<meta name="robots" content="noarchive">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" type="text/css" href="style/ManualDownload.css">
</head>
<body>

<!--#include file="../Navi/NaviTop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents" class="feedback">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li class="now">マニュアルダウンロード</li>
      </ul>
    </div></div></div>

    <h1 class="title">マニュアルダウンロード</h1>

<!-- メーカー一覧 -->
<%=wMakerHTML%>

<div id="reader">
  <p>PDFファイルをご覧いただくためにはAdobe Readerが必要です。<br>お持ちでない方は<a href="http://www.adobe.co.jp/products/acrobat/readstep.html" target="_blank">こちら</a>からダウンロードしてください。</p>
  <a href="http://www.adobe.co.jp/products/acrobat/readstep.html" target="_blank"><img src="images/get_adobe_reader.png" width="158" height="39" alt="Get Adobe Reader"></a>
</div>

<!-- マニュアル一覧 -->
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