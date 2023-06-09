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
'	試聴・Movieポップアップページ
'       試聴またがMovieのURLが2個以上(,区切り)の場合のみ表示される
'
'更新履歴
'2006/01/10 試聴、動画リンクにhttpが含まれている場合は外部リンクとする。
'2008/05/07 EOFチェック追加
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/20 an SELECT文へLACクエリー案を適用
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

'---- パラメータ取り込み
ItemCnt = cf_unstring(ReplaceInput(Trim(Request("item"))), ItemList, "^")
MakerCd = ReplaceInput(ItemList(0))
ProductCd = ReplaceInput(ItemList(1))

'---- Execute main
call connect_db()
call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
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
'	試聴･Movieリンク作成
'
'========================================================================
Function main()
Dim i

'---- 商品データSELECT
w_sql = ""
w_sql = w_sql & "SELECT b.メーカー名"
w_sql = w_sql & "     , a.商品名"
w_sql = w_sql & "     , a.商品画像ファイル名_小"
w_sql = w_sql & "     , a.試聴フラグ"
w_sql = w_sql & "     , a.試聴URL"
w_sql = w_sql & "     , a.動画フラグ"
w_sql = w_sql & "     , a.動画URL"
w_sql = w_sql & "  FROM Web商品                a WITH (NOLOCK)"   '2012/01/20 an mod s
w_sql = w_sql & "         INNER JOIN  メーカー b WITH (NOLOCK)"
w_sql = w_sql & "           ON     b.メーカーコード = a.メーカーコード"
'w_sql = w_sql & "     , メーカー b"
'w_sql = w_sql & " WHERE b.メーカーコード = a.メーカーコード"     '2012/01/20 an mod e
w_sql = w_sql & "   AND a.メーカーコード = '" & MakerCd & "'"
w_sql = w_sql & "   AND a.商品コード = '" & ProductCd & "'"

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic

if RS.EOF = true then
	exit function
end if

'---- 明細HTML作成
MakerName = RS("メーカー名")
ProductName = RS("商品名")
ImageFileName = RS("商品画像ファイル名_小")

'----試聴リンク
if RS("試聴フラグ") = "Y" AND RS("試聴URL") <> "" then
	ItemCnt = cf_unstring(RS("試聴URL"), ItemList, ",")
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

'----動画リンク
if RS("動画フラグ") = "Y" AND RS("動画URL") <> "" then
	ItemCnt = cf_unstring(RS("動画URL"), ItemList, ",")
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
<title>試聴・Movie選択</title>

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
            <img src='images/Shichou.gif' width='18' height='18' border='0' alt='試聴する'>
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
            <img src='images/Movie.jpg' width='18' height='18' border='0' alt='動画を見る'>
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
