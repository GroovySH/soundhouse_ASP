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
<!--#include file="../3rdParty/aspJSON1.17.asp"-->
<%
'========================================================================
'
'	購入履歴一覧ページ
'
'
'変更履歴
'2014/09/22 GV 新規作成
'
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wCustomerNo					' 顧客番号

Dim oJSON						' JSONオブジェクト
Dim wOrderNo					' 受注番号

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
'wUserID = ReplaceInput(Trim(Request("cno")))
wCustomerNo = ReplaceInput(Trim(Request("customer_no")))
wOrderNo = ReplaceInput(Trim(Request("order_no")))

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- エラーメッセージをセッションデータに登録   ' member系の他のページ処理にならう
If Err.Description <> "" Then
'	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
'	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If

Call close_db()

If Err.Description <> "" Then

End If


'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.close
Set Connection= Nothing

ConnectionEmax.close
Set ConnectionEmax= Nothing

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

Dim vSQL
Dim vRS
Dim oJSON
Dim flag

Set oJSON = New aspJSON

flag = "0"


'--- 出荷完了データの情報取出し
vSQL = ""
vSQL = vSQL & "SELECT *"
vSQL = vSQL & "  FROM 受注 a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "     a.顧客番号 = " & wCustomerNo & " AND " 
vSQL = vSQL & "     a.受注番号 = " & wOrderNo
'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic


If vRS.RecordCount > 0 Then
	flag = "1"
Else
	flag = "0"
End If

'レコードセットを閉じる
vRS.Close

'レコードセットのクリア
Set vRS = Nothing

' リスト追加
oJSON.data.Add "result" ,flag


' -------------------------------------------------
' JSONデータの返却
' -------------------------------------------------
' ヘッダ出力
Response.AddHeader "Content-Type", "application/json"
' JSONデータの出力
Response.Write oJSON.JSONoutput()

End Function
%>
