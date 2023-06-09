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
'	Emaxの[受注].[Web注文変更キャンセル中フラグ]を"Y"または"N"に更新する
'
'変更履歴
'2016/02/04 GV 新規作成。(注文変更キャンセル機能)
'2020.12.07 GV 受注明細のWebキャンセルフラグ更新改修。(#2619)
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim wDefer						' 変更モード(Y/N)
Dim wFlg						' 実行フラグ
Dim oJSON						' JSONオブジェクト

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
wFlg = True

' Getパラメータ
' 顧客番号
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
	wFlg = False
End If

' 受注番号
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wFlg = False
End If

'保留モード
wDefer = ReplaceInput_NoCRLF(Trim(Request("defer")))
wDefer = UCase(wDefer)
If wFlg = True Then
	Select Case wDefer
		Case "Y"
			wFlg = True
		Case "N"
			wFlg = True
		Case Else
			wFlg = False
	End Select
End If

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- エラーメッセージをセッションデータに登録   ' member系の他のページ処理にならう
If Err.Description <> "" Then
End If

Call close_db()

Call sendResponse()

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

ConnectionEmax.close
Set ConnectionEmax= Nothing

End function

'========================================================================
'
'	Function	Main
'
'変更履歴
'2020.12.07 GV 受注明細のWebキャンセルフラグ更新改修。(#2619)
'========================================================================
Function main()

Dim vSQL
Dim vRS
Dim vRS2 '2020.12.07 GV add
Dim i '2020.12.07 GV add
Set oJSON = New aspJSON


' 入力値が正常の場合
If (wFlg = True) Then
	'---- トランザクション開始
	ConnectionEmax.BeginTrans

	'受注の取り出し
	vSQL = ""
	vSQL = vSQL & "SELECT a.* "
	'vSQL = vSQL & "  FROM 受注 a WITH(UPDLOCK) "
	vSQL = vSQL & "  FROM 受注 a "
	vSQL = vSQL & " WHERE 受注番号 = " & wOrderNo
	vSQL = vSQL & "  AND 顧客番号 = " & wCustomerNo
	vSQL = vSQL & "  AND 削除日 IS NULL "
	'@@@@Response.Write vSQL & "<br>"

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在している場合
	If vRS.EOF = False Then
		'vSQL = ""
		'vSQL = vSQL & "UPDATE 受注 "
		'vSQL = vSQL & " SET "
		'vSQL = vSQL & " Web注文変更キャンセル中フラグ = " & wDefer
		'vSQL = vSQL & " WHERE 受注番号 = " & wOrderNo
		'vSQL = vSQL & "  AND 顧客番号 = " & wCustomerNo
		'vSQL = vSQL & "  AND 削除日 = IS NULL "
		vRS("Web注文変更キャンセル中フラグ") = wDefer
		vRS.update

		wFlg = True
	Else
		wFlg = False
	End If

	'2020.12.07 GV add start
	If (wDefer = "N") Then
		'受注明細の取り出し
		vSQL = ""
		vSQL = vSQL & "SELECT a.* "
		vSQL = vSQL & "  FROM 受注明細 a "
		vSQL = vSQL & " WHERE 受注番号 = " & wOrderNo
		vSQL = vSQL & " AND  Webキャンセルフラグ = 'Y' "
		vSQL = vSQL & " ORDER BY 受注明細番号 "

		Set vRS2 = Server.CreateObject("ADODB.Recordset")
		vRS2.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

		'レコードが存在している場合
		If vRS2.EOF = False Then
			For i = 0 To (vRS2.RecordCount - 1)
				vRS2("Webキャンセルフラグ") = wDefer
				vRS2.update

				' 次のレコード行へ移動
				vRS2.MoveNext

				If vRS2.EOF Then
					Exit For
				End If
			Next
		End If
	End If
	' 2020.12.07 GV add end


	'成功の場合
	If (wFlg = True) Then
		'コミット
		ConnectionEmax.CommitTrans
	Else
		'ロールバック
		ConnectionEmax.RollbackTrans
	End If

	'レコードセットを閉じる
	vRS.Close

	'レコードセットのクリア
	Set vRS = Nothing

	'2020.12.07 GV add start
	'レコードセットを閉じる
	vRS2.Close
	'レコードセットのクリア
	Set vRS2 = Nothing
	'2020.12.07 GV add end

End If
End Function


'========================================================================
'
'	Function	JSON返却
'
'========================================================================
Function sendResponse()

	' 全件数をJSONデータにセット
	oJSON.data.Add "ono" ,wOrderNo
	oJSON.data.Add "cno" ,wCustomerNo
	oJSON.data.Add "defer" ,wDefer

	If wFlg = True Then
		oJSON.data.Add "result" ,"Y"
	Else
		oJSON.data.Add "result" ,"N"
	End If

	' -------------------------------------------------
	' JSONデータの返却
	' -------------------------------------------------
	' ヘッダ出力
	Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
	Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
	Response.AddHeader "Pragma", "no-cache"
	Response.AddHeader "X-Content-Type-Options", "nosniff"

	' JSONデータの出力
	Response.Write oJSON.JSONoutput()

End Function

'========================================================================
%>
