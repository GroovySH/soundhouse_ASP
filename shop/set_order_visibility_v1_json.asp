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
'	購入履歴一覧ページ上に非表示/表示のフラグをセットする。
'
'
'変更履歴
'2016.02.12 GV 新規作成
'
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wFlg						' 実行フラグ
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim wDetailNo					' 明細番号
Dim wMode						' モード(Y ... 追加、N ... 削除)
Dim oJSON						' JSONオブジェクト


' 初期設定
wFlg = True

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
' 顧客番号
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
	wFlg = False
End If

' 注文番号
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wFlg = False
End If

'明細番号
wDetailNo = ReplaceInput_NoCRLF(Trim(Request("dno")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wDetailNo) = False) Or (cf_checkNumeric(wDetailNo) = False) Then
	wFlg = False
End If

' モード
wMode = ReplaceInput_NoCRLF(Trim(Request("mode")))
wMode = UCase(wMode) ' 大文字化
' チェック
If cf_checkHankaku2(wMode) = False Then
	wFlg = False
Else
	If (wMode = "Y") Or (wMode = "N") Then
	Else
		wFlg = False
	End If
End If

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
'========================================================================
Function main()

Dim vSQL
Dim i
Dim j
Dim vRS1
Dim vRS2
Dim vOrderNo
Dim vDetailNo
Dim vCustomerNo

Set oJSON = New aspJSON


' 入力値が正常の場合
If (wFlg = True) Then
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "	 T1.顧客番号 "
	vSQL = vSQL & " ,T1.受注番号 "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "  受注 T1 WITH (NOLOCK)"
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "  T1.顧客番号 = " & wCustomerNo
	vSQL = vSQL & " AND T1.受注番号 = " & wOrderNo

	Set vRS1 = Server.CreateObject("ADODB.Recordset")
	vRS1.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在する場合
	If vRS1.EOF = False Then
		vSQL = ""
		vSQL = vSQL & "SELECT "
		vSQL = vSQL & "	 T1.顧客番号 "
		vSQL = vSQL & " ,T1.受注番号 "
		vSQL = vSQL & " ,T1.受注明細番号 "
		vSQL = vSQL & " ,T1.非表示フラグ "
		vSQL = vSQL & "FROM "
		vSQL = vSQL & "  受注非表示リスト T1 "
		vSQL = vSQL & "WHERE "
		vSQL = vSQL & "  T1.顧客番号 = " & wCustomerNo
		vSQL = vSQL & " AND T1.受注番号 = " & wOrderNo
		vSQL = vSQL & " AND T1.受注明細番号 = " & wDetailNo

		Set vRS2 = Server.CreateObject("ADODB.Recordset")
		vRS2.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

		'レコードが存在する場合
		If vRS2.EOF = False Then
			' モードが非表示リスト追加の場合
			If (wMode = "Y") Then
				'何もしない
			ElseIf (wMode = "N") Then
				' モードが非表示リストから削除の場合、削除
				vRS2.Delete
			End If
		Else
			' レコードが存在しなかった場合
			' モードが非表示リスト追加の場合
			If (wMode = "Y") Then
				vRS2.AddNew
				vRS2("顧客番号") = wCustomerNo
				vRS2("受注番号") = wOrderNo
				vRS2("受注明細番号") = wDetailNo
				vRS2.Update
			End If
		End If

		'レコードセットを閉じる
		vRS2.Close

		'レコードセットのクリア
		Set vRS2 = Nothing
	Else
		'受注が存在しなかった場合
		wFlg = false
	End If

	'レコードセットを閉じる
	vRS1.Close

	'レコードセットのクリア
	Set vRS1 = Nothing
End If
	' 結果
	oJSON.data.Add "result" ,wFlg

	'顧客番号
	If (IsNull(wCustomerNo)) Then
		vCustomerNo = ""
	Else
		vCustomerNo = CStr(Trim(wCustomerNo))
	End If

	oJSON.data.Add "cno" ,vCustomerNo


	'受注番号
	If (IsNull(wOrderNo)) Then
		vOrderNo = ""
	Else
		vOrderNo = CStr(Trim(wOrderNo))
	End If

	'受注明細番号
	If (IsNull(wDetailNo)) Then
		vDetailNo = ""
	Else
		vDetailNo = CStr(Trim(wDetailNo))
	End If

	oJSON.data.Add "ono" ,vOrderNo
	oJSON.data.Add "dno" ,vDetailNo

	'モード
	oJSON.data.Add "mode" ,wMode

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
