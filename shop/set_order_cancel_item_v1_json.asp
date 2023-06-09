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
'	購入履歴
'	キャンセルした商品にフラグをセットする。
'
' 変更可能状態
'   Web注文（インターネット、スマートフォン）である。
'   支払い方法が「ローン」以外である。
'   受注ステータスが「受注」(出荷指示あり)でない。
'   メーカー直送品が含まれていない。
'   注残かつ適正在庫数量=0でない。
'
'
'変更履歴
'2016/02/04 GV 新規作成。(注文変更キャンセル機能)
'2020.10.20 GV 注文変更キャンセル時のEmax購入履歴データ修正。(#2572)
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim wOrderDetailNo				' 受注番号
Dim wSetFlg						' 変更モード(Y/N)
Dim wFlg						' 実行フラグ
Dim oJSON						' JSONオブジェクト
Dim modifyFlag					' 変更可能フラグ
Dim wNgReason					' 不可理由
Dim wDepositFlag   				' 入金完了フラグ
Dim wDepositAmount 				' 入金合計金額
Dim isUpdateOrderTable			' 受注テーブルも更新する(Y/N) ' 2020.10.20 GV add

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


'フラグ
wSetFlg = ReplaceInput_NoCRLF(Trim(Request("f")))
wSetFlg = UCase(wSetFlg)
If (wSetFlg <> "Y") And (wSetFlg <> "N") And (wSetFlg <> "") Then
	wFlg = False
End If


' 受注明細番号(カンマ区切り)
wOrderDetailNo = ReplaceInput_NoCRLF(Trim(Request("od")))
If (wOrderDetailNo = "") Then
	wFlg = False
Else
	Dim pos
	' アンダースコアの位置を取得
	pos = InStr(wOrderDetailNo, "_")

	'文字列があり、アンダースコアが含まれている場合
	If (Len(wOrderDetailNo) > 0) And (pos > 0) Then
		wOrderDetailNo = Replace(wOrderDetailNo, "_", ",")
	End If
End If

'2020.10.20 GV add
'受注テーブル更新フラグ
isUpdateOrderTable = ReplaceInput_NoCRLF(Trim(Request("up_od")))
If (isUpdateOrderTable <> "") Then
	isUpdateOrderTable = UCase(isUpdateOrderTable)

	If (isUpdateOrderTable <> "Y") And (isUpdateOrderTable <> "N") Then
		wFlg = False
	End If
End If


wNgReason = ""
wDepositFlag = ""
wDepositAmount = 0


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
'========================================================================
Function main()

Dim vSQL
Dim vRS1
Dim vRS2
Dim okFlag
Dim wSQL
Dim orderDate
Dim deleteDate
Dim promote
Dim i

' JSONオブジェクト生成
Set oJSON = New aspJSON

okFlag = True

' 入力値が正常の場合
If (wFlg = True) Then
	'---- トランザクション開始
	ConnectionEmax.BeginTrans

	'受注の取り出し
	wSQL = ""
	wSQL = wSQL & "SELECT "
	wSQL = wSQL & "  受注形態 "
	wSQL = wSQL & " ,支払方法 "
	wSQL = wSQL & " ,受注日 "
	wSQL = wSQL & " ,削除日 "
	wSQL = wSQL & " ,削除日 "
	wSQL = wSQL & " ,Web注文変更キャンセル中フラグ "
	wSQL = wSQL & " ,入金合計金額 "
	wSQL = wSQL & " ,入金完了フラグ "
'	wSQL = wSQL & "  FROM 受注 WITH(NOLOCK) " ' 2020.10.20 GV mod
	wSQL = wSQL & "  FROM 受注 " ' 2020.10.20 GV add

	'2020.10.20 GV add start
	' 受注データ更新フラグがYでない
	If (isUpdateOrderTable <> "Y") Then
		wSQL = wSQL & " WITH (NOLOCK) "
	End If
	'2020.10.20 GV add end

	wSQL = wSQL & " WHERE 受注番号 = " & wOrderNo
	wSQL = wSQL & "  AND 顧客番号 = " & wCustomerNo
	'Response.Write wSQL & "<br>"

	Set vRS1 = Server.CreateObject("ADODB.Recordset")
	vRS1.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在している場合
	If vRS1.EOF = False Then
		If (okFlag = True) Then
			'2020.10.20 GV add start
			' 受注データ更新フラグがYである
			If (isUpdateOrderTable = "Y") Then
				vRS1("Web注文変更キャンセル中フラグ") = wSetFlg
				vRS1.update
			End If
			'2020.10.20 GV add end

			wSQL = ""
			wSQL = wSQL & "SELECT "
			wSQL = wSQL & "  受注明細番号 "
			wSQL = wSQL & " ,Webキャンセルフラグ "
			wSQL = wSQL & "FROM "
			wSQL = wSQL & "  受注明細 "
			wSQL = wSQL & "WHERE "
			wSQL = wSQL & "   受注番号 = " & wOrderNo
			wSQL = wSQL & "   AND 受注明細番号 IN (" & wOrderDetailNo & ") "

			'@@@@Response.Write wSQL & "<br>"

			Set vRS2 = Server.CreateObject("ADODB.Recordset")
			vRS2.Open wSQL, ConnectionEmax, adOpenStatic, adLockOptimistic
	
			'レコードが存在している場合
			If vRS2.EOF = False Then
				For i = 0 To (vRS2.RecordCount - 1)
					vRS2("Webキャンセルフラグ") = wSetFlg
					vRS2.update
					okFlag = True

					' 次のレコード行へ移動
					vRS2.MoveNext

					If vRS2.EOF Then
						Exit For
					End If
				Next
			Else
				okFlag = False
				wNgReason = "6"
			End If
		End If
	Else
		'レコードがない場合、NG
		okFlag = False
		wNgReason = "7"
	End If

	If okFlag = True Then
		'コミット
		ConnectionEmax.CommitTrans

		'レコードセットを閉じる
		vRS2.Close

		'レコードセットのクリア
		Set vRS2 = Nothing
	Else
		'ロールバック
		ConnectionEmax.RollbackTrans
	End If

	'レコードセットを閉じる
	vRS1.Close

	'レコードセットのクリア
	Set vRS1 = Nothing
Else
	'入力値がNGの場合
	okFlag = False
	wNgReason = "99"
End If

wFlg = okFlag

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
	oJSON.data.Add "od_no" ,wOrderDetailNo
	oJSON.data.Add "set_flg" ,wSetFlg
	oJSON.data.Add "reason" ,wNgReason
	oJSON.data.Add "is_up_od_tbl" ,isUpdateOrderTable

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
