get_order_cancel<%@ LANGUAGE="VBScript" %>
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
'	Emax受注明細注文キャンセル 取得API
'========================================================================
'On Error Resume Next

Const PAGE_SIZE = 10			' 購入履歴情報の1ページあたりの表示行数

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim oJSON						' JSONオブジェクト
Dim wFlg						' 実行フラグ

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

'受注番号
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' 数値のみチェック (ASPは全角でも数字ならTrueを返す)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wOrderNo = null
Else
	wOrderNo = CLng(wOrderNo)
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
Dim vWHERE
Dim i
Dim j
Dim vRS
Dim point

Set oJSON = New aspJSON

' 初期化
i = 0
j = 0

' 入力値が正常の場合
If (wFlg = True) Then

	'--- 受注明細注文キャンセル情報取出し
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "      order_cancel_details.注文キャンセル番号 "
	vSQL = vSQL & "    , order_cancel_details.注文キャンセル明細番号 "
	vSQL = vSQL & "    , order_cancel_details.受注番号 "
	vSQL = vSQL & "    , order_cancel_details.受注明細番号 "
	vSQL = vSQL & "    , order_cancel_details.メーカーコード "
	vSQL = vSQL & "    , order_cancel_details.商品コード "
	vSQL = vSQL & "    , order_cancel_details.商品名 "
	vSQL = vSQL & "    , order_cancel_details.色 "
	vSQL = vSQL & "    , order_cancel_details.規格 "
	vSQL = vSQL & "    , order_cancel_details.セット品フラグ "
	vSQL = vSQL & "    , order_cancel_details.セット品親明細番号 "
	vSQL = vSQL & "    , order_cancel_details.注文キャンセル数量 "
	vSQL = vSQL & "    , order_cancel_details.注文キャンセル単価 "
	vSQL = vSQL & "    , order_cancel_details.注文キャンセル金額 "
	vSQL = vSQL & "    , order_cancel_details.外税 "
	vSQL = vSQL & "    , order_cancel_details.内税 "
	vSQL = vSQL & "    , order_cancel_details.ポイント "
	vSQL = vSQL & "    , order_cancel_details.クーポン値引き "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "      " & gLinkServer & "受注注文キャンセル order_cancel WITH (NOLOCK) "
	vSQL = vSQL & " INNER JOIN " & gLinkServer & "受注 orders WITH (NOLOCK) "
	vSQL = vSQL & "   ON orders.受注番号 = order_cancel.受注番号 "
	vSQL = vSQL & " INNER JOIN " & gLinkServer & "受注明細注文キャンセル order_cancel_details WITH (NOLOCK) "
	vSQL = vSQL & "   ON order_cancel.受注番号 = order_cancel.受注番号 "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "        orders.受注番号 = " & wOrderNo
	vSQL = vSQL & "    AND orders.顧客番号 = " & wCustomerNo & " "

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在している場合
	If vRS.EOF = False Then

		' リスト追加
		oJSON.data.Add "list" ,oJSON.Collection()

		For i = 0 To (vRS.RecordCount - 1)
			' ポイント
			If (IsNull(vRS("ポイント"))) Then
				point = 0
			Else
				point = CDbl(vRS("ポイント"))
			End If

			'--- 明細行生成
			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
					.Add "order_cancel_no" ,CStr(Trim(vRS("注文キャンセル番号")))
					.Add "order_cancel_detail_no" ,CStr(Trim(vRS("注文キャンセル明細番号")))
					.Add "o_no" ,CStr(Trim(vRS("受注番号")))
					.Add "od_no" ,CStr(Trim(vRS("受注明細番号")))
					.Add "maker_code" ,CStr(Trim(vRS("メーカーコード")))
					.Add "i_cd" ,CStr(Trim(vRS("商品コード")))
					.Add "i_name" ,CStr(Trim(vRS("商品名")))
					.Add "iro" ,CStr(Trim(vRS("色")))
					.Add "kikaku" ,CStr(Trim(vRS("規格")))
					.Add "set_item_flg" ,CStr(Trim(vRS("セット品フラグ")))
					.Add "set_item_detail_no" ,CStr(Trim(vRS("セット品親明細番号")))
					.Add "i_suu" ,CStr(Trim(vRS("注文キャンセル数量")))
					.Add "i_tanka" ,CStr(Trim(vRS("注文キャンセル単価")))
					.Add "i_cancel_am" ,CStr(Trim(vRS("注文キャンセル金額")))
					.Add "ext_tax" ,CStr(Trim(vRS("外税")))
					.Add "inc_tax" ,CStr(Trim(vRS("内税")))
					.Add "point" ,point 'ポイント
					.Add "coupon_discount" ,CStr(Trim(vRS("クーポン値引き")))
				End With
			End With

			' 次のレコード行へ移動
			vRS.MoveNext

			If vRS.EOF Then
				Exit For
			End If

			j = j + 1
		Next
	End If

	'レコードセットを閉じる
	vRS.Close

	'レコードセットのクリア
	Set vRS = Nothing
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
%>
