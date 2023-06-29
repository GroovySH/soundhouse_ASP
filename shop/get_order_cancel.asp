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
'	Emax受注注文キャンセル 取得API
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)

Dim oJSON						' JSONオブジェクト
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
Dim wGiftCustomerNo				' ギフト顧客番号
Dim wGiftNo						' ギフト番号
Dim wOrderGift					' ギフト注文フラグ

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
wCustomerNo = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))

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
Dim vRS

Dim createDate
Dim totalOrderAmount
Dim usedPoint
Dim orderTotalOrderAmount

Set oJSON = New aspJSON

'--- 受注注文キャンセル情報取出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      order_cancel.注文キャンセル番号 "
vSQL = vSQL & "    , order_cancel.受注番号 "
vSQL = vSQL & "    , order_cancel.本登録日 "
vSQL = vSQL & "    , order_cancel.商品合計金額 "
vSQL = vSQL & "    , order_cancel.その他合計金額 "
vSQL = vSQL & "    , order_cancel.送料 "
vSQL = vSQL & "    , order_cancel.代引手数料 "
vSQL = vSQL & "    , order_cancel.外税合計金額 "
vSQL = vSQL & "    , order_cancel.内税合計金額 "
vSQL = vSQL & "    , order_cancel.受注合計金額 "
vSQL = vSQL & "    , order_cancel.過不足相殺金額 "
vSQL = vSQL & "    , order_cancel.利用ポイント "
vSQL = vSQL & "    , order_cancel.合計金額 "
vSQL = vSQL & "    , order_cancel.値引き後消費税 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注注文キャンセル order_cancel WITH (NOLOCK) "
vSQL = vSQL & " INNER JOIN " & gLinkServer & "受注 orders WITH (NOLOCK) "
vSQL = vSQL & "   ON orders.受注番号 = order_cancel.受注番号 "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        orders.受注番号 = " & wOrderNo
vSQL = vSQL & "    AND orders.顧客番号 = " & wCustomerNo & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' リスト追加
	oJSON.data.Add "data" ,oJSON.Collection()

	' 本登録日
	If (IsNull(vRS("本登録日"))) Then
		createDate = ""
	Else
		createDate = CStr(Trim(vRS("本登録日")))
	End If

	' 合計金額
	If (IsNull(vRS("合計金額"))) Then
		totalOrderAmount = 0
	Else
		totalOrderAmount = CDbl(vRS("合計金額"))
	End If

	' 利用ポイント
	If (IsNull(vRS("利用ポイント"))) Then
		usedPoint = 0
	Else
		usedPoint = CDbl(vRS("利用ポイント"))
	End If

	' 受注合計金額
	If (IsNull(vRS("受注合計金額"))) Then
		orderTotalOrderAmount = 0
	Else
		orderTotalOrderAmount = CDbl(vRS("受注合計金額"))
	End If

	With oJSON.data("data")
		.Add "order_no", CStr(Trim(vRS("受注番号")))
		.Add "order_cancel_no", CStr(Trim(vRS("注文キャンセル番号")))
		.Add "create_date", createDate
		.Add "item_total_amount", CStr(Trim(vRS("商品合計金額")))
		.Add "other_total_amount", CStr(Trim(vRS("その他合計金額")))
		.Add "ff_charge", CDbl(vRS("送料")) 
		.Add "cod_charge", CStr(Trim(vRS("代引手数料")))
		.Add "tax_am", CStr(Trim(vRS("外税合計金額")))
		.Add "tax_in", CStr(Trim(vRS("内税合計金額")))
		.Add "order_total_order_amount", orderTotalOrderAmount ' 受注合計金額
		.Add "kabusoku_am", CStr(Trim(vRS("過不足相殺金額")))
		.Add "used_point", usedPoint ' 利用ポイント
		.Add "total_order_amount", totalOrderAmount ' 合計金額
		.Add "after_discount_tax", CStr(Trim(vRS("値引き後消費税")))
	End With
End If

'レコードセットを閉じる
vRS.Close

'レコードセットのクリア
Set vRS = Nothing

' -------------------------------------------------
' JSONデータの返却
' -------------------------------------------------
' ヘッダ出力
Response.AddHeader "Content-Type", "application/json"
' JSONデータの出力
Response.Write oJSON.JSONoutput()

End Function
%>
