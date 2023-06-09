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
'2014/09/16 GV 新規作成
'
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wUserID

Dim oJSON						' JSONオブジェクト
Dim wOrderNo					' 受注番号

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
wUserID = ReplaceInput(Trim(Request("cno")))
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

Dim orderDate
Dim shippingDate
Dim estimateDate
Dim one_time_todokesaki
Dim final_nouki_date_time
Dim receiptName
Dim receiptNote

Set oJSON = New aspJSON


one_time_todokesaki = ""
final_nouki_date_time = ""
receiptName = ""
receiptNote = ""

'--- ヘッダ部分の情報取出し
vSQL = ""
vSQL = vSQL & "SELECT TOP 1 "
vSQL = vSQL & "      a.受注番号 "
vSQL = vSQL & "    , a.見積日 "
vSQL = vSQL & "    , a.受注日 "
vSQL = vSQL & "    , a.出荷完了日 "
vSQL = vSQL & "    , a.受注形態 "
vSQL = vSQL & "    , a.支払方法 "
vSQL = vSQL & "    , a.商品合計金額 "
vSQL = vSQL & "    , a.送料 "
vSQL = vSQL & "    , a.代引手数料 "
vSQL = vSQL & "    , a.受注合計金額 "
vSQL = vSQL & "    , a.一括出荷フラグ "
vSQL = vSQL & "    , a.領収書宛先 "
vSQL = vSQL & "    , a.領収書但し書き "
vSQL = vSQL & "    , a.Web受注変更開始日 "
vSQL = vSQL & "    , a.消費税率 "
vSQL = vSQL & "    , a.運送会社コード "
vSQL = vSQL & "    , a.担当者コード "
vSQL = vSQL & "    , b.今回限り届先郵便番号 "
vSQL = vSQL & "    , b.今回限り届先都道府県 "
vSQL = vSQL & "    , b.今回限り届先住所 "
vSQL = vSQL & "    , b.今回限り届先名前 "
vSQL = vSQL & "    , b.最終指定納期 "
vSQL = vSQL & "    , b.最終時間指定 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注     a WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "受注明細 b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        b.受注番号 = a.受注番号 "
vSQL = vSQL & "    AND a.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "    AND a.顧客番号 = " & wUserID & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' リスト追加
	oJSON.data.Add "data" ,oJSON.Collection()

	' 受注日
	If (IsNull(vRS("受注日"))) Then
		orderDate = ""
	Else
		orderDate = CStr(Trim(vRS("受注日")))
	End If

	' 見積日
	If (IsNull(vRS("見積日"))) Then
		estimateDate = ""
	Else
		estimateDate = CStr(Trim(vRS("見積日")))
	End If

	' 出荷完了日
	If (IsNull(vRS("出荷完了日"))) Then
		shippingDate = ""
	Else
		shippingDate = CStr(Trim(vRS("出荷完了日")))
	End If

	' 今回限り届先
	one_time_todokesaki = vRS("今回限り届先郵便番号")&"^"&_
		vRS("今回限り届先都道府県")&"^"&_
		vRS("今回限り届先住所")&"^"&_
		vRS("今回限り届先名前")

	one_time_todokesaki = Replace(one_time_todokesaki, """", "”")

	' 最終指定納期と時刻
	final_nouki_date_time = vRS("最終指定納期")&"_"&vRS("最終時間指定") 

	' 領収書宛先
	If (IsNull(vRS("領収書宛先"))) Then
		receiptName = ""
	Else
		receiptName = CStr(Trim(vRS("領収書宛先")))
		receiptName = Replace(receiptName, """", "”")
	End If

	' 領収書但し書き
	If (IsNull(vRS("領収書但し書き"))) Then
		receiptNote = ""
	Else
		receiptNote = CStr(Trim(vRS("領収書但し書き")))
		receiptNote = Replace(receiptNote, """", "”")
	End If

	With oJSON.data("data")
		.Add "order_no", CStr(Trim(vRS("受注番号")))
		.Add "estimate_date", estimateDate
		.Add "order_date", orderDate
		.Add "shipping_date", shippingDate
		.Add "order_type", CStr(Trim(vRS("受注形態")))
		.Add "payment_method",  CStr(Trim(vRS("支払方法")))
		.Add "total_item_amount", CDbl(Trim(vRS("商品合計金額")))
		.Add "freight_charge", CDbl(vRS("送料")) 
		.Add "daibiki_charge", CDbl(vRS("代引手数料"))
		.Add "total_order_amount", CDbl(vRS("受注合計金額"))
		.Add "combined_shipping_flag", CStr(Trim(vRS("一括出荷フラグ")))
		.Add "receipt_name", receiptName
		.Add "receipt_note", receiptNote
'				.Add "web_order_modify_start_date", vRS("Web受注変更開始日")
		.Add "tax_rate", CDbl(vRS("消費税率"))
		.Add "freight_forwarder_cd", CStr(vRS("運送会社コード"))
		.Add "tantou_cd", CStr(vRS("担当者コード"))
		.Add "one_time_todokesaki", one_time_todokesaki
		.Add "final_nouki_date_time", final_nouki_date_time
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
