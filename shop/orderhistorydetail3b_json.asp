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
'2015.03.09 GV 出荷が別れた場合の表示不具合対応。
'2016.05.09 GV セット品で出荷が別れた場合の表示不具合対応。
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
'wUserID = ReplaceInput(Trim(Request("cno")))
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
Dim webOutline
Dim source
Dim shippingText
Dim itemPicSmall
Dim slipNo
Dim i
Dim invoiceNo
Dim shippingNo
Dim makerName
Dim itemName
Dim setItem '2016.05.09 GV add
Dim shipSuu '2016.05.09 GV add

Set oJSON = New aspJSON


one_time_todokesaki = ""
final_nouki_date_time = ""
receiptName = ""
receiptNote = ""
webOutline = ""
source = ""
shippingText = ""
itemPicSmall = ""
slipNo = ""
invoiceNo = ""
shippingNo = ""
makerName = ""
itemName = ""
setItem = "" '2016.05.09 GV add
i = 0

'--- 出荷完了データの情報取出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      b.受注明細番号 "
vSQL = vSQL & "    , b.メーカーコード "
vSQL = vSQL & "    , b.商品コード "
vSQL = vSQL & "    , b.色 "
vSQL = vSQL & "    , b.規格 "
vSQL = vSQL & "    , z.商品ID "		'2015.03.09 GV add
vSQL = vSQL & "    , b.受注単価 "
vSQL = vSQL & "    , b.受注数量 "
vSQL = vSQL & "    , b.受注引当合計数量 "
vSQL = vSQL & "    , b.出荷合計数量 "
vSQL = vSQL & "    , sum(f.出荷数量) AS 出荷数量 "
vSQL = vSQL & "    , c.メーカー名 "
vSQL = vSQL & "    , d.商品名 "
vSQL = vSQL & "    , d.商品概略Web "
vSQL = vSQL & "    , d.商品画像ファイル名_小 "
vSQL = vSQL & "    , d.Web商品フラグ "
vSQL = vSQL & "    , d.セット商品フラグ " '2016.05.09 GV add
'vSQL = vSQL & "    , e.送り状番号 "	'2015.03.09 GV mod
vSQL = vSQL & "    , NULL AS 出荷予定日 "
vSQL = vSQL & "    , NULL AS 出荷予定テキスト "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注明細     b WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "メーカー     c WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "商品         d WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "受注送り状   e WITH (NOLOCK) "	'2015.03.09 GV mod
vSQL = vSQL & "    , " & gLinkServer & "出荷明細View f WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "色規格別在庫 z WITH (NOLOCK) "	'2015.03.09 GV add
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        c.メーカーコード = b.メーカーコード "
vSQL = vSQL & "    AND d.メーカーコード = b.メーカーコード "
vSQL = vSQL & "    AND d.商品コード = b.商品コード "
vSQL = vSQL & "    AND e.受注番号 = b.受注番号 "	'2015.03.09 GV mod
vSQL = vSQL & "    AND f.出荷番号 = e.出荷番号 "	'2015.03.09 GV mod
vSQL = vSQL & "    AND f.受注番号 = b.受注番号 "
vSQL = vSQL & "    AND f.受注明細番号 = b.受注明細番号 "
vSQL = vSQL & "    AND f.セット品親明細番号 = 0 "
'2015.03.09 GV add start
vSQL = vSQL & "    AND z.メーカーコード = b.メーカーコード "
vSQL = vSQL & "    AND z.商品コード = b.商品コード "
vSQL = vSQL & "    AND z.色 = b.色 "
vSQL = vSQL & "    AND z.規格 = b.規格 "
'2015.03.09 GV add end
vSQL = vSQL & "    AND b.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "GROUP BY  "
vSQL = vSQL & "      b.受注明細番号 "
vSQL = vSQL & "    , b.メーカーコード "
vSQL = vSQL & "    , b.商品コード "
vSQL = vSQL & "    , b.色 "
vSQL = vSQL & "    , b.規格 "
vSQL = vSQL & "    , z.商品ID "		'2015.03.09 GV add
vSQL = vSQL & "    , b.受注単価 "
vSQL = vSQL & "    , b.受注数量 "
vSQL = vSQL & "    , b.受注引当合計数量 "
vSQL = vSQL & "    , b.出荷合計数量 "
vSQL = vSQL & "    , c.メーカー名 "
vSQL = vSQL & "    , d.商品名 "
vSQL = vSQL & "    , d.商品概略Web "
vSQL = vSQL & "    , d.商品画像ファイル名_小 "
vSQL = vSQL & "    , d.Web商品フラグ "
vSQL = vSQL & "    , d.セット商品フラグ " '2016.05.09 GV add
vSQL = vSQL & "ORDER BY  "
'vSQL = vSQL & "      e.送り状番号 "	'2015.03.09 GV mod
vSQL = vSQL & "      b.受注明細番号 "	'2015.03.09 GV add
vSQL = vSQL & "    , c.メーカー名 "
vSQL = vSQL & "    , d.商品名 "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' リスト追加
	oJSON.data.Add "list" ,oJSON.Collection()

	Do Until vRS.EOF = True
		' 出荷予定日
		If (IsNull(vRS("出荷予定日"))) Then
			shippingDate = ""
		Else
			shippingDate = CStr(Trim(vRS("出荷予定日")))
		End If


		If (IsNull(vRS("商品概略Web"))) Then
			webOutline = ""
		Else
			webOutline = CStr(vRS("商品概略Web"))
			webOutline = Replace(Trim(webOutline), """", "”")
			webOutline = Replace(Trim(webOutline), vbCrLf, "")
		End If

		If (IsNull(vRS("出荷予定テキスト"))) Then
			shippingText = ""
		Else
			shippingText = CStr(vRS("出荷予定テキスト"))
			shippingText = Replace(Trim(shippingText), """", "”")
		End If

		If (IsNull(vRS("商品画像ファイル名_小"))) Then
			itemPicSmall = ""
		Else
			itemPicSmall = CStr(vRS("商品画像ファイル名_小"))
		End If

		'2015.03.09 GV mod start
		'If (IsNull(vRS("送り状番号"))) Then
		'	shippingNo = ""
		'Else
		'	shippingNo = CStr(vRS("送り状番号"))
		'End If
		'2015.03.09 GV mod start

		makerName = Replace(Trim(vRS("メーカー名")), """", "”")
		makerName = CStr(makerName)

		itemName = Replace(Trim(vRS("商品名")), """", "”")
		itemName = CStr(itemName)

		'2016.05.09 GV add start
		If (IsNull(vRS("セット商品フラグ"))) Then
			setItem = ""
		Else
			setItem = CStr(vRS("セット商品フラグ"))
		End If

		' セット品の場合
		If setItem = "Y" Then
			shipSuu = CDbl(vRS("出荷合計数量"))
		Else
			shipSuu = CDbl(vRS("出荷数量"))
		End If
		'2016.05.09 GV add end

		With oJSON.data("list")
			'.Add i ,oJSON.Collection()
			'With .item(i)
			.Add "d"&CStr(vRS("受注明細番号")) ,oJSON.Collection()
			With .item("d"&CStr(vRS("受注明細番号")))
				.Add "order_detail_no", CStr(Trim(vRS("受注明細番号")))
				.Add "maker_cd", CStr(vRS("メーカーコード"))
				.Add "item_cd", CStr(vRS("商品コード"))
				.Add "iro", CStr(Trim(vRS("色")))
				.Add "kikaku",  CStr(Trim(vRS("規格")))
				.Add "item_id",  CStr(Trim(vRS("商品ID")))	'2015.03.09 GV add
				.Add "order_tanka", CDbl(Trim(vRS("受注単価")))
				.Add "order_suu", CDbl(vRS("受注数量")) 
				.Add "total_order_hikiate_suu", CDbl(vRS("受注引当合計数量"))
				.Add "total_shipping_suu", CDbl(vRS("出荷合計数量"))
				'.Add "shipping_suu", CDbl(vRS("出荷数量"))		'2015.03.09 GV add
				.Add "shipping_suu", shipSuu  '2016.05.09 GV add
				.Add "maker_name", makerName
				.Add "item_name", itemName
				.Add "web_outline", webOutline
				.Add "item_pic_small", itemPicSmall
				.Add "web_flag", CStr(vRS("Web商品フラグ"))
				.Add "set_item", setItem '2016.05.09 GV add
				.Add "shipping_yotei_text", shippingText
				.Add "invoice_no", invoiceNo
				.Add "shipping_no", shippingNo
			End With
		End With

		i = i + 1

		vRS.MoveNext
	Loop





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
