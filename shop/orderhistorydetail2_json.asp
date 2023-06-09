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
Dim makerName
Dim itemName
Dim i

Set oJSON = New aspJSON


one_time_todokesaki = ""
final_nouki_date_time = ""
receiptName = ""
receiptNote = ""
webOutline = ""
source = ""
shippingText = ""
itemPicSmall = ""
makerName = ""
itemName = ""
i = 0

'--- 未出荷データの情報取出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      b.受注明細番号 "
vSQL = vSQL & "    , b.メーカーコード "
vSQL = vSQL & "    , b.商品コード "
vSQL = vSQL & "    , b.色 "
vSQL = vSQL & "    , b.規格 "
vSQL = vSQL & "    , b.受注単価 "
vSQL = vSQL & "    , b.受注数量 "
vSQL = vSQL & "    , b.受注引当合計数量 "
vSQL = vSQL & "    , b.出荷合計数量 "
vSQL = vSQL & "    , c.メーカー名 "
vSQL = vSQL & "    , d.商品名 "
vSQL = vSQL & "    , d.商品概略Web "
vSQL = vSQL & "    , d.商品画像ファイル名_小 "
vSQL = vSQL & "    , d.Web商品フラグ "
vSQL = vSQL & "    , x.出荷予定日 "
vSQL = vSQL & "    , x.ソース "
vSQL = vSQL & "    , x.出荷予定テキスト "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注明細 b WITH (NOLOCK) "
vSQL = vSQL & "        LEFT JOIN " & gLinkServer & "受注明細出荷予定 x WITH (NOLOCK) "
vSQL = vSQL & "          ON     x.受注番号     = b.受注番号 "
vSQL = vSQL & "             AND x.受注明細番号 = b.受注明細番号 "
vSQL = vSQL & "             AND x.出荷予定連番 = 1 "
vSQL = vSQL & "             AND x.変更日       = (SELECT MAX(y.変更日) "
vSQL = vSQL & "                                   FROM   " & gLinkServer & "受注明細出荷予定 y WITH (NOLOCK) "
vSQL = vSQL & "                                   WHERE      y.受注番号     = b.受注番号 "
vSQL = vSQL & "                                          AND y.受注明細番号 = b.受注明細番号) "
vSQL = vSQL & "    , " & gLinkServer & "メーカー c WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "商品 d WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        c.メーカーコード = b.メーカーコード "
vSQL = vSQL & "    AND d.メーカーコード = b.メーカーコード "
vSQL = vSQL & "    AND d.商品コード = b.商品コード "
vSQL = vSQL & "    AND b.セット品親明細番号 = 0 "
vSQL = vSQL & "    AND b.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "    AND b.受注数量 > b.出荷合計数量 "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "      c.メーカー名 "
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
		End If

		If (IsNull(vRS("ソース"))) Then
			source = ""
		Else
			source = CStr(vRS("ソース"))
			source = Replace(Trim(source), """", "”")
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

		makerName = Replace(Trim(vRS("メーカー名")), """", "”")
		makerName = CStr(makerName)

		itemName = Replace(Trim(vRS("商品名")), """", "”")
		itemName = CStr(itemName)

		With oJSON.data("list")
			.Add i ,oJSON.Collection()
			With .item(i)
				.Add "order_detail_no", CStr(Trim(vRS("受注明細番号")))
				.Add "maker_cd", CStr(vRS("メーカーコード"))
				.Add "item_cd", CStr(vRS("商品コード"))
				.Add "iro", CStr(Trim(vRS("色")))
				.Add "kikaku",  CStr(Trim(vRS("規格")))
				.Add "order_tanka", CDbl(Trim(vRS("受注単価")))
				.Add "order_suu", CDbl(vRS("受注数量")) 
				.Add "total_order_hikiate_suu", CDbl(vRS("受注引当合計数量"))
				.Add "total_shipping_suu", CDbl(vRS("出荷合計数量"))
				.Add "maker_name", makerName
				.Add "item_name", itemName
				.Add "web_outline", webOutline
				.Add "item_pic_small", itemPicSmall
				.Add "web_flag", CStr(vRS("Web商品フラグ"))
				.Add "shipping_yotei_date", shippingDate
				.Add "source", source
				.Add "shipping_yotei_text", shippingText
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
