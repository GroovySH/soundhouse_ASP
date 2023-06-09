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
'	Emax受注明細　取得API
'
'
'変更履歴
'2016/03/29 GV 新規作成
'2016.09.06 GV キャンセル時の引当数戻し処理の改修対応。
'2020.02.28 GV クーポンとポイント適用条件(引当可能数チェック)対応。
'2020.11.20 GV 注文変更時の配送日指定チェック改修。(#2602)
'
'========================================================================
'On Error Resume Next

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
Dim i
Dim j

Dim iro
Dim kikaku
Dim makerName
Dim itemName
Dim kosuuGenteiTankaFlg
Dim bItemFlg
Dim estimateHikiateSuu ' 2016.09.06 GV add
Dim hikiateSuuAtOrder  ' 2020.02.28 GV add

Set oJSON = New aspJSON
i = 0
j = 0

'--- 明細部分の情報取出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  od.受注番号 "
vSQL = vSQL & " ,od.受注明細番号 "
vSQL = vSQL & " ,od.メーカーコード "
vSQL = vSQL & " ,od.商品コード "
vSQL = vSQL & " ,od.色 "
vSQL = vSQL & " ,od.規格 "
vSQL = vSQL & " ,mk.メーカー名 "
vSQL = vSQL & " ,od.商品名 "
vSQL = vSQL & " ,od.受注数量 "
vSQL = vSQL & " ,od.受注単価 "
vSQL = vSQL & " ,od.受注金額 "
vSQL = vSQL & " ,od.個数限定単価フラグ "
vSQL = vSQL & " ,od.B品フラグ "
vSQL = vSQL & " ,od.見積引当合計数量 " ' 2016.09.06 GV add
vSQL = vSQL & " ,od.受注時引当可能在庫数量 " ' 2016.09.06 GV add

vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注 o WITH (NOLOCK) "
vSQL = vSQL & "INNER JOIN " & gLinkServer & "受注明細 od WITH (NOLOCK) "
vSQL = vSQL & "  ON od.受注番号 =o.受注番号 "
vSQL = vSQL & "  AND od.セット品親明細番号 = 0 "
vSQL = vSQL & "INNER JOIN " & gLinkServer & "メーカー mk WITH (NOLOCK) "
vSQL = vSQL & "  ON mk.メーカーコード = od.メーカーコード "

vSQL = vSQL & "WHERE "
vSQL = vSQL & "      o.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "  AND o.顧客番号 = " & wUserID & " "
vSQL = vSQL & "  AND od.受注数量 > 0 "
vSQL = vSQL & "  AND od.受注金額 > 0 "

vSQL = vSQL & " ORDER BY "
vSQL = vSQL & "        od.受注明細番号 ASC "


'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' リスト追加
	oJSON.data.Add "list" ,oJSON.Collection()

	' --------------------
	For i = 0 To (vRS.RecordCount - 1)
		'色
		iro = Replace(Trim(vRS("色")), """", "”")
		iro = CStr(iro)

		'規格
		kikaku = Replace(Trim(vRS("規格")), """", "”")
		kikaku = CStr(kikaku)

		'メーカー名
		makerName = Replace(Trim(vRS("メーカー名")), """", "”")
		makerName = CStr(makerName)

		'商品名
		'itemName = Replace(Trim(vRS("商品名")), """", "”")
		'itemName = CStr(itemName)

		itemName = Replace(Trim(vRS("商品名")), """", "”")
		itemName = CStr(itemName)

		If (IsNull(vRS("個数限定単価フラグ"))) Then
			kosuuGenteiTankaFlg = ""
		Else
			kosuuGenteiTankaFlg = CStr(Trim(vRS("個数限定単価フラグ")))
		End If

		If (IsNull(vRS("B品フラグ"))) Then
			bItemFlg = ""
		Else
			bItemFlg = CStr(Trim(vRS("B品フラグ")))
		End If

		If (IsNull(vRS("受注時引当可能在庫数量"))) Then
			hikiateSuuAtOrder = 0
		Else
			hikiateSuuAtOrder = CDbl(vRS("受注時引当可能在庫数量"))
		End If

		'2020.11.20 GV add
		If (IsNull(vRS("見積引当合計数量"))) Then
			estimateHikiateSuu = 0
		Else
			estimateHikiateSuu = CDbl(vRS("見積引当合計数量"))
		End If


		'--- 明細行生成
		With oJSON.data("list")
			.Add j ,oJSON.Collection()
			With .item(j)
				.Add "o_no" ,CStr(Trim(vRS("受注番号")))
				.Add "od_no" ,CStr(Trim(vRS("受注明細番号")))
				.Add "m_cd" ,CStr(Trim(vRS("メーカーコード")))
				.Add "i_cd" ,CStr(Trim(vRS("商品コード")))
				.Add "iro" ,iro
				.Add "kikaku" ,kikaku
				.Add "m_nm" ,makerName
				.Add "i_nm" ,itemName
				.Add "i_suu", CDbl(vRS("受注数量")) 
				.Add "i_tanka", CDbl(Trim(vRS("受注単価")))
				.Add "i_am", CDbl(Trim(vRS("受注金額")))
				.Add "kosuu_lmt", kosuuGenteiTankaFlg '個数限定単価フラグ
				.Add "b_item", bItemFlg 'B品フラグ
				.Add "est_hikiate_suu", estimateHikiateSuu  ' 2020.11.20 GV add
				.Add "hikiate_suu_at_order", hikiateSuuAtOrder ' 2020.04.15 GV add
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

' -------------------------------------------------
' JSONデータの返却
' -------------------------------------------------
' ヘッダ出力
Response.AddHeader "Content-Type", "application/json"
Response.AddHeader "X-Content-Type-Options", "nosniff"

' JSONデータの出力
Response.Write oJSON.JSONoutput()

End Function
%>
