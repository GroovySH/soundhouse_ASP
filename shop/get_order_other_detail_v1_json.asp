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
'	Emax受注その他明細 取得API
'
'
'変更履歴
'2018.12.10 GV 新規作成
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


Dim orderNo		' 受注番号
Dim otherNo		' 受注その他明細番号
Dim otherCd		' 受注その他コード
Dim name		' 受注その他名称
Dim suu			' 受注その他数量
Dim tanka		' 受注その他単価
Dim amount		' 受注その他金額
Dim excTax		' 外税
Dim incTax		' 内税
Dim shippingNo	' 出荷番号
Dim modified	' 最終更新日
Dim modifyTantouCd
Dim modifyProcessName

Set oJSON = New aspJSON

'--- 受注その他明細データ取得
vSQL = ""
vSQL = vSQL & "SELECT TOP 1 "
vSQL = vSQL & " other.受注番号 "			' order_no
vSQL = vSQL & " ,other.受注その他明細番号 "	'other_no
vSQL = vSQL & " ,other.受注その他コード "	'other_cd
vSQL = vSQL & " ,other.受注その他名称 "		'name
vSQL = vSQL & " ,other.受注その他数量 "		'suu
vSQL = vSQL & " ,other.受注その他単価 "		'tanka
vSQL = vSQL & " ,other.受注その他金額 "		'amount
vSQL = vSQL & " ,other.外税 "				'exc_tax
vSQL = vSQL & " ,other.内税 "				'inc_tax
vSQL = vSQL & " ,other.出荷番号 "			'shipping_no
vSQL = vSQL & " ,other.最終更新日 "			'modified
vSQL = vSQL & " ,other.最終更新者コード "	'modify_tantou_cd
vSQL = vSQL & " ,other.最終更新処理名 "		'modify_process_name
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注その他明細     other WITH (NOLOCK) "

vSQL = vSQL & " WHERE "
vSQL = vSQL & "   other.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "  AND "
vSQL = vSQL & "   other.受注その他コード = 'COUPON' "
'vSQL = vSQL & "   other.受注その他コード = 's006' " ' test !!!

vSQL = vSQL & " ORDER BY "
vSQL = vSQL & "        other.受注その他明細番号 ASC "


'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' リスト追加
	oJSON.data.Add "data" ,oJSON.Collection()

	' --------------------
	' 受注番号
	orderNo = CStr(Trim(vRS("受注番号")))

	' 受注その他明細番号
	otherNo = CStr(Trim(vRS("受注その他明細番号")))

	'受注その他コード
	If (IsNull(vRS("受注その他コード"))) Then
		otherCd = ""
	Else
		otherCd = CStr(Trim(vRS("受注その他コード")))
	End If

	'受注その他名称
	If (IsNull(vRS("受注その他名称"))) Then
		name = ""
	Else
		name = CStr(Trim(vRS("受注その他名称")))
	End If

	'受注その他数量
	If (IsNull(vRS("受注その他数量"))) Then
		suu = ""
	Else
		suu = CDbl(vRS("受注その他数量")) 
	End If

	'受注その他単価
	If (IsNull(vRS("受注その他単価"))) Then
		tanka = 0
	Else
		tanka = CDbl(vRS("受注その他単価"))
	End If

	'受注その他金額
	If (IsNull(vRS("受注その他金額"))) Then
		amount = 0
	Else
		amount = CDbl(vRS("受注その他金額"))
	End If

	'外税
	If (IsNull(vRS("外税"))) Then
		excTax = 0
	Else
		excTax = CDbl(vRS("外税"))
	End If

	'内税
	If (IsNull(vRS("内税"))) Then
		incTax = 0
	Else
		incTax = CDbl(vRS("内税"))
	End If

	'出荷番号
	If (IsNull(vRS("出荷番号"))) Then
		shippingNo = ""
	Else
		shippingNo = CDbl(vRS("出荷番号"))
	End If

	' 最終更新日
	If (IsNull(vRS("最終更新日"))) Then
		modified = ""
	Else
		modified = CStr(Trim(vRS("最終更新日")))
	End If

	' 最終更新者コード
	If (IsNull(vRS("最終更新者コード"))) Then
		modifyTantouCd = ""
	Else
		modifyTantouCd = CStr(Trim(vRS("最終更新者コード")))
	End If

	' 最終更新処理名
	If (IsNull(vRS("最終更新処理名"))) Then
		modifyProcessName = ""
	Else
		modifyProcessName = CStr(Trim(vRS("最終更新処理名")))
	End If

	With oJSON.data("data")
		.Add "o_no", orderNo
		.Add "other_no", otherNo
		.Add "other_cd",  otherCd
		.Add "name", name
		.Add "suu", suu
		.Add "tanka", tanka
		.Add "am", amount
		.Add "ext_tax", excTax
		.Add "inc_tax", incTax
		.Add "ship_no", shippingNo
		.Add "modified", modified
		.Add "mod_tantou", modifyTantouCd
		.Add "mod_proc", modifyProcessName
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
Response.AddHeader "X-Content-Type-Options", "nosniff"

' JSONデータの出力
Response.Write oJSON.JSONoutput()

End Function
%>
