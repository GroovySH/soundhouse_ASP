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
'	[受注]テーブルの「領収書」関連カラムの値を更新する。
'
'
'変更履歴
'2020.02.05 GV 新規作成
'2020.11.07 GV CStr関数の修正。(#2589)
'
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wFlg						' 実行フラグ
Dim wCustomerNo					' 顧客番号
Dim wOrderNo					' 受注番号
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
Dim vRS1
Dim vRS2
Dim vCustomerNo
Dim vOrderNo
Dim vReceiptAmount
Dim vReceiptFlag
Dim vReceiptNo
Dim vReceiptDate
Dim vReceiptName
Dim vReceiptNote
Dim vModified
Dim vModifyTantouCd

vReceiptAmount = "0"
vReceiptFlag = ""
vReceiptNo = "-1"
vReceiptDate = null
vReceiptName = ""
vReceiptNote = ""
vModified = Now()
vModifyTantouCd = "Internet"

Set oJSON = New aspJSON


' 入力値が正常の場合
If (wFlg = True) Then
	vSQL = ""
	vSQL = vSQL & "SELECT DISTINCT "
	vSQL = vSQL & "  T1.顧客番号 "
	vSQL = vSQL & " ,c.顧客名 "
	vSQL = vSQL & " ,T1.受注番号 "
	vSQL = vSQL & " ,T1.支払方法 "
	vSQL = vSQL & " ,(CASE WHEN T1.支払方法 = '現金' AND T1.受注形態 = '来店' AND T1.入金合計金額 = 0 "
	vSQL = vSQL & "          THEN T1.合計金額 "
	vSQL = vSQL & "        WHEN T1.支払方法 = '現金' "
	vSQL = vSQL & "          THEN T1.入金合計金額 "
	vSQL = vSQL & "        WHEN T1.支払方法 = '銀行振込' "
	vSQL = vSQL & "          THEN T1.入金合計金額 "
' 2020.11.07 GV add start
	vSQL = vSQL & "        WHEN T1.支払方法 = 'コンビニ支払' "
	vSQL = vSQL & "          THEN T1.入金合計金額 "
' 2020.11.07 GV add end
	vSQL = vSQL & "        WHEN T1.支払方法 = 'クレジットカード' "
	vSQL = vSQL & "          THEN T1.合計金額 "
	vSQL = vSQL & "        WHEN T1.支払方法 = 'ローン(頭金あり)' "
	vSQL = vSQL & "          THEN "
	vSQL = vSQL & "            (SELECT ol.ローン頭金入金額 "
	vSQL = vSQL & "               FROM 受注_ローン情報 ol WITH (NOLOCK) "
	vSQL = vSQL & "              WHERE ol.受注番号 = T1.受注番号) "
	vSQL = vSQL & "        END) 領収書金額 "
	vSQL = vSQL & " ,T1.領収書発行フラグ "
	vSQL = vSQL & " ,T1.領収書番号 "
	vSQL = vSQL & " ,T1.領収書発行日 "
	vSQL = vSQL & " ,T1.領収書宛先 "
	vSQL = vSQL & " ,T1.領収書但し書き "
	vSQL = vSQL & " ,T1.最終更新日 "
	vSQL = vSQL & " ,T1.最終更新者コード "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "  受注 T1 WITH (NOLOCK) "
	vSQL = vSQL & "INNER JOIN 顧客 c WITH (NOLOCK) "
	vSQL = vSQL & "   ON c.顧客番号 = T1.顧客番号 "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "  T1.顧客番号 = " & wCustomerNo
	vSQL = vSQL & " AND T1.受注番号 = " & wOrderNo

	Set vRS1 = Server.CreateObject("ADODB.Recordset")
	vRS1.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在する場合
	If vRS1.EOF = False Then
		'領収書金額
		'vReceiptAmount = CStr(Trim(vRS1("領収書金額"))) ' 2020.11.07 GV mod
		' 2020.11.07 GV add start
		If (IsNull(vRS1("領収書金額"))) Then
			vReceiptAmount = ""
		Else
			vReceiptAmount = CStr(Trim(vRS1("領収書金額")))
		End If
		' 2020.11.07 GV add end

		'領収書発行フラグ
		vReceiptFlag = getReceiptFlag(vRS1("支払方法"), wOrderNo)
		If vReceiptFlag <> "" Then
			'領収書番号
			vReceiptNo = getReceiptNo()
			If vReceiptNo <> "-1" Then
				'領収書発行日
				'毎度更新
'				If (IsNull(vRS1("領収書発行日"))) Then
					vReceiptDate = vModified
'				Else
'					vReceiptDate = CStr(Trim(vRS1("領収書発行日")))
'				End If

				'領収書宛先
				If ((IsNull(vRS1("領収書宛先"))) Or (Trim(vRS1("領収書宛先")) = "")) Then
					vReceiptName = CStr(Trim(vRS1("顧客名")))
					vReceiptName = Replace(vReceiptName, """", "”")
				Else
					vReceiptName = CStr(Trim(vRS1("領収書宛先")))
					vReceiptName = Replace(vReceiptName, """", "”")
				End If

				'領収書但し書き
				If ((IsNull(vRS1("領収書宛先"))) Or (Trim(vRS1("領収書但し書き")) = "")) Then
					vReceiptNote = getReceiptNote()
				Else
					vReceiptNote = CStr(Trim(vRS1("領収書但し書き")))
					vReceiptNote = Replace(vReceiptNote, """", "”")
				End If

				'更新
				vSQL = ""
				vSQL = vSQL & "SELECT "
				vSQL = vSQL & "  T1.受注番号 "
				vSQL = vSQL & " ,T1.領収書発行フラグ "
				vSQL = vSQL & " ,T1.領収書番号 "
				vSQL = vSQL & " ,T1.領収書発行日 "
				vSQL = vSQL & " ,T1.領収書宛先 "
				vSQL = vSQL & " ,T1.領収書但し書き "
				vSQL = vSQL & " ,T1.最終更新日 "
				vSQL = vSQL & " ,T1.最終更新者コード "
				vSQL = vSQL & "FROM "
				vSQL = vSQL & "  受注 T1 "
				vSQL = vSQL & "WHERE "
				vSQL = vSQL & "  T1.顧客番号 = " & wCustomerNo
				vSQL = vSQL & " AND T1.受注番号 = " & wOrderNo

				Set vRS2 = Server.CreateObject("ADODB.Recordset")
				vRS2.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

				If vRS2.EOF = False Then
					vRS2("領収書発行フラグ") = vReceiptFlag
					vRS2("領収書番号") = vReceiptNo
					vRS2("領収書発行日") = vReceiptDate
					'vRS2("領収書宛先") = vReceiptName 更新しない
					'vRS2("領収書但し書き") = vReceiptNote 更新しない
					vRS2("最終更新日") = vModified
					vRS2("最終更新者コード") = vModifyTantouCd
					vRS2.Update
				End If

				'レコードセットを閉じる
				vRS2.Close

				'レコードセットのクリア
				Set vRS2 = Nothing
			End If
		End If
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

oJSON.data.Add "ono" ,vOrderNo

'「領収書」関連カラム
oJSON.data.Add "receipt_am" ,vReceiptAmount
oJSON.data.Add "receipt_flg" ,vReceiptFlag
oJSON.data.Add "receipt_no" ,vReceiptNo
oJSON.data.Add "receipt_dt" ,vReceiptDate
oJSON.data.Add "receipt_name" ,vReceiptName
oJSON.data.Add "receipt_note" ,vReceiptNote

'「最終更新」関連カラム
oJSON.data.Add "modified" ,vModified
oJSON.data.Add "modify_tantou_cd" ,vModifyTantouCd

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
'
'	Function	領収書番号の生成
'
'========================================================================

Function getReceiptNo()

Dim vSQL
Dim vRS
Dim vReceiptNo

vReceiptNo = -1

vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  a.item_num1 "
vSQL = vSQL & "  FROM コントロールマスタ a WITH (ROWLOCK) "
vSQL = vSQL & " WHERE a.sub_system_cd = '共通'"
vSQL = vSQL & "   AND a.item_cd = '番号'"
vSQL = vSQL & "   AND a.item_sub_cd = '領収書'"

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then
	vReceiptNo = CLng(vRS("item_num1")) + 1
	vRS("item_num1") = vReceiptNo
	vRS.Update
End If

vRS.Close

getReceiptNo = CStr(Trim(vReceiptNo))

End Function

'========================================================================
'
'	Function	領収書但し書きの生成
'
'========================================================================

Function getReceiptNote()

Dim vSQL
Dim vRS
Dim vReceiptNote

vReceiptNote = "音響機器代として"

vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  a.item_char1 "
vSQL = vSQL & "  FROM コントロールマスタ a WITH (NOLOCK) "
vSQL = vSQL & " WHERE a.sub_system_cd = '領収書'"
vSQL = vSQL & "   AND a.item_cd = '但し書き'"
vSQL = vSQL & "   AND a.item_sub_cd = '1'"

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then
	vReceiptNote = CStr(Trim(vRS("item_char1")))
	vReceiptNote = Replace(vReceiptNote, """", "”")
End If

vRS.Close

getReceiptNote = vReceiptNote

End Function

'========================================================================
%>
