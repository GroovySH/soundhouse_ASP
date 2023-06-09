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
'	購入履歴一覧ページ (担当者一覧)
'
'
'変更履歴
'2022.03.23 GV 新規作成。(業者向けサイト)(#3110)
'
'========================================================================
'On Error Resume Next

Dim ConnectionEmax

Dim wErrDesc
Dim wFlg						' 実行フラグ
Dim wCustomerNo					' 顧客番号
Dim wOrderHidden				' 非表示フラグ
Dim wOrderCancelled				' キャンセル注文フラグ
Dim wOrderShipping				' 未発送注文フラグ
Dim wOrderGift					' ギフト注文フラグ
Dim oJSON						' JSONオブジェクト

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

'非表示フラグ
wOrderHidden = ReplaceInput_NoCRLF(Trim(Request("hide")))
If ((IsNull(wOrderHidden) = True) Or (UCase(wOrderHidden) <> "Y")) Then
	wOrderHidden = "N"
Else
	wOrderHidden = "Y"
End If

'キャンセル注文フラグ
wOrderCancelled = ReplaceInput_NoCRLF(Trim(Request("cancelled")))
If ((IsNull(wOrderCancelled) = True) Or (UCase(wOrderCancelled) <> "Y")) Then
	wOrderCancelled = "N"
Else
	wOrderCancelled = "Y"
End If

'未発送注文フラグ
wOrderShipping = ReplaceInput_NoCRLF(Trim(Request("shipping")))
If ((IsNull(wOrderShipping) = True) Or (UCase(wOrderShipping) <> "Y")) Then
	wOrderShipping = "N"
Else
	wOrderShipping = "Y"
End If

'ギフト注文フラグ
wOrderGift = ReplaceInput_NoCRLF(Trim(Request("gift")))
If ((IsNull(wOrderGift) = True) Or (UCase(wOrderGift) <> "Y")) Then
	wOrderGift = "N"
Else
'	wOrderGift = "Y" 'TODO: ギフト注文フラグを有効にする場合、この行のコメントアウトを外す
	wOrderGift = "N" 'TODO: ギフト注文フラグを有効にする場合、この行を消す
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
Dim vRS
Dim tantouParam
Dim tantouColumn

Set oJSON = New aspJSON

' 入力値が正常の場合
If (wFlg = True) Then
	'-----------------------------------------------------------
	' 該当顧客の受注の担当者氏名一覧取り出し
	'-----------------------------------------------------------
	tantouParam  = "tantou_name"
	tantouColumn = "相手先担当者"

	vSQL = createTantouListSql(tantouParam, tantouColumn)

	'@@@@Response.Write vSQL & "<br>"

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在している場合
	If vRS.EOF = False Then
		createJsonObject vRS, tantouParam
	End If

	'レコードセットを閉じる
	vRS.Close

	'-----------------------------------------------------------
	' 該当顧客の受注の担当者e_mail一覧取り出し
	'-----------------------------------------------------------
	tantouParam  = "tantou_email"
	tantouColumn = "顧客E_mail"

	vSQL = createTantouListSql(tantouParam, tantouColumn)

	'@@@@Response.Write vSQL & "<br>"

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'レコードが存在している場合
	If vRS.EOF = False Then
		createJsonObject vRS, tantouParam
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


'========================================================================
'
'	Function	担当者一覧の取得SQL
'
'========================================================================
Function createTantouListSql(tantouParam, tantouColumn)
	Dim vSQL
	Dim orderType

	' 受注形態(カンマ区切りで指定)
	orderType = ""
	orderType = orderType & "  'E-mail'"
	orderType = orderType & " ,'FAX'"
	orderType = orderType & " ,'インターネット'"
	orderType = orderType & " ,'携帯'"
	orderType = orderType & " ,'電話'"
	orderType = orderType & " ,'郵送'"
	orderType = orderType & " ,'来店'"
	orderType = orderType & " ,'スマートフォン'"
	orderType = orderType & " ,'ギフト'"

	vSQL = ""
	vSQL = vSQL & "SELECT DISTINCT o." & tantouParam & " "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & " (SELECT DISTINCT "
	vSQL = vSQL & "   o1.顧客番号 "
	vSQL = vSQL & "  ,o1.受注番号 "
	vSQL = vSQL & "  ,o1.見積日 "
	vSQL = vSQL & "  ,o1.削除日 "
	vSQL = vSQL & "  ,ov.非表示フラグ "
	vSQL = vSQL & "  ,ISNULL(o1." & tantouColumn & ", '') AS " & tantouParam & " "
	vSQL = vSQL & "  FROM 受注 AS o1 "
	vSQL = vSQL & "    INNER JOIN 受注明細 od1 WITH (NOLOCK) "
	vSQL = vSQL & "      ON od1.受注番号 = o1.受注番号 "
	vSQL = vSQL & "     AND od1.セット品親明細番号 = 0 "
	vSQL = vSQL & "    LEFT JOIN 受注非表示リスト ov WITH (NOLOCK) "
	vSQL = vSQL & "      ON ov.受注番号 = od1.受注番号 "
	vSQL = vSQL & "     AND ov.受注明細番号 = od1.受注明細番号 "
	vSQL = vSQL & "  WHERE o1.顧客番号 = " & wCustomerNo & " "
	vSQL = vSQL & "    AND o1.受注形態 IN (" & orderType & ") "

	' 未発送注文フラグ
	If wOrderShipping = "Y" Then
		vSQL = vSQL & "    AND od1.受注数量 > od1.出荷合計数量 "
	End If

	' 非表示フラグ
	If wOrderHidden = "Y" Then
		vSQL = vSQL & "    AND ov.非表示フラグ = 'Y' "
	Else
		'ギフトモードではない
		If (wOrderGift = "N") Then
			vSQL = vSQL & "    AND ov.非表示フラグ IS NULL "
		End If
	End If

	' キャンセル注文フラグ
	If wOrderCancelled = "Y" Then
		'vSQL = vSQL & "  AND o1.削除日 IS NOT NULL "
		vSQL = vSQL & "  AND od1.Webキャンセルフラグ = 'Y' "
	Else
		If wOrderHidden = "Y" Then
		'非表示フラグがYの場合、無指定
		Else
			vSQL = vSQL & "  AND o1.削除日 IS NULL "
			vSQL = vSQL & "  AND ISNULL(od1.Webキャンセルフラグ, 'N') <> 'Y' "
		End If
	End If

	vSQL = vSQL & " ) AS o "
	vSQL = vSQL & "WHERE o." & tantouParam & " <> '' "
	vSQL = vSQL & "ORDER BY o." & tantouParam & " ASC "

	createTantouListSql = vSQL
End Function

'========================================================================
'
'	Function	DBから取得したデータからオブジェクトを生成
'
'========================================================================
Function createJsonObject(vRS, tantouParam)
	Dim j
	Dim tantouListParam

	j = 0
	tantouListParam = tantouParam & "_list"

	' リスト追加
	oJSON.data.Add tantouListParam ,oJSON.Collection()

	' レコードセットの最後までループ
	Do Until vRS.EOF
		'--- 明細行生成
		With oJSON.data(tantouListParam)
			.Add j, CStr(Trim(vRS(tantouParam)))
		End With

		j = j + 1

		' レコードセットのポインタを次の行へ移動
		vRS.MoveNext
	Loop
End Function
'========================================================================
%>
