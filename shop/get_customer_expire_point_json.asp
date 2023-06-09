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
' ユーザーのポイント情報を取得
'
'
'変更履歴
'2015/01/26 GV 新規作成
'2016.04.27 GV 日付修正
'
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wCustomerNo					' 顧客番号
Dim oJSON						' JSONオブジェクト


'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
' Getパラメータ
wCustomerNo = ReplaceInput(Trim(Request("cno")))

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
Dim vPointDate
Dim vPointZan

Set oJSON = New aspJSON

'-----------------------------------------------------------
' 獲得ポイント情報の取得
'-----------------------------------------------------------
If (IsNumeric(wCustomerNo)) Then
	vSQL = createPointSql(wCustomerNo)
	'@@@@Response.Write(vSQL&"<br>")

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	// JSONオブジェクトを作成
	createJsonObject vRS

	'レコードセットを閉じる
	vRS.Close

	'レコードセットのクリア
	Set vRS = Nothing
Else
	' ポイント残
	oJSON.data.Add "point_zan" ,"0"

	' ポイント期限
	oJSON.data.Add "point_expire_date" ,""
End If


' -------------------------------------------------
' JSONデータの返却
' -------------------------------------------------
' ヘッダ出力
Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
Response.AddHeader "Pragma", "no-cache"
' JSONデータの出力
Response.Write oJSON.JSONoutput()

End Function


'========================================================================
'
'	Function	日付けのフォーマット (YYYY年MM月DD日)
'
'========================================================================
Function formatDateYYYYMMDD(pdatDate)

Dim vDate

If IsNull(pdatDate) = True Then
	' Null は計算不能
	Exit Function
End If

If IsDate(pdatDate) = False Then
	' 日付けでなければ計算不能
	Exit Function
End If

vDate = DatePart("yyyy", pdatDate) & "年"

If DatePart("m", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("m", pdatDate)
Else
	vDate = vDate & DatePart("m", pdatDate)
End If

vDate = vDate & "月"

If DatePart("d", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("d", pdatDate)
Else
	vDate = vDate & DatePart("d", pdatDate)
End If

vDate = vDate & "日"

formatDateYYYYMMDD = vDate

End Function

'========================================================================
'
'	Function	ポイント情報の取得SQL
'
'========================================================================
Function createPointSql(customerNo)
	Dim vSQL

	vSQL = ""
	vSQL = "SELECT sum(ポイント残) AS point_zan "
	vSQL = vSQL & " , min(ポイント期限) AS point_expire_date "
	vSQL = vSQL & " FROM " & gLinkServer & "ポイント明細 WITH (NOLOCK) "
	vSQL = vSQL & " WHERE "
	vSQL = vSQL & " (ポイント期限 = "
	vSQL = vSQL & "  (SELECT min(ポイント期限) FROM " & gLinkServer & "ポイント明細 WITH (NOLOCK)"
	vSQL = vSQL & "   WHERE 顧客番号 = " & customerNo
	vSQL = vSQL & "     AND ポイント日付 Is Not Null "
	vSQL = vSQL & "     AND ポイント残 Is Not Null "
	vSQL = vSQL & "     AND ポイント残 <> 0 "
	vSQL = vSQL & "     AND (ポイント期限 Is Null "
'	vSQL = vSQL & "      OR ポイント期限 >= CONVERT(datetime, '" & Now() & "')))) " ' 2016.04.27 GV mod
	vSQL = vSQL & "      OR ポイント期限 >= CONVERT(datetime, '" & Date() & "')))) " '2016.04.27 GV add
	vSQL = vSQL & " AND 顧客番号 = " & customerNo
	vSQL = vSQL & " AND ポイント日付 Is Not Null "
	vSQL = vSQL & " AND ポイント残 Is Not Null "
	vSQL = vSQL & " AND ポイント残 <> 0 "

	createPointSql = vSQL
End Function

'========================================================================
'
'	Function	DBから取得したデータからオブジェクトを生成
'
'========================================================================
Function createJsonObject(vRS)
	Dim pointZan
	Dim pointExpireDate

	pointZan = 0
	pointExpireDate = ""

	If vRS.EOF = False Then
		' ポイント残
		If (IsNull(vRS("point_zan"))) Then
			pointZan = 0
		Else
			pointZan = CStr(Trim(vRS("point_zan")))
		End If

		' ポイント期限
		If (IsNull(vRS("point_expire_date"))) Then
			pointExpireDate = ""
		Else
			pointExpireDate = CStr(Trim(vRS("point_expire_date")))
		End If
	End If


	' ポイント残
	oJSON.data.Add "point_zan" ,pointZan

	' ポイント期限
	oJSON.data.Add "point_expire_date" ,pointExpireDate
End Function
'========================================================================
%>
