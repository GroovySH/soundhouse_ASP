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
'	購入履歴一覧ページにおける利用したポイント情報を取得
'
'
'変更履歴
'2016/03/11 GV 新規作成。(Web注文変更キャンセル機能)
'2020.06.30 GV 欲しい物リスト対応。(#2841)
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
Dim wOrderNo					' 受注番号
Dim oJSON						' JSONオブジェクト


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
Dim i
Dim vRS
Dim vHTML
Dim vPointDate
Dim vPoint
Dim vPointZan
Dim vUseOrderNo
Dim vBeforeOrderNo
Dim vBeforeOrderDetailNo
Dim vAddFlag
Dim vTotalObtainPoint	' 合計獲得ポイント

Set oJSON = New aspJSON

' 獲得リスト追加
oJSON.data.Add "obtain" ,oJSON.Collection()
' 利用リスト追加
oJSON.data.Add "used" ,oJSON.Collection()

' イテレータ初期化
i = 0

vBeforeOrderNo = null
vBeforeOrderDetailNo = null
vAddFlag = false

vTotalObtainPoint = 0

'-----------------------------------------------------------
' 獲得ポイント情報の取得
'-----------------------------------------------------------
'vSQL = ""
'vSQL = vSQL & "SELECT "
'vSQL = vSQL & "    a.受注番号 "
'vSQL = vSQL & "  , a.受注明細番号"
'vSQL = vSQL & "      , a.受注明細枝番"
'vSQL = vSQL & "      , a.ポイント区分"
'vSQL = vSQL & "      , a.ポイント日付"
'vSQL = vSQL & "      , a.ポイント"
'vSQL = vSQL & "      , a.ポイント残"
'vSQL = vSQL & "      , a.使用受注番号"
'vSQL = vSQL & "      , a.履歴登録日"
'vSQL = vSQL & " FROM "
'vSQL = vSQL & "    " & gLinkServer & "ポイント明細履歴 a WITH (NOLOCK) "
'vSQL = vSQL & " WHERE "
'vSQL = vSQL & "  (CONVERT(VARCHAR(100), 受注番号)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100),受注明細番号)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100), 受注明細枝番)+"
'vSQL = vSQL & "CONVERT(varchar(100), 履歴登録日,121))"
'vSQL = vSQL & "   IN ("
'vSQL = vSQL & "     SELECT"
'vSQL = vSQL & "       (CONVERT(VARCHAR(100), 受注番号)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100),受注明細番号)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100), 受注明細枝番)+"
'vSQL = vSQL & "CONVERT(varchar(100), MAX(履歴登録日),121))"
'vSQL = vSQL & "     FROM"
'vSQL = vSQL & "       " & gLinkServer & "ポイント明細履歴 WITH (NOLOCK) "
'vSQL = vSQL & "     WHERE"
'vSQL = vSQL & "           顧客番号 =  " & wCustomerNo
'vSQL = vSQL & "       AND 更新区分 = 'Updated'"
'vSQL = vSQL & "       AND ポイント区分 = '獲得'"
'vSQL = vSQL & "       AND ポイント残 is not null "
'vSQL = vSQL & "       AND ポイント期限 is not null "
'vSQL = vSQL & "       AND 最終更新処理 = '出荷指示'"
'vSQL = vSQL & "       AND 受注番号 IN (" & wOrderNo & ")"
'vSQL = vSQL & "     GROUP BY"
'vSQL = vSQL & "      受注番号, 受注明細番号, 受注明細枝番"
'vSQL = vSQL & "   )"
'vSQL = vSQL & " ORDER BY"
'vSQL = vSQL & "   受注番号 ASC, 受注明細番号 ASC, 受注明細枝番 ASC"
vSQL = createPointSql(wCustomerNo, wOrderNo, "獲得")

'@@@@Response.Write(vSQL&"<br>")

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

// レコードが存在した場合、JSONオブジェクトを作成
If vRS.EOF = False Then
	createJsonObject vRS, "obtain"
End If


'レコードセットを閉じる
vRS.Close


'-----------------------------------------------------------
' 利用ポイント情報の取得
'-----------------------------------------------------------
' イテレータ初期化
i = 0

vBeforeOrderNo = null

'--- 該当顧客のポイント明細の取り出し
'vSQL = ""
'vSQL = vSQL & "SELECT "
'vSQL = vSQL & "    a.受注番号 "
'vSQL = vSQL & "  , a.受注明細番号"
'vSQL = vSQL & "      , a.受注明細枝番"
'vSQL = vSQL & "      , a.ポイント区分"
'vSQL = vSQL & "      , a.ポイント日付"
'vSQL = vSQL & "      , a.ポイント"
'vSQL = vSQL & "      , a.ポイント残"
'vSQL = vSQL & "      , a.使用受注番号"
'vSQL = vSQL & "      , a.履歴登録日"
'vSQL = vSQL & " FROM "
'vSQL = vSQL & "    " & gLinkServer & "ポイント明細履歴 a WITH (NOLOCK) "
'vSQL = vSQL & " WHERE "
'vSQL = vSQL & "  (CONVERT(VARCHAR(100), 受注番号)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100),受注明細番号)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100), 受注明細枝番)+"
'vSQL = vSQL & "CONVERT(varchar(100), 履歴登録日,121))"
'vSQL = vSQL & "   IN ("
'vSQL = vSQL & "     SELECT"
'vSQL = vSQL & "       (CONVERT(VARCHAR(100), 受注番号)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100),受注明細番号)+"
'vSQL = vSQL & "CONVERT(VARCHAR(100), 受注明細枝番)+"
'vSQL = vSQL & "CONVERT(varchar(100), MAX(履歴登録日),121))"
'vSQL = vSQL & "     FROM"
'vSQL = vSQL & "       " & gLinkServer & "ポイント明細履歴 WITH (NOLOCK) "
'vSQL = vSQL & "     WHERE"
'vSQL = vSQL & "           顧客番号 =  " & wCustomerNo
'vSQL = vSQL & "       AND 更新区分 = 'Inserted'"
'vSQL = vSQL & "       AND ポイント区分 = '利用'"
'vSQL = vSQL & "       AND ポイント残 is null "
'vSQL = vSQL & "       AND ポイント期限 is null "
'vSQL = vSQL & "       AND 最終更新処理 = '受注'"
'vSQL = vSQL & "       AND 受注番号 IN (" & wOrderNo & ")"
'vSQL = vSQL & "     GROUP BY"
'vSQL = vSQL & "      受注番号, 受注明細番号, 受注明細枝番"
'vSQL = vSQL & "   )"
'vSQL = vSQL & " ORDER BY"
'vSQL = vSQL & "   受注番号 ASC, 受注明細番号 ASC, 受注明細枝番 ASC"
vSQL = createPointSql(wCustomerNo, wOrderNo, "利用")

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

// レコードが存在した場合、JSONオブジェクトを作成
If vRS.EOF = False Then
	createJsonObject vRS, "used"
End If

'レコードセットを閉じる
vRS.Close


'レコードセットのクリア
Set vRS = Nothing

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
Function createPointSql(customerNo, orderNo, kubun)
	Dim vSQL
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "    a.受注番号 "
	vSQL = vSQL & "  , a.受注明細番号"
	vSQL = vSQL & "      , a.受注明細枝番"
	vSQL = vSQL & "      , a.ポイント区分"
	vSQL = vSQL & "      , a.ポイント日付"
	vSQL = vSQL & "      , a.ポイント"
	vSQL = vSQL & "      , a.ポイント残"
	vSQL = vSQL & "      , a.使用受注番号"
	vSQL = vSQL & "      , a.ポイント番号"
	vSQL = vSQL & " FROM "
	vSQL = vSQL & "    " & gLinkServer & "ポイント明細 a WITH (NOLOCK) "
	vSQL = vSQL & "     WHERE"
	vSQL = vSQL & "           顧客番号 =  " & customerNo
	vSQL = vSQL & "       AND ポイント区分 = '" & kubun & "'"
	vSQL = vSQL & "       AND 受注番号 IN (" & orderNo & ")"
	vSQL = vSQL & " ORDER BY"
	vSQL = vSQL & "   受注番号 ASC, 受注明細番号 ASC, 受注明細枝番 ASC"

	createPointSql = vSQL
End Function

'========================================================================
'
'	Function	DBから取得したデータからオブジェクトを生成
' JSONオブジェクトで配列を追加するには、メンバ変数のキーを数値にするが、
' 受注番号の桁では追加できない。
'
'========================================================================
Function createJsonObject(vRS, kubun)
	Dim pointDate
	Dim point
	Dim pointZan
	Dim useOrderNo
	Dim addFlag
	Dim beforeOrderNo
	Dim beforeOrderDetailNo
	Dim totalPoint
	Dim beforePointNo '2021.06.30 GV add

	beforeOrderNo = null
	beforeOrderDetailNo = null
	addFlag = false
	totalPoint = 0
	beforePointNo = null ' 2021.06.30 GV add

	' レコードセットの最後までループ
	Do Until vRS.EOF

		' ポイント日付
		If (IsNull(vRS("ポイント日付"))) Then
			pointDate = ""
		Else
			pointDate = CStr(Trim(vRS("ポイント日付")))
		End If

		' ポイント
		If (IsNull(vRS("ポイント"))) Then
			point = 0
		Else
			point = CStr(Trim(vRS("ポイント")))
		End If

		totalPoint = totalPoint + vRS("ポイント")

		' ポイント残
		If (IsNull(vRS("ポイント残"))) Then
			pointZan = 0
		Else
			pointZan = CStr(Trim(vRS("ポイント残")))
		End If

		' 使用受注番号
		If (IsNull(vRS("使用受注番号"))) Then
			useOrderNo = ""
		Else
			useOrderNo = CStr(Trim(vRS("使用受注番号")))
		End If

		' 受注番号が１つ前のループ時と違う場合
		If (IsNull(beforeOrderNo) = True) Then
			addFlag = True
		ElseIf (beforeOrderNo <> vRS("受注番号")) Then
			addFlag = True
		Else
			addFlag = false
		End If

		If (addFlag) Then
			beforeOrderNo = vRS("受注番号")

			beforeOrderDetailNo = null
			beforePointNo = null '2021.06.30 GV add

			With oJSON.data(kubun)
				.Add "o"&CStr(beforeOrderNo) ,oJSON.Collection()
			End With
		End If

		'受注明細番号が1つ前のループ時と違う場合
		If (IsNull(beforeOrderDetailNo) = True) Then
			addFlag = True
		ElseIf (beforeOrderDetailNo <> vRS("受注明細番号")) Then
			addFlag = True
		Else
			addFlag = false
		End If

		'ポイント番号が1つ前のループ時と違う場合
		If (IsNull(beforePointNo) = True) Then
			addFlag = True
		ElseIf (beforePointNo <> vRS("ポイント番号")) Then
			addFlag = True
		Else
			addFlag = false
		End If



		If (addFlag = True) Then
			beforeOrderDetailNo = vRS("受注明細番号")
			beforePointNo = vRS("ポイント番号") '2021.06.30 GV add

			With oJSON.data(kubun)
				'With .item(beforeOrderNo)
				'With .item("o"&CStr(beforeOrderNo))  '2021.06.30 GV mod
				'	.Add "d"&beforeOrderDetailNo ,oJSON.Collection()  '2021.06.30 GV mod
				With .item("o"&CStr(beforeOrderNo))
					.Add "d"&beforePointNo ,oJSON.Collection()
				End With
			End With
		End If

		' 獲得リスト追加
		With oJSON.data(kubun).item("o"&CStr(beforeOrderNo))
			'With .item("d"&beforeOrderDetailNo) '2021.06.30 GV mod
			'	.Add "sub"&vRS("受注明細枝番"), oJSON.Collection()  '2021.06.30 GV mod
			With .item("d"&beforePointNo)
				.Add "sub"&vRS("受注明細枝番"), oJSON.Collection()
			End With
		End With

		With oJSON.data(kubun).item("o"&CStr(beforeOrderNo))
			'With .item("d"&beforeOrderDetailNo) '2021.06.30 GV mod
			With .item("d"&beforePointNo)
				With .item("sub"&CStr(vRS("受注明細枝番")))
					.Add "o_no" ,CStr(Trim(vRS("受注番号")))
					.Add "od_no" ,CStr(Trim(vRS("受注明細番号")))
					.Add "od_sub_no" ,CStr(Trim(vRS("受注明細枝番")))
					.Add "kubun" ,CStr(Trim(vRS("ポイント区分")))
					.Add "pt_dt" ,formatDateYYYYMMDD(pointDate)
					.Add "pt" ,point
					.Add "pt_zan" ,pointZan
					.Add "use_o_no" ,useOrderNo
				End With
			End With
		End With

		' レコードセットのポインタを次の行へ移動
		vRS.MoveNext
	Loop

	If kubun = "obtain" Then
		oJSON.data.Add "total_obtain_pt" ,totalPoint
	ElseIf kubun = "used" Then
		oJSON.data.Add "total_used_pt" ,totalPoint
	End If

'createJsonObject = oJSON
End Function
'========================================================================
%>
