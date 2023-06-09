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
<!--#include file="../common/HttpsSecurity.inc"-->

<%
'========================================================================
'
'	オーダーローン情報登録
'		入力されたデーターのチェック。
'		OKなら入力されたローン情報を仮受注へ追加。
'
'変更履歴
'2011/01/28 GV(ay) 新規作成
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache

'---- Session情報
Dim wUserID
Dim wUserName
Dim wMsg

Dim wErrMsg

'---- 受け渡し情報を受取る変数
Dim loan_downpayment_fl
Dim loan_downpayment_am
Dim loan_term
Dim loan_am
Dim loan_term_payment
Dim loan_apply_fl
Dim loan_company
Dim wErrDesc   '2011/08/01 an add

'---- DB
Dim Connection

'=======================================================================
'	受け渡し情報取り出し
'=======================================================================
'---- Session変数
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session("msg")

'---- 受け渡し情報取り出し
loan_downpayment_fl = Left(ReplaceInput(Trim(Request("loan_downpayment_fl"))), 1)
loan_downpayment_am = ReplaceInput(Trim(Request("loan_downpayment_am")))
loan_apply_fl = Left(ReplaceInput(Trim(Request("loan_apply_fl"))), 1)
loan_company = Left(ReplaceInput(Trim(Request("loan_company"))), 10)
loan_term_payment = Left(ReplaceInput(Trim(Request("loan_term_payment"))), 1)
loan_term = ReplaceInput(Trim(Request("loan_term")))
loan_am = ReplaceInput(Trim(Request("loan_am")))

'---- セッション切れチェック
If wUserID = ""Then
	Response.Redirect g_HTTP
End If

Session("msg") = ""
wErrMsg = ""

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()
Call main()

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderLoanStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- エラーが無いときは注文内容確認ページ、エラーがあれば注文内容指定ページへ
If wErrMsg = "" Then
	Server.Transfer "OrderConfirm.asp"
Else
	Session("msg") = wErrMsg
	Server.Transfer "OrderLoan.asp"
End If

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection= Nothing    '2011/08/01 an add

End Function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

'---- 仮受注情報更新
Call update_order_header()

''---- 入力データーのチェック
Call validate_data()

End Function

'========================================================================
'
'	Function	仮受注情報の更新
'
'========================================================================
Function update_order_header()

Dim RSv
Dim vSQL

'---- 仮受注Recordset取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    仮受注"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

RSv("ローン頭金ありフラグ") = loan_downpayment_fl
If loan_downpayment_fl = "Y" Then
	If IsNumeric(loan_downpayment_am) = False Then
		RSv("ローン頭金") = 0
	Else
		RSv("ローン頭金") = CCur(loan_downpayment_am)
	End If
Else
	RSv("ローン頭金") = 0
End If

RSv("オンラインローン申込フラグ") = loan_apply_fl

Select Case loan_apply_fl
	Case "Y"
		RSv("ローン会社") = loan_company
		RSv("希望ローン回数") = 0
		RSv("ローン金額") = 0

	Case "N"
		RSv("ローン会社") = ""
		Select Case loan_term_payment
			Case "T"		' 希望ローン回数の場合
				RSv("希望ローン回数") = CLng(loan_term)
				RSv("ローン金額") = 0

			Case "P"		' 月額支払金額の場合
				RSv("希望ローン回数") = 0
				If IsNumeric(loan_am) = False Then
					RSv("ローン金額") = 0
				Else
					RSv("ローン金額") = CCur(loan_am)
				End If
			Case Else
				RSv("希望ローン回数") = 0
				RSv("ローン金額") = 0

		End Select

End Select

RSv("最終更新日") = Now()

RSv.Update
RSv.Close

End Function

'========================================================================
'
'	Function	入力データーのチェック
'
'========================================================================
Function validate_data()

If isNumeric(loan_term) = False Then
	loan_term = 0
End If

' 頭金あり／なし
If loan_downpayment_fl = "" Then
	wErrMsg = wErrMsg & "ローン頭金あり/なしを選択してください。<br>"
End If

' 頭金ありの場合ローン頭金のチェック
If loan_downpayment_fl = "Y" Then
	If loan_downpayment_am = "" Then
		wErrMsg = wErrMsg & "ローン頭金を入力してください。<br>"
	Else
		If isNumeric(loan_downpayment_am) = False Then
			wErrMsg = wErrMsg & "ローン頭金を数字のみで入力してください。<br>"
			loan_downpayment_am = 0
		Else
			If loan_downpayment_am = 0 Then
				wErrMsg = wErrMsg & "ローン頭金を入力してください。<br>"
			End If
		End If
	End If
End If

' オンラインローン
If loan_apply_fl = "" Then
	wErrMsg = wErrMsg & "オンラインでローンを申し込むかどうかを選択してください。<br>"
End If

' オンラインローンを申込む場合ローン会社のチェック
If loan_apply_fl = "Y" Then

	' 未使用項目をクリア
	loan_am = 0
	loan_term = 0

	Select Case loan_company
		Case ""
			wErrMsg = wErrMsg & "ローン会社を選択してください。<br>"

		Case "ジャックス"
			If loan_downpayment_fl = "Y" Then
				wErrMsg = wErrMsg & "ジャックスでのお申し込みの場合、頭金を指定することはできません。<br>頭金なしをご選択ください。<br>"
			End If

	End Select

Else

	Select Case loan_term_payment
		Case ""
			wErrMsg = wErrMsg & "ローン回数か月額支払額を選択してください。<br>"

		Case "T"		' 希望ローン回数
			' 未使用項目をクリア
			loan_am = 0

			If loan_term = 0 Then
				wErrMsg = wErrMsg & "希望ローン回数を選択してください。<br>"
				loan_term = 0
			End If

		Case "P"		' 月額支払金額
			' 未使用項目をクリア
			loan_term = 0

			If loan_am = "" Then
				wErrMsg = wErrMsg & "月額支払金額を入力してください。<br>"
			Else
				If IsNumeric(loan_am) = False Then
					wErrMsg = wErrMsg & "月額支払金額を数字のみで入力してください。<br>"
					loan_am = 0
				End If
				If loan_am = 0 Then
					wErrMsg = wErrMsg & "月額支払金額を入力してください。<br>"
				End If
			End If

	End Select

End If

'If wErrMsg <> "" Then
'	wErrMsg = "<b>以下の入力エラーを訂正して下さい。</b><br /><br />" & wErrMsg
'End If

End Function
%>
