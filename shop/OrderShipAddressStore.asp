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
'	オーダー届先情報登録
'		入力されたデーターのチェック。
'		OKなら入力された届先情報をWeb顧客住所、Web顧客住所電話番号へ追加。
'
'変更履歴
'2011/01/31 GV(ay) 新規作成
'2011/04/14 hn SessionID関連変更
'2011/08/01 an #1087 Error.aspログ出力対応
'========================================================================
On Error Resume Next
Response.Expires = -1			' Do not cache

'---- Session情報
Dim wUserID
Dim wUserName
Dim wMsg

Dim wErrMsg
Dim wErrDesc   '2011/08/01 an add

'---- 受け渡し情報を受取る変数
Dim ship_name
Dim ship_zip
Dim ship_prefecture
Dim ship_address
dim ship_telephone

'---- DB
Dim Connection

'=======================================================================
'	受け渡し情報取り出し
'=======================================================================
'---- Session変数
wUserID = Session("UserID")
wUserName = Session("userName")
wMsg = Session.contents("msg")

'---- 受け渡し情報取り出し
ship_name = Left(ReplaceInput(Trim(Request("ship_name"))), 30)
ship_zip = Left(ReplaceInput(Trim(Request("ship_zip"))), 10)
ship_prefecture = Left(ReplaceInput(Trim(Request("ship_prefecture"))), 4)
ship_address = Left(ReplaceInput(Trim(Request("ship_address"))), 40)
ship_telephone = Left(ReplaceInput(Trim(Request("ship_telephone"))), 20)

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
	wErrDesc = "OrderShipAddressStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- エラーが無いときは注文内容確認ページ、エラーがあれば注文内容指定ページへ
If wErrMsg = "" Then
	Server.Transfer "OrderInfoEnter.asp"
Else
	Session("msg") = wErrMsg
	Server.Transfer "OrderShipAddress.asp"
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

Dim vAddNo

''---- 入力データーのチェック
Call validate_data()

If wErrMsg = "" Then
	'---- Web顧客住所情報登録
	vAddNo = insert_todokesaki()

	'---- 仮受注情報登録
	Call insert_Order(vAddNo)

End If

End Function

'========================================================================
'
'	Function	届先情報の登録
'
'========================================================================
Function insert_todokesaki()

Dim vSQL
Dim RSv
Dim i
Dim vMaxNo

'---- MAX住所連番の取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    MAX(住所連番) AS MAX住所連番"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web顧客住所 WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    顧客番号 = " & wUserID

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

vMaxNo = RSv("MAX住所連番") + 1

RSv.Close

'---- insert 顧客住所
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web顧客住所"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    1 = 2"
 
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

RSv.AddNew

RSv("顧客番号") = wUserID
RSv("住所連番") = vMaxNo
RSv("住所区分") = "届先"
RSv("住所名称") = ship_name
RSv("顧客郵便番号") = ship_zip
RSv("顧客都道府県") = ship_prefecture
RSv("顧客住所") = ship_address
RSv("勤務先フラグ") = "N"
RSv("規定届先フラグ") = "N"
RSv("最終更新日") = Now()
RSv("最終更新者コード") = "Internet"

RSv.Update
RSv.Close

'---- insert 顧客住所電話番号
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web顧客住所電話番号"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    1 = 2"
 
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

RSv.AddNew

RSv("顧客番号") = wUserID
RSv("住所連番") = vMaxNo
RSv("電話連番") = 1
RSv("電話区分") = "電話"
RSv("顧客電話番号") = ship_telephone
RSv("検索用顧客電話番号") = cf_numeric_only(ship_telephone)
RSv("最終更新日") = Now()
RSv("最終更新者コード") = "Internet"

RSv.Update
RSv.Close

insert_todokesaki = vMaxNo

End function

'========================================================================
'
'	Function	仮受注情報の登録
'
'========================================================================
Function insert_Order(vAddNo)

Dim RSv
Dim vSQL

'----仮受注データ取り出し
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    仮受注"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

RSv("届先区分") = "D"
RSv("届先住所連番") = vAddNo
RSv("届先名前") = ship_name
RSv("届先郵便番号") = ship_zip
RSv("届先都道府県") = ship_prefecture
RSv("届先住所") = ship_address
RSv("届先電話番号") = ship_telephone

RSv.Update
RSv.Close

End Function

'========================================================================
'
'	Function	入力データーのチェック
'
'========================================================================
Function validate_data()

Dim vSQL
Dim RSv

Dim vTel
Dim vAddress
Dim vCnt
Dim vBanchFl
Const cNumber = "0123456789０１２３４５６７８９一二三四五六七八九十"

If ship_name = "" Then
	wErrMsg = wErrMsg & "お届け先のお名前を入力してください。<br>"
Else

	If Len(ship_name) > 30 Then
		wErrMsg = wErrMsg & "お届け先のお名前は30文字以内で入力してください。<br>"
	End If

End If

If ship_zip = "" Then
	wErrMsg = wErrMsg & "お届け先の郵便番号を入力してください。<br>"
Else
	If IsNumeric(Replace(ship_zip, "-", "")) = False Or cf_checkHankaku2(ship_zip) = False Then
		wErrMsg = wErrMsg & "お届け先の郵便番号を半角で入力してください。<br>"
	Else
		If Len(ship_zip) > 10 Then
			wErrMsg = wErrMsg & "お届け先の郵便番号は10文字以内で入力してください。<br>"
		Else
			If check_zip(ship_zip, vAddress) = False Then
				wErrMsg = wErrMsg & "お届け先の郵便番号が郵便番号辞書にありません。<br>"
			Else
				If InStr(vAddress, Trim(ship_prefecture)) <= 0  Then
					wErrMsg = wErrMsg & "入力された郵便番号と都道府県が一致しません。<br>"
				End If
			End If
		End If
	End If
End If

If ship_prefecture = "" Then
	wErrMsg = wErrMsg & "お届け先の都道府県を選択してください。<br>"
Else

	If Len(ship_prefecture) > 4 Then
		wErrMsg = wErrMsg & "お届け先の都道府県は4文字以内で入力してください。<br>"
	End If

End If

If ship_address = "" Then
	wErrMsg = wErrMsg & "お届け先の住所を入力してください。<br>"
Else

	If Len(ship_address) > 40 Then
		wErrMsg = wErrMsg & "お届け先の住所は40文字以内で入力してください。<br>"
	End If

	If Len(ship_address) > 0 Then

		vBanchFl = False

		For vCnt = 1 To Len(cNumber)

			If InStr(ship_address, Mid(cNumber, vCnt, 1)) > 0 Then
				vBanchFl = True
				Exit For
			End If

		Next

		If vBanchFl = False Then
			wErrMsg= wErrMsg & "番地を入力してください。<br>"
		End If

	End If

End If

If ship_telephone = "" Then
	wErrMsg= wErrMsg & "お届け先の電話番号を入力してください。<br>"
Else

	If IsNumeric(Replace(ship_telephone, "-", "")) = False Or cf_checkHankaku2(ship_telephone) = False Then
		wErrMsg = wErrMsg & "お届け先の電話番号は半角数字とハイフン(－)で入力してください。<br>"
	Else

		If Len(ship_telephone) > 20 Then
			wErrMsg = wErrMsg & "お届け先の電話番号は20文字以内で入力してください。<br>"
		Else
			vTel = Replace(ship_telephone, "-", "")

			If Len(vTel) = 10 Or Len(vTel) = 11 Then
			Else
				wErrMsg = wErrMsg & "入力された電話番号をご確認ください。<br>"
			End If

		End If

	End If

End If

'---- 同一住所があるかどうかチェック
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    住所連番"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    Web顧客住所 WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    顧客番号 = " & wUserID
vSQL = vSQL & "    AND 住所名称 = '" & ship_name & "'"
vSQL = vSQL & "    AND 顧客郵便番号 = '" & ship_zip & "'"
vSQL = vSQL & "    AND 顧客都道府県 = '" & ship_prefecture & "'"
vSQL = vSQL & "    AND 顧客住所 = '" & ship_address & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = False Then				'同一住所あり
	wErrMsg = wErrMsg & "同一住所が既に登録されています。<br>"
	Exit Function
End If

RSv.Close

'If wErrMsg <> "" Then
'	wErrMsg = "<b>以下の入力エラーを訂正して下さい。</b><br /><br />" & wErrMsg
'End If

End Function

'========================================================================
'
'	Function	郵便番号辞書検索
'
'========================================================================
Function check_zip(pZip, pAddress)

Dim vSQL
Dim RSv

'---- 郵便番号辞書検索
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    都道府県名漢字"
vSQL = vSQL & "  , 市区町村名漢字"
vSQL = vSQL & "  , 町域名漢字"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    郵便番号辞書 WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    郵便番号 = '" & Replace(pZip, "-", "") & "'"
	  
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If RSv.EOF = False Then
	check_zip = True
	pAddress = Trim(RSv("都道府県名漢字")) & Trim(RSv("市区町村名漢字"))
Else
	check_zip = False
	pAddress = ""
End If

RSv.Close

End Function
%>
