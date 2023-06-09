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
'	オーダーログイン情報チェック
'
'	更新履歴
'2004/12/20 NaviLeftログインボタンからの呼び出し時は呼び出しもとへもどる。
'2005/06/06 SQLインジェクション対策
'2006/08/08 入力データチェック強化
'2006/09/18 LoginFl追加　ログイン情報保持
'2007/03/22 パスワードにTrim追加　（スペース入力時の対処)
'2007/08/07 Web不掲載フラグ NOT= Y　のデータのみ採用
'2008/04/01 パスワードハッシュ
'2008/04/05 パスワードリセットを1度もしていない人は、強制的にリセットページへ
'2008/04/14 ログイン履歴作成、3回続けて失敗したらロック
'2008/05/13 ログイン後HTTPS用セッションIDをセット
'2008/05/14 HTTPSチェック対応
'2009/04/30 エラー時にerror.aspへ移動
'2009/05/29 パスワードハッシュをSHA-256に変更
'2010/03/17 hn FeedBack.aspへの呼び出しの戻り追加
'2010/07/28 an ログインログにSessionID, 顧客番号追加
'2010/07/30 st RtnURLがある場合はそのまま呼び出し元へリダイレクト
'2011/02/21 hn RtnURLのチェック強化（PCIDSS)
'2011/04/14 hn SessionID 関連チェック
'2011/04/20 an #843 ログイン時、Emailの代わりにユーザーIDを使用/ログイン履歴を共通関数化
'2011/06/13 hn 同一ユーザーIDが有ればエラー
'2011/08/01 an #1087 Error.aspログ出力対応
'2011/10/01 an #722 ログイン時に顧客名をCookieにセット
'2011/12/19 hn DC対応
'2011/12/25 hn マイページ対応  member.asp　→ mypage.asp
'2012/01/17 hn cookieにDomain属性追加
'2012/02/15 GV Cookie の LIFL を ULIFL にキー名変更
'2012/03/07 GV #1234 エラーログ出力を共通プロシージャを用いて出力するよう変更
'2012/03/08 GV #1234 エラーチェック等 出力メッセージの文言を変更
'2012/03/26 GV #1254 ユーザーIDとパスワードが同一のお客様はログイン出来ないように変更
'2012/03/26 GV #1254 ユーザーIDに一致する顧客情報あり、パスワード違いの際のメッセージ変更
'2012/03/26 GV 過去の不要なコメントアウト処理およびコメントを削除 (2011/8/1以前分)
'2012/09/07 nt ウィッシュリストからのリダイレクトを追加
'
'========================================================================
On Error Resume Next

Response.Expires = -1			' Do not cache

Dim userID
Dim MemberID
Dim LoginFl
Dim LoginCount
Dim member_email
Dim member_password
Dim called_from
Dim RtnURL

Dim Connection
Dim RS_customer
Dim RS_order_header

DIm wPasswordResetFl

Dim w_sql
Dim w_msg
Dim w_html

Dim w_userID
Dim w_userName
Dim wMemberEmail
Dim wErrDesc   '2011/08/01 an add

' CAPICOM's hash algorithm constants.
Const CAPICOM_HASH_ALGORITHM_SHA1      = 0
Const CAPICOM_HASH_ALGORITHM_MD2       = 1
Const CAPICOM_HASH_ALGORITHM_MD4       = 2
Const CAPICOM_HASH_ALGORITHM_MD5       = 3
Const CAPICOM_HASH_ALGORITHM_SHA256    = 4
Const CAPICOM_HASH_ALGORITHM_SHA384    = 5
Const CAPICOM_HASH_ALGORITHM_SHA512    = 6

'=======================================================================

userID = Session("userID")
LoginFl = Session("LoginFl")
LoginCount = Session("LoginCount")

If IsNumeric(LoginCount) = False Then
	LoginCount = 0
Else
	LoginCount = CLng(LoginCount)
End If

'---- 入力データーの取り出し
MemberID = ReplaceInput_NoCRLF(Trim(Request("MemberID")))
member_password = ReplaceInput(Trim(Request("member_password")))
called_from = ReplaceInput(Request("called_from"))
RtnURL = replace(ReplaceInput(Request("RtnURL")), "＆", "&")

'---- RtnURLが不正な場合はエラー @@@暫定対応 アカマイ結果待ち@@@ 2011/02/28 hn add
If RtnURL <> "" Then
	If InStr(LCase(RtnURL), LCase(g_HTTP)) <> 1 _
	And InStr(LCase(RtnURL), LCase(g_HTTPS)) <> 1 _
	And InStr(LCase(RtnURL), "http://hotplaza.soundhouse.co.jp") <> 1 _
	And InStr(LCase(RtnURL), "http://guide.soundhouse.co.jp") <> 1 Then
' 2012/03/07 GV Add Start
		' エラーログ出力
		Call fwriteErrorLog("引数 RtnURL が不正 (RtnURL=" & RtnURL & ")")
' 2012/03/07 GV Add End
		Response.Redirect g_HTTP & "shop/Error.asp"
	End If
End If

'---- メイン処理
Session("msg") = ""
w_msg = ""

If userID = "" Or LoginFl <> "Y" Then
	Call connect_db()
	Call main()

	'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
	If Err.Description <> "" Then
' 2012/03/07 GV Mod Start
'		wErrDesc = "LoginCheck.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		wErrDesc = Err.Description
		wErrDesc = Replace(wErrDesc, vbNewLine, " ")
		' エラーログ出力
		Call fwriteErrorLog("ログインチェック処理でエラー " & wErrDesc)

		wErrDesc = "LoginCheck.asp" & " " & wErrDesc
' 2012/03/07 GV Mod End

		Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
	End If                                           '2011/08/01 an add e

	Call close_db()
End If

If Err.Description <> "" Then
' 2012/03/07 GV Add Start
	' エラーログ出力
	wErrDesc = Err.Description
	Call fwriteErrorLog(wErrDesc)
' 2012/03/07 GV Add End
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

If w_msg = "" Then

	If wPasswordResetFl = "Y" Then
		Response.Redirect g_HTTPS & "member/MemberPasswordResetRequest.asp?member_email=" & wMemberEmail
	Else
		If RtnURL <> "" Then
			Response.Redirect RtnURL
		Else
			Select Case called_from
			Case "order"
				Response.Redirect g_HTTPS & "shop/OrderInfoEnter.asp"
			Case "catalog"
				Response.Redirect g_HTTPS & "shop/CatalogRequest.asp"
			Case "present"
				Response.Redirect g_HTTPS & "shop/PresentOubo.asp"
			Case "feedback"
				Response.Redirect g_HTTPS & "shop/FeedBack.asp"
			Case "top"
				Response.Redirect g_HTTPS & "member/Mypage.asp?called_from=" & called_from
			'2012/09/07 nt add Start
			'---- ウィッシュリストからのリダイレクトを追加
			Case "wishlist"
				Response.Redirect g_HTTPS & "shop/WishList.asp?called_from=" & called_from
			'2012/09/07 nt add End
			Case "navi"
				Response.Redirect g_HTTP			'Topへ
			Case Else
				Response.Redirect g_HTTPS & "member/Mypage.asp?called_from=" & called_from
			End Select
		End If
	End If

Else

	If w_msg <> "NoData" Then
		Session("msg") = w_msg
	End If

	Response.Redirect g_HTTPS & "shop/Login.asp?called_from=" & called_from & "&RtnURL=" & RtnURL		'Login error at First Login

End If

'=======================================================================

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End Function

'========================================================================
'
'	Function	Main
'
'========================================================================
'
Function main()

Dim HashedData
Dim vCustNo
Dim vHashPassword										' 2012/03/26 GV Add

Const LOGIN_FLAG_KEY = "ULIFL"							' 2012/02/15 GV Add

If MemberID = "" And member_password = "" Then
	w_msg = "NoData"
	Exit Function
End If

' 2012/03/08 GV Mod Start
'If MemberID = "" or member_password = "" Then
'	w_msg = "入力されたユーザーIDまたは、パスワードが正しくありません。"
'	Exit Function
'End If
If Len(MemberID) <= 0 Then
	w_msg = "ユーザーIDを入力して下さい。"
	Exit Function
End If
If Len(member_password) <= 0 Then
	w_msg = "パスワードを入力して下さい。"
	Exit Function
End If
' 2012/03/08 GV Mod End

'---- 顧客情報チェック
w_sql = ""
w_sql = w_sql & "SELECT 顧客番号"
w_sql = w_sql & "       , 顧客名"
w_sql = w_sql & "       , パスワード2"
w_sql = w_sql & "       , Web不掲載フラグ"
w_sql = w_sql & "       , リセットトークン"
w_sql = w_sql & "       , リセットトークン登録日"
w_sql = w_sql & "       , パスワードロック日"
w_sql = w_sql & "       , ハッシュアルゴリズム"
w_sql = w_sql & "       , 顧客E_mail1"
w_sql = w_sql & "  FROM Web顧客"
w_sql = w_sql & " WHERE ユーザーID = '" & MemberID & "'"
w_sql = w_sql & "   AND Web不掲載フラグ != 'Y'"
w_sql = w_sql & " ORDER BY 顧客番号 DESC"

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open w_sql, Connection, adOpenStatic, adLockOptimistic

If RS_customer.EOF = True Then
' 2012/03/08 GV Mod Start
'	w_msg = "入力されたユーザーIDまたは、パスワードが正しくありません。"
	w_msg = "入力されたユーザーIDのお客様情報がみつかりませんでした。" _
	      & "<br>ご登録されたユーザーIDをお忘れの場合は、下のフォームよりご確認頂けます。"
' 2012/03/08 GV Mod End
	Call fInsertLoginHistory(MemberID, "ログイン", "失敗", gSessionID, "")
Else
	If IsNULL(RS_customer("パスワードロック日")) = False Then
' 2012/03/08 GV Mod Start
'		w_msg = "この会員情報はロックされています｡<br>パスワードリセットを行ってください。"
		w_msg = "このお客様情報はロックされています。" _
		      & "<br>ロックを解除するには、下のフォームよりパスワードの再設定を行ってください。"
' 2012/03/08 GV Mod End
	Else

		'---- 重複があればエラー
		If RS_customer.RecordCount > 1 Then
			w_msg = "同じユーザーIDが登録されています。<br>申し訳ありませんが、新たに会員登録をお願いいたします。"
		Else

			'--- オブジェクト作成
			Set HashedData = CreateObject("CAPICOM.HashedData")
			'--- アルゴリズムにSHA1を指定
			If RS_customer("ハッシュアルゴリズム") = "SHA1" or RS_customer("ハッシュアルゴリズム") = "" Then
				HashedData.Algorithm = CAPICOM_HASH_ALGORITHM_SHA1
			End If
			'--- アルゴリズムにSHA-256 を指定
			If RS_customer("ハッシュアルゴリズム") = "SHA-256" Then
				HashedData.Algorithm = CAPICOM_HASH_ALGORITHM_SHA256
			End If
			'--- ハッシュ値を計算
			HashedData.Hash member_password

' 2012/03/26 GV Add Start
			vHashPassword = HashedData.Value

			Set HashedData = Nothing
' 2012/03/26 GV Add End

' 2012/03/26 GV Mod Start
'			If RS_customer("パスワード2") <> HashedData.Value Then
			If RS_customer("パスワード2") <> vHashPassword Then
' 2012/03/26 GV Mod End
' 2012/03/08 GV Mod Start
'				w_msg = "入力されたユーザーIDまたは、パスワードが正しくありません。"
' 2012/03/26 GV Mod Start
'				w_msg = "入力されたユーザーID、または、パスワードのお客様情報がみつかりませんでした。" _
'				      & "<br>ご登録されたユーザーID、またはパスワードをお忘れの場合は、下のフォームよりご確認頂けます。"
				w_msg = "入力されたパスワードのお客様情報がみつかりませんでした。" _
				      & "<br>パスワードをお忘れの場合は、下のフォームよりご確認頂けます。"
' 2012/03/26 GV Mod End
' 2012/03/08 GV Mod End
			Else
				If RS_customer("Web不掲載フラグ") = "Y" Then
					w_msg = "この会員情報は削除されています｡<br>再登録などのお問合せは<a href='" & g_HTTP & "shop/Inquiry.asp'>こちら</a>"
				Else

' 2012/03/26 GV Add Start
					If LCase(MemberID) = LCase(member_password) Then

						' ログインIDとパスワードが同一の場合、ログイン不可(パスワード変更を促す)
						w_msg = "入力されたユーザーIDとパスワードが同じです。" _
						      & "<br>変更をお願いいたします。"

					Else
' 2012/03/26 GV Add End

						If (IsNull(RS_customer("リセットトークン登録日")) = True) Or _
						   (IsNull(RS_customer("リセットトークン登録日")) = False And RS_customer("リセットトークン") <> "") Then
							wPasswordResetFl = "Y"
							wMemberEmail = RS_customer("顧客E_mail1")
						Else
							Session("userID") = RS_customer("顧客番号")		'OKのとき顧客番号をセット
							Session("userName") = RS_customer("顧客名")		'OKのとき顧客名をセット
							Session("LoginFl") = "Y"		'ログインした時はY
							Session("LoginCount") = 0		'ログインした時は0

							'---- Cookieに顧客名セット										'2011/10/01 an add s
							Response.Cookies("CustName") = RS_customer("顧客名")			'2011/10/01 an add e
							Response.Cookies("CustName").Domain = gCookieDomain				'2012/01/17 hn add

							'---- Cookieにログインフラグセット								'2011/12/19 hn add
' 2012/02/15 GV Mod Start
'							Response.Cookies("LIfl") = "Y"
'							Response.Cookies("LIfl").Domain = gCookieDomain					'2012/01/17 hn add
							Response.Cookies(LOGIN_FLAG_KEY) = "Y"
							Response.Cookies(LOGIN_FLAG_KEY).Domain = gCookieDomain
' 2012/02/15 GV Mod End

							'---- セッションデータに顧客番号セット       '2011/12/19 hn add
							vCustNo = RS_customer("顧客番号")
							Call fSetSessionData(gSessionID, "顧客番号", vCustNo)

							'---- HTTPS用セッションIDセット
							'Call SetSSID()                  '2011/10/01 an del
						End If

' 2012/03/26 GV Add Start
					End If
' 2012/03/26 GV Add End

				End If
			End If
		End If
	End If

	If w_msg <> "" Then

		Session("LoginCount") = LoginCount + 1

		If Session("LoginCount") >= 5 Then

			RS_customer("パスワードロック日") = Now()
			RS_customer.update

			' ログイン履歴作成
			Call fInsertLoginHistory(MemberID, "パスワードロック", "--", gSessionID, RS_customer("顧客番号"))

		Else
			' ログイン履歴作成
			Call fInsertLoginHistory(MemberID, "ログイン", "失敗", gSessionID, RS_customer("顧客番号"))
		End If

	Else

		' ログイン履歴作成
		Call fInsertLoginHistory(MemberID, "ログイン", "成功", gSessionID, RS_customer("顧客番号"))

	End If

End If

RS_customer.Close

End Function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.Close
Set Connection = Nothing    '2011/08/01 an add

End Function
%>
