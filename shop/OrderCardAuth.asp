<%@ LANGUAGE="VBScript" %>
<%
 Option Explicit
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<%
'========================================================================
'
'	カードオーダー与信確認処理
'
'		カードの与信を取りOkならorder_submitへコントロールを渡す。
'
'------------------------------------------------------------------------
'	
'		このプログラムはCardnetからのサンプルプログラムを元に作られています。
'		入力されたカードの与信チェックを行う｡
'		OKなら受注登録処理へ
'
'------------------------------------------------------------------------
'	更新履歴
'2005/04/05 カード情報を受注データから取り出すように変更
'2006/06/30 受注情報なしのときはエラー
'2009/04/30 エラー時にerror.aspへ移動
'
'========================================================================

On Error Resume Next

Dim w_sessionID
Dim userID
Dim msg

Dim card_no
Dim card_exp_dt
Dim card_exp_dt1
Dim card_exp_dt2
Dim card_holder_nm
Dim order_total_am
Dim card_order_no
Dim card_net_no
Dim card_auth_no

Dim Connection
Dim RS_order_header

Dim w_sql
Dim w_html
Dim w_msg
Dim w_next_URL

'=======================================================================

w_sessionID = Session.SessionId
userID = Request.cookies("UserID")

Session("msg") = ""
w_msg = ""

'---- execute main process
call connect_db()
call main()
call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- エラーが無いときは注文登録処理ページ、エラーがあれば確認ページへ

if w_msg = "" then
	Response.Redirect "OrderSubmit.asp"
else
	Session("msg") = w_msg
	Response.Redirect "OrderInfoEnter.asp"
end if

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

End function

'========================================================================
'
'	Function	Main カード与信確認
'
'========================================================================
'
Function main()

'---- カード情報取り出し
call get_card()

if w_msg <> "" then
	exit function
end if

'************************************** 変更する
card_order_no = 00001

'---- 与信チェック
call card_auth()

'---- 受注情報に与信確認番号をセット
if w_msg = "" then
	call update_order_header()
end if

'---- 念のためチェック
if RS_order_header("カード与信確認番号") = "" then
		w_msg = "<font color='#ff0000'>カード与信の取得が出来ませんでした。<br>別のカードで再度御注文ください。</font>"
end if

RS_order_header.close

End Function

'========================================================================
'
'	Function	カード情報取り出し
'
'========================================================================
'
Function get_card()

'---- 仮受注取り出し
w_sql = ""
w_sql = w_sql & "SELECT a.カード番号"
w_sql = w_sql & "     , a.カード有効期限"
w_sql = w_sql & "     , a.カード名義人"
w_sql = w_sql & "     , a.受注合計金額"
w_sql = w_sql & "     , a.カード与信確認番号"
w_sql = w_sql & "     , a.カードネット伝票番号"
w_sql = w_sql & "  FROM 仮受注 a"
w_sql = w_sql & " WHERE SessionID = " & w_sessionID
	  
Set RS_order_header = Server.CreateObject("ADODB.Recordset")
RS_order_header.Open w_sql, Connection, adOpenStatic, adLockOptimistic

if RS_order_header.EOF = true then
	w_msg = "<font color='#ff0000'>NoData</font>"
	exit function
end if

card_no = RS_order_header("カード番号")
card_exp_dt = RS_order_header("カード有効期限")
card_exp_dt1 = Left(card_exp_dt, 2)
card_exp_dt2 = Right(card_exp_dt, 2)
card_holder_nm = RS_order_header("カード名義人")
order_total_am = RS_order_header("受注合計金額")

End function

'========================================================================
'
'	Function	仮受注情報の更新
'
'========================================================================
'
Function update_order_header()

'---- update 仮受注
RS_order_header("カード与信確認番号") = card_auth_no
RS_order_header("カードネット伝票番号") = card_net_no

RS_order_header.update

End function

'========================================================================
'
'	Function	カード与信確認
'
'========================================================================
'
Function card_auth()

REM	/*==================================================================*/

REM -- 決済パッケージを展開したディレクトリを設定します。
REM -- (参考)	sgsv003z.exe, sgsv004z.exe, sgsv012z.exe, sgsv00001a.prm
REM --			を展開したディレクトリです。

Dim HomeDir
HomeDir = "d:\soundhouse\wwwroot\cardnet"			'本番@@@@@@@@@@@@@@@@@@@@@@@@
'''''HomeDir = "\\Emax2\Web\SH_New\cardnet"		'テスト@@@@@@@@@@@@@@@@@@@@@@@@


REM	/*==================================================================*/
REM	/* この設定はパッケージシステムの設定値です。						*/
REM	/*==================================================================*/

REM	--	ＡＳＰ連携ＤＬＬレジストリ登録名

Dim DllRegist
DllRegist = "Sgsv011z.SSLAuth"

REM	--	システムエラーが発生した場合に表示するＨＴＭＬ

Dim ErrorURL
ErrorURL = "OrderInfoEnter.asp"

REM	--	与信が成功した場合に表示するＨＴＭＬ

Dim SuccessURL
SuccessURL = "OrderSubmit.asp"

REM	--	与信が拒否された場合に表示するＨＴＭＬ

Dim FailureURL
FailureURL = "OrderInfoEnter.asp"


REM	/*==================================================================*/
REM	/*                                                                  */
REM	/*		ファイル名		：	sgsv00013a.asp		(Original name)           */
REM	/*                                                                  */
REM	/*		概要			：	カートリッジ連携実行ＡＳＰ                      */
REM	/*                                                                  */
REM	/*		作成日			： 2000/01/04                                     */
REM	/*                                                                  */
REM	/*		更新履歴                                                      */
REM	/*		  日付	  変更者				理由                                  */
REM	/*                                                                  */
REM	/*==================================================================*/

REM --
REM -- オブジェクトを実体化します。(最初に必ず必要です。)
REM --

'@@@@@@On Error Resume Next

Dim SSLAuth
Set SSLAuth = CreateObject(DllRegist)

REM --
REM -- 決済パッケージを展開したディレクトリを設定します。
REM -- (参考)
REM -- sgsv003z.exe, sgsv004z.exe, sgsv012z.exe, sgsv00001a.prmが
REM -- 存在するディレクトリです。
SSLAuth.HomeDir = HomeDir

REM --
REM -- ＥＣ決済センターへ送信するための情報を次のように設定します。
REM --

REM -- （必須）サーバＩＤを設定（株式会社日本カードネットから通知される）
REM -- <<固定値>>
SSLAuth.ServerID = "1680"

REM -- （必須）ショップＩＤを設定（株式会社日本カードネットから通知される）
REM -- <<固定値>>
SSLAuth.ShopID = "0001"

REM -- （必須）オーソリ金額を設定
SSLAuth.Amount = order_total_am

REM -- （必須）支払方法を設定
SSLAuth.PayMode = "10"			' 一括

REM -- （支払区分が分割(61)の時のみ必須）
REM -- 分割回数を設定します。分割以外の場合に設定しても構いません。
If Request("card_payment_method") = "61" Then
	SSLAuth.InstallCount = 1
End If

REM -- （必須）カード番号を設定
SSLAuth.PAN = card_no

REM -- （必須）カード有効期限を設定 ☆注意：(Month2桁+Year２桁)
SSLAuth.CardExp = card_exp_dt1 & card_exp_dt2

REM -- （必須）伝票番号を設定
REM -- <<加盟店様がカスタマイズして、伝票番号を設定してください。>>
SSLAuth.SalesSlipNo = cf_NumToChar(card_order_no, 5)

REM -- （オプション）商品番号を設定
SSLAuth.GoodsCode = "0990"

REM -- （オプション）加盟店契約カード会社コードを設定
If Request("RECV_CO_COD") <> "" Then
	SSLAuth.CardCoCode = ""
End If

REM -- （オプション）加盟店端末番号設定
If Request("MER_TERM_NUM") <> "" Then
	SSLAuth.MerchantID = ""
End If

REM -- （オプション）端末識別番号
If Request("TERM_NUM") <> "" Then
	SSLAuth.TerminalID = ""
End If

REM --
REM -- ＥＣ決済センターへデータを送信します。
REM --
SSLAuth.Send()

REM -- システム的なエラーが発生したかを調べます。
Dim SystemErrorMessage
SystemErrorMessage = ""
If Err.Description <> "" Then
	w_next_url = ErrorURL
	w_msg = "<font color='#ff0000'>" _
				& "system error申し訳ございませんが､センターシステム側で受付を停止しております。<br>" _
				& "しばらくしてから御注文ください。<br>" _
				& "Code: " & p_ErrorCode _
				& "</font>"
	Exit Function
End If

REM --
REM -- 取扱結果表示
REM --

If Err.Number <> 0 Then
	REM -- システムエラーが発生した場合
	w_next_url = ErrorURL
	call card_error(SSLAuth.ErrorCode)
Else
	REM -- システム的に正常な場合
	'If SSLAuth.ErrorCode = "   " Then
	If Trim(SSLAuth.ErrorCode) = "" Then
		if Trim(SSLAuth.AuthCode) = "" then
			REM オーソリ取得に失敗
			w_next_url = FailureURL
			w_msg = "<font color='#ff0000'" _
						& "カード与信の取得が出来ませんでした。<br>" _
						& "別のカードで再度御注文ください。<br>" _
						& "</font>"
		else
			REM オーソリ取得に成功
			card_auth_no = SSLAuth.AuthCode
			card_net_no = SSLAuth.CardNetNo
			if trim(card_auth_no) = "" then			'念のためチェック	'020924
				w_next_url = FailureURL
				w_msg = "<font color='#ff0000'" _
							& "カード与信の取得は出来ましたが、与信番号取得中にエラーが発生しました。<br>" _
							& "弊社営業までご連絡ください<br>" _
							& "</font>"
			else
				w_next_url = SuccessURL
				w_msg = ""
			end if
		end if
	else
		REM オーソリ取得に失敗
		w_next_url = FailureURL
		call card_error(SSLAuth.ErrorCode)
	End If
End If

REM --
REM -- 最後に必ずオブジェクトを解放します。
REM --
Set SSLAuth = Nothing

end function

'========================================================================
'
' カード決済 Error
'
'========================================================================

Dim ErrorCode

Dim error_input
Dim error_card
Dim error_system
Dim error_system_l
Dim error_system_h
Dim error_package

error_input = "G65,G83"

error_card = "G12,G55,G56,G60,G61,S06"

error_system = "V12,J01,J02,J10,J11,J20,J21,J22,J30,J31,J32,S01,S02,S03," _
			 & "S04,S05,S10,S12,S13,S15,S90,S99,P01,P12,P30,P31,P50,P51," _
			 & "P52,P53,P54,P55,P65,P68,P69,P70,P71,P72,P73,P74,P75,P76," _
			 & "P78,P80,P81,P83,P84,P90,E90,K01,K02,K40,K50"
			 
error_system_l = "C00"
error_system_h = "C99"

error_package = "V01,V02,V03,V10,V11,V14,V15,V99"

'=======================================================================

'========================================================================
'
'	Function	カードエラーメッセージ作成
'
'========================================================================
'
Function card_error(p_ErrorCode)

'---- set error message
'---- 入力エラー
if InStr(error_input, p_ErrorCode) > 0 then	'input error
	w_msg = "<font color='#ff0000'>" _
				& "カード番号または有効期限の入力に誤りがありました。<br>" _
				& "入力内容を確認してから再度御注文ください。<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "よくあるご質問は<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>こちら</a>" _
				& "</font>"
	exit function
end if

'---- カードエラー
if InStr(error_card, p_ErrorCode) > 0 then	'card error
	w_msg = "<font color='#ff0000'>" _
				& "申し訳ございませんが､御指定のカードでは御注文できません。<br>" _
				& "別のカードまたは､別のお支払方法で御注文願います。<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "よくあるご質問は<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>こちら</a>" _
				& "</font>"
	exit function
end if

'---- パッケージエラー
if InStr(error_package, p_ErrorCode) > 0 then	'package error
	w_msg = "<font color='#ff0000'>" _
				& "申し訳ございませんが､処理中にエラーが発生しました。<br>" _
				& "再度御注文ください。（二重申込にはなりません）<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "よくあるご質問は<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>こちら</a>" _
				& "</font>"
	exit function
end if

'---- システムエラー
if InStr(error_system, p_ErrorCode) > 0 then	'system error
	w_msg = "<font color='#ff0000'>" _
				& "申し訳ございませんが､センターシステム側で受付を停止しております。<br>" _
				& "しばらくしてから御注文ください。<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "よくあるご質問は<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>こちら</a>" _
				& "</font>"
	exit function
end if

'---- システムエラー
if (p_ErrorCode >= error_system_l) AND (p_ErrorCode <= error_system_h) then	'system error
	w_msg = "<font color='#ff0000'>" _
				& "申し訳ございませんが､センターシステム側で受付を停止しております。<br>" _
				& "しばらくしてから御注文ください。<br>" _
				& "Code: " & p_ErrorCode & "<br>" _
				& "よくあるご質問は<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>こちら</a>" _
				& "</font>"
	exit function
end if

'---- その他エラー
w_msg = "<font color='#ff0000'>" _
			& "申し訳ございませんが､今回の御注文はお受けできませんでした。<br>" _
			& "別のお支払方法で御注文願います。<br>" _
			& "Code: " & p_ErrorCode & "<br>" _
			& "よくあるご質問は<a href='http://www.soundhouse.co.jp/information/t_qanda.htm#card'>こちら</a>" _
			& "</font>"

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close

End function

%>
