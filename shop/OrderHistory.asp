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
'	購入履歴一覧ページ
'
'	会員メニューの｢ご購入履歴｣および、ページ指定された場合は、自分から呼び出される。
'	該当顧客の購入一覧を表示する。
'	顧客番号をSession("userID")から取出し、引数として用いる。
'
'	HTTPSでないとエラー
'	ログインしていないとエラー
'	ログインしていれば、Session("userID")に顧客番号がセットされている。
'	Session("userID")が空文字の時はエラー　｢ログインしてください。｣
'	Session("userID")で顧客情報が取出せなければエラー　｢ログインしてください。｣
'	エラーメッセージをセットしLogin.aspへRedirect
'
'	・該当顧客の受注情報を検索する
'	・受注情報はEmaxDBを使用する。(WebDBではない。)
'	・購入履歴の場合1ページへ表示する件数は、20件（プログラム内で指定）
'	・見積中と出荷準備中を1SQLで、購入履歴を別SQLで作成する
'	・各エリア内の表示順は見積日降順
'
'変更履歴
'2011/12/22 GV #1149 新規作成
'2012/08/11 if-web リニューアルレイアウト調整
'2012/11/24 ok 受注形態にスマートフォンを追加
'
'========================================================================
'On Error Resume Next

Const THIS_PAGE_NAME = "OrderHistory.asp"
Const PAGE_SIZE = 20						' 購入履歴情報の1ページあたりの表示行数

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wUserID

Dim wNotLogin					' ログインしていない

Dim wOrderHistryListHTML

Dim wIPage						' 表示するページ位置 (パラメータ)

'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
'---- Session変数
wDispMsg = Session("DispMsg")
Session("DispMsg") = ""
wErrMsg = Session("ErrMsg")
Session("ErrMsg") = ""

wUserID = Session("userID")

' Getパラメータ
wIPage = ReplaceInput(Trim(Request("IPage")))	' ページ位置

If wIPage = "" Or IsNumeric(wIPage) = False Then
	wIPage = 1
Else
	wIPage = CLng(wIPage)
End If

wNotLogin = False				' 初期状態はログインしている事を前提とする

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- エラーメッセージをセッションデータに登録   ' member系の他のページ処理にならう
If Err.Description <> "" Then
	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

If wNotLogin = True Then
	'---- ログインしていない場合はログインページへ
	Session("msg") = wMsg
	Server.Transfer "../shop/Login.asp"
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
Dim vRS_Cust
Dim vParam
Dim vTitleWord
Dim vTitleWordSave
Dim vOrderDateLabel
Dim vHistoryCount
Dim vHTML

If wUserID = "" Then
	'--- ログインしていなければエラー　｢ログインしてください。｣
	wNotLogin = True		' ログインされていない
	wMsg = "ログインしてください。"
	Exit Function
End If

' 顧客情報取得
Set vRS_Cust = get_customer()

If vRS_Cust.EOF = True Then
	'--- Session("userID")で顧客情報が取出せなければエラー　｢ログインしてください。｣
	wNotLogin = True		' ログインされていない
	wMsg = "ログインしてください。"
	Exit Function
End If

vRS_Cust.Close

Set vRS_Cust = Nothing

' 注文履歴件数の初期化
vHistoryCount = 0

'--- 該当顧客の受注一覧取り出し1 (見積中・出荷準備中)
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.受注番号 "
vSQL = vSQL & "    , a.受注日 "
vSQL = vSQL & "    , a.見積日 "
vSQL = vSQL & "    , a.出荷完了日 "
vSQL = vSQL & "    , a.受注形態 "
vSQL = vSQL & "    , a.支払方法 "
vSQL = vSQL & "    , a.Web受注変更開始日 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "受注 a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.削除日     IS NULL "
vSQL = vSQL & "    AND a.出荷完了日 IS NULL "
'vSQL = vSQL & "    AND a.受注形態 in ('E-mail','FAX','インターネット','携帯','電話','郵送','来店')"	'2012/11/24 ok Del
vSQL = vSQL & "    AND a.受注形態 in ('E-mail','FAX','インターネット','携帯','電話','郵送','来店','スマートフォン')"	'2012/11/24 ok Add
vSQL = vSQL & "    AND a.顧客番号   = " & wUserID & " "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "      CASE WHEN a.受注日 IS NULL "
vSQL = vSQL & "          THEN 1 "
vSQL = vSQL & "          ELSE 2 "
vSQL = vSQL & "      END "
vSQL = vSQL & "    , 見積日 DESC "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

vHTML = ""

If vRS.EOF = False Then

	vTitleWordSave = ""

	Do Until vRS.EOF = True

		'--- 出荷状況(タイトル) の判定
		vTitleWord = make_titleWord(vRS("受注日"), vRS("出荷完了日"))

		If vTitleWord <> vTitleWordSave Then

			If vTitleWordSave <> "" Then
				vHTML = vHTML & "</table>" & vbNewLine
			End If

			' 現在処理中のタイトルを待避
			vTitleWordSave = vTitleWord

			'--- 注文日列のタイトルラベル決定
			If vTitleWord = "お見積" Then
				vOrderDateLabel = "お見積日"
			ElseIf vTitleWord = "出荷準備中" Then
				vOrderDateLabel = "ご注文日"
			ElseIf vTitleWord = "ご購入履歴" Then
				vOrderDateLabel = "ご注文日"
			Else
				vOrderDateLabel = "ご注文日"
			End If

			'--- タイトル生成
			vHTML = vHTML & "<p class='table_bar'>" & vTitleWord & "</p>" & vbNewLine

			vHTML = vHTML & "<table class='order_history_list'>" & vbNewLine
			vHTML = vHTML & "  <tr>" & vbNewLine
			vHTML = vHTML & "    <th>" & vOrderDateLabel & "</th>" & vbNewLine
			vHTML = vHTML & "    <th>ご注文番号</th>" & vbNewLine
			vHTML = vHTML & "    <th>ご注文方法</th>" & vbNewLine
			vHTML = vHTML & "    <th>お支払方法</th>" & vbNewLine
			vHTML = vHTML & "  </tr>" & vbNewLine

		End If

		'--- 明細行生成
		vHTML = vHTML & make_orderHistoryHTML(vRS)

		vRS.MoveNext

	Loop

	vHTML = vHTML & "</table>" & vbNewLine

	' 注文履歴件数を確認用に待避
	vHistoryCount = vRS.RecordCount

End If

vRS.Close

'--- 該当顧客の受注一覧取り出し2 (ご購入履歴)
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.受注番号 "
vSQL = vSQL & "    , a.受注日 "
vSQL = vSQL & "    , a.見積日 "
vSQL = vSQL & "    , a.出荷完了日 "
vSQL = vSQL & "    , a.受注形態 "
vSQL = vSQL & "    , a.支払方法 "
vSQL = vSQL & "    , a.Web受注変更開始日 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "受注 a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.削除日     IS NULL "
vSQL = vSQL & "    AND a.出荷完了日 IS NOT NULL "
'vSQL = vSQL & "    AND a.受注形態 in ('E-mail','FAX','インターネット','携帯','電話','郵送','来店')"	'2012/11/24 ok Del
vSQL = vSQL & "    AND a.受注形態 in ('E-mail','FAX','インターネット','携帯','電話','郵送','来店','スマートフォン')"	'2012/11/24 ok Add
vSQL = vSQL & "    AND a.顧客番号   = " & wUserID & " "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "    見積日 DESC "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	'--- 出荷状況(タイトル) 文字生成
	vTitleWord = "ご購入履歴"
	vOrderDateLabel = "ご注文日"

	'--- タイトル生成
	vHTML = vHTML & "<p class='table_bar'>" & vTitleWord & "</p>" & vbNewLine

	vHTML = vHTML & "<table class='order_history_list'>" & vbNewLine
	vHTML = vHTML & "  <tr>" & vbNewLine
	vHTML = vHTML & "    <th>" & vOrderDateLabel & "</th>" & vbNewLine
	vHTML = vHTML & "    <th>ご注文番号</th>" & vbNewLine
	vHTML = vHTML & "    <th>ご注文方法</th>" & vbNewLine
	vHTML = vHTML & "    <th>お支払方法</th>" & vbNewLine
	vHTML = vHTML & "  </tr>" & vbNewLine

	'--- 指定ページを表示する為のレコード位置付け(SearchListの処理に倣う)
	vRS.PageSize = PAGE_SIZE
	If wIPage > ((vRS.RecordCount + (PAGE_SIZE - 1)) / PAGE_SIZE) Then		'MAXページを超える場合は最終ページへ
		wIPage = Fix(vRS.RecordCount / PAGE_SIZE)
	End If

	' レコード位置の位置付け
	vRS.AbsolutePage = wIPage

	For i = 0 To (vRS.PageSize - 1)

		'--- 明細行生成
		vHTML = vHTML & make_orderHistoryHTML(vRS)

		vRS.MoveNext

		If vRS.EOF Then
			Exit For
		End If

	Next

	vHTML = vHTML & "</table>" & vbNewLine

	' 注文履歴件数を確認用に待避
	vHistoryCount = vHistoryCount + vRS.RecordCount

	'--- ページ遷移部HTML生成
	vHTML = vHTML & make_pageNaviHTML(vRS, wIPage)

End If

vRS.Close

Set vRS = Nothing

'--- 購入履歴の存在確認
If vHistoryCount <= 0 Then

	wMsg = "購入履歴がありません。"
	Exit Function

End If

wOrderHistryListHTML = vHTML

End function

'========================================================================
'
'	Function	顧客情報の取り出し
'
'========================================================================
Function get_customer()

Dim vRS
Dim vSQL

'---- 顧客情報取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.顧客番号 "
vSQL = vSQL & "    , a.ユーザーID "
vSQL = vSQL & "    , a.顧客名 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web顧客 a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.顧客番号 = " & wUserID
vSQL = vSQL & "    AND a.Web不掲載フラグ <> 'Y'"

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

Set get_customer = vRS

End Function

'========================================================================
'
'	Function	注文履歴用HTML生成 (データ部1行分)
'
'========================================================================
Function make_orderHistoryHTML(pobjRS)

Dim vHTML

If pobjRS.EOF = True Then
    Exit Function
End If

vHTML = ""

vHTML = vHTML & "  <tr>" & vbNewLine

vHTML = vHTML & "    <td>" & formatDateYYYYMMDD(pobjRS("見積日")) & "</td>" & vbNewLine
vHTML = vHTML & "    <td><a href='OrderHistoryDetail.asp?OrderNo=" & pobjRS("受注番号") & "'>" & pobjRS("受注番号") & "</a></td>" & vbNewLine
vHTML = vHTML & "    <td>" & pobjRS("受注形態") & "</td>" & vbNewLine
vHTML = vHTML & "    <td>" & get_paymetMethodWord(pobjRS("支払方法")) & "</td>" & vbNewLine

vHTML = vHTML & "  </tr>" & vbNewLine

make_orderHistoryHTML = vHTML

End Function

'========================================================================
'
'	Function	ページ遷移部HTML生成
'
'========================================================================
Function make_pageNaviHTML(pobjRS, plngPage)

Dim vHTML
Dim i

vHTML = ""
vHTML = vHTML & "  <ol id='pagenavi'>" & vbNewLine

If plngPage <> 1 Then
	' 前のページ
	vHTML = vHTML & "    <li id='go'><a href='JavaScript:page_onClick(" & plngPage - 1 & ");' title='前のページに戻る'><span>&laquo;</span></a></li>" & vbNewline
End If

For i = 1 To pobjRS.PageCount
	If i = plngPage Then
		' 現在のページ
		vHTML = vHTML & "    <li id='now'><a href='JavaScript:void(0);'>" & i & "</a></li>" & vbNewLine
	Else
		' ページ番号指定
		vHTML = vHTML & "    <li><a href='JavaScript:page_onClick(" & i & ");'>" & i & "</a></li>" & vbNewLine
	End If
next

If plngPage <> pobjRS.PageCount Then
	' 次のページ
	vHTML = vHTML & "    <li id='go'><a href='JavaScript:page_onClick(" & plngPage + 1 & ");' title='次のページに進む'><span>&raquo;</span></a></li>" & vbNewline
End If

vHTML = vHTML & "  </ol>" & vbNewLine

make_pageNaviHTML = vHTML

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
'	Function	表示用支払い方法文字の生成
'
'	Note
'	  支払方法              表示文字
'	──────────────────────
'	  コンビニ支払       → "コンビニ払い"
'	  ネットバンキング   → "コンビニ払い"
'	  ゆうちょ           → "コンビニ払い"
'	  ローン(頭金あり)   → "ローン"
'	  ローン(頭金なし)   → "ローン"
'	  ローン(頭金無し)   → "ローン"
'	  銀行振込           → "銀行振込"
'	  代引き             → "代金引換"
'	  現金               → (支払方法そのまま)
'	  売掛               → (支払方法そのまま)
'	  アマゾン           → (支払方法そのまま)
'	  クレジットカード   → (支払方法そのまま)
'
'========================================================================
Function get_paymetMethodWord(pstrPaymetMethod)

Dim vDisplayWord

If IsNull(pstrPaymetMethod) = True Then
	' Null は判定不能
	Exit Function
End If

If pstrPaymetMethod = "代引き" Then
	vDisplayWord = "代金引換"
ElseIf pstrPaymetMethod = "コンビニ支払" Then
	vDisplayWord = "コンビニ払い"
ElseIf pstrPaymetMethod = "ネットバンキング" Then
	vDisplayWord = "コンビニ払い"
ElseIf pstrPaymetMethod = "ゆうちょ" Then
	vDisplayWord = "コンビニ払い"
ElseIf pstrPaymetMethod = "銀行振込" Then
	vDisplayWord = "銀行振込"
ElseIf InStr(pstrPaymetMethod, "ローン") > 0 Then
	vDisplayWord = "ローン"
Else
	vDisplayWord = pstrPaymetMethod
End If

get_paymetMethodWord = vDisplayWord

End Function

'========================================================================
'
'	Function	購入履歴のタイトル文字生成
'
'========================================================================
Function make_titleWord(pdatOrderDate, pdatShipCompleteDate)

Dim vTitleWord

If IsNull(pdatOrderDate) Then
	'--- 受注日がNullの場合
	vTitleWord = "お見積"
ElseIf IsNull(pdatOrderDate) = False And IsNull(pdatShipCompleteDate) Then
	'--- 受注日がNullでなく、出荷完了日がNullの場合
	vTitleWord = "出荷準備中"
ElseIf IsNull(pdatShipCompleteDate) = False Then
	'--- 出荷完了日がNullの場合
	vTitleWord = "ご購入履歴"
Else
	vTitleWord = "ご購入履歴"
End If

make_TitleWord = vTitleWord

End Function

'========================================================================
%>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>ご購入履歴｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel='stylesheet' href='../member/style/mypage.css?20120818' type='text/css'>
<script type='text/javascript'>
function page_onClick(p_page){
	document.f_pagenavi.IPage.value = p_page;
	document.f_pagenavi.submit();
}
</script>
</head>
<body>
<!--#include file="../Navi/NaviTop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  
  <!-- コンテンツstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li><a href="../member/Mypage.asp">マイページ</a></li>
        <li class="now">ご購入履歴</li>
      </ul>
    </div></div></div>

    <h1 class="title">ご購入履歴</h1>

<div class="center_pane">

<% If wErrMsg <> "" Then %>
<p class="error"><% = wErrMsg %></p>
<% Else %>
<%     If wDispMsg <> "" Then %>
<p class="renew"><% = wDispMsg %></p>
<%     End If %>
<%     If wMsg <> "" Then %>
<p class="error"><% = wMsg %></p>
<%     End If %>
  <% = wOrderHistryListHTML %>
<% End If %>
</div>

<!-- #include file="../Navi/MyPageMenu.inc"-->

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<form name='f_pagenavi' method='get' action='OrderHistory.asp'>
	<input type='hidden' name='IPage' value='1'>
</form>
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>