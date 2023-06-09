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
'	個別注文情報ページ
'
'	OrderHistory.aspの受注番号リンクから呼び出される
'	購入情報の詳細情報を表示する。
'	コントロールマスタより取出した変更可能時間帯で、かつ受注.担当者コード = 'internet' の場合、キャンセル変更を可能とする。
'
'	HTTPSでないとエラー
'	ログインしていないとエラー
'	ログインしていれば、Session("userID")に顧客番号がセットされている。
'	Session("userID")が空文字の時はエラー　｢ログインしてください。｣
'	Session("userID")で顧客情報が取出せなければエラー　｢ログインしてください。｣
'	エラーメッセージをセットしLogin.aspへRedirect
'
'	・該当顧客の受注情報を検索する。ヘッダ部分、未出荷部分、出荷完了部分を別々に取出す。
'	・ヘッダーが取り出せない場合はエラー｢該当の注文情報がありません。｣ OrderHistory.asp へ Redirect
'	・受注情報はEmaxDBを使用する。(WebDBではない。)
'
'変更履歴
'2011/12/27 GV #1149 新規作成
'2012/08/11 if-web リニューアルレイアウト調整
'2013/04/30 if-web 送り状番号表示をコメントアウト
'2013/07/11 GV #1507 レビュー編集機能
'
'========================================================================
On Error Resume Next

Const THIS_PAGE_NAME = "OrderHistoryDetail.asp"
Const UPDATEABLE_STAFF_CD = "Internet"			' キャンセル・注文変更 可能な 受注.担当者コード

Const FIRST_STEP = True			' 1st step 対処

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wUserID

Dim wNotLogin					' ログインしていない
Dim wUpdateable					' オーダーの変更が可能
Dim wDeleteable					' オーダーのキャンセルが可能
Dim wTaxRate					' 消費税率
Dim wOrderUpdateStartTime		' 受注変更開始時間(コントロールマスタ)
Dim wOrderUpdateEndTime			' 受注変更終了時間(コントロールマスタ)

Dim wOrderDetailHTML

Dim wOrderNo					' 受注番号 (パラメータ)

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
wOrderNo = ReplaceInput(Trim(Request("OrderNo")))	' 受注番号

If wOrderNo = "" Or IsNumeric(wOrderNo) = False Then
	wOrderNo = 0				' main でエラーとして取り扱う
Else
	wOrderNo = CLng(wOrderNo)
End If

wNotLogin = False				' 初期状態はログインしている事を前提とする

wUpdateable = False				' オーダーの変更は不可
wDeleteable = False				' オーダーのキャンセルは不可

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

If wMsg <> "" Then
	'--- ヘッダーが取り出せない,受注が見つからない等の場合はエラー　OrderHistory.aspへRedirect
	Session("ErrMsg") = wMsg
	Response.Redirect "OrderHistory.asp"
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
Dim vCurrentTime			' 現在の時刻
Dim vStaffCd				' 担当者コード
Dim vParam
Dim vTitleWord
Dim vTitleWordSave
Dim vOrderDateLabel
Dim vTrackingNumber			' 送り状番号
Dim vTransporterCd			' 運送会社コード
Dim vTransporterName		' 運送会社名
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

' パラメータのチェック (受注番号)
If wOrderNo <= 0 Then
	'--- 不正な受注番号の場合　｢該当の注文情報がありません。｣　OrderHistory.aspへRedirect
	wMsg = "該当の注文情報がありません。"
	Exit Function
End If

'--- コントロールマスタより「受注変更開始時間」「受注変更終了時間」取出し
If get_updateTimeSlot(wOrderUpdateStartTime, wOrderUpdateEndTime) = False Then
	'--- コントロールマスタに定義無し
	wMsg = "エラーが発生しました。"
	Exit Function
End If

'--- ヘッダ部分の情報取出し
vSQL = ""
vSQL = vSQL & "SELECT TOP 1 "
vSQL = vSQL & "      a.受注番号 "
vSQL = vSQL & "    , a.見積日 "
vSQL = vSQL & "    , a.受注日 "
vSQL = vSQL & "    , a.出荷完了日 "
vSQL = vSQL & "    , a.受注形態 "
vSQL = vSQL & "    , a.支払方法 "
vSQL = vSQL & "    , a.商品合計金額 "
vSQL = vSQL & "    , a.送料 "
vSQL = vSQL & "    , a.代引手数料 "
vSQL = vSQL & "    , a.受注合計金額 "
vSQL = vSQL & "    , a.一括出荷フラグ "
vSQL = vSQL & "    , a.領収書宛先 "
vSQL = vSQL & "    , a.領収書但し書き "
vSQL = vSQL & "    , a.Web受注変更開始日 "
vSQL = vSQL & "    , a.消費税率 "
vSQL = vSQL & "    , a.運送会社コード "
vSQL = vSQL & "    , a.担当者コード "
vSQL = vSQL & "    , b.今回限り届先郵便番号 "
vSQL = vSQL & "    , b.今回限り届先都道府県 "
vSQL = vSQL & "    , b.今回限り届先住所 "
vSQL = vSQL & "    , b.今回限り届先名前 "
vSQL = vSQL & "    , b.最終指定納期 "
vSQL = vSQL & "    , b.最終時間指定 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注     a WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "受注明細 b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        b.受注番号 = a.受注番号 "
vSQL = vSQL & "    AND a.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "    AND a.顧客番号 = " & wUserID & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	'--- ヘッダーが取り出せない場合はエラー｢該当の注文情報がありません。｣　OrderHistory.aspへRedirect
	vRS.Close
	Set vRS = Nothing
	wMsg = "該当の注文情報がありません。"
	Exit Function
End If

'--- 消費税率取出し
wTaxRate = CLng(vRS("消費税率"))

'--- 運送会社コード取出し
vTransporterCd = vRS("運送会社コード")

vHTML = ""

'--- 出荷状況(タイトル) の判定
vTitleWord = make_titleWord(vRS("受注日"), vRS("出荷完了日"), vRS("Web受注変更開始日"))

vHTML = vHTML & "<p class='table_bar'>" & vTitleWord & "</p>" & vbNewLine

'--- ご注文日 ～ お支払方法
vHTML = vHTML & "<table class='order_history_list'>" & vbNewLine
vHTML = vHTML & "  <tr>" & vbNewLine
vHTML = vHTML & "    <th>ご注文日</th>" & vbNewLine
vHTML = vHTML & "    <th>ご注文番号</th>" & vbNewLine
vHTML = vHTML & "    <th>ご注文方法</th>" & vbNewLine
vHTML = vHTML & "    <th>お支払方法</th>" & vbNewLine
vHTML = vHTML & "  </tr>    " & vbNewLine
vHTML = vHTML & "  <tr>" & vbNewLine
vHTML = vHTML & "    <td>" & formatDateYYYYMMDD_J(vRS("見積日")) & "</td>" & vbNewLine
vHTML = vHTML & "    <td class='number'>" & vRS("受注番号") & "</td>" & vbNewLine
vHTML = vHTML & "    <td>" & vRS("受注形態") & "</td>" & vbNewLine
vHTML = vHTML & "    <td>" & get_paymetMethodWord(vRS("支払方法")) & "</td>" & vbNewLine
vHTML = vHTML & "  </tr>" & vbNewLine
vHTML = vHTML & "</table>" & vbNewLine

'--- お届け先・配送方法・日時指定・領主書・合計金額(商品合計,送料,代引手数料,ご購入合計金額)
vHTML = vHTML & "<dl class='modify_list'>" & vbNewLine
vHTML = vHTML & "  <dt class='address'>お届け先</dt>" & vbNewLine
vHTML = vHTML & "  <dd class='address'>" & vbNewLine
vHTML = vHTML & "〒" & vRS("今回限り届先郵便番号") & "<br>" & vbNewLine
vHTML = vHTML & "" & vRS("今回限り届先都道府県") & vRS("今回限り届先住所") & "<br>" & vbNewLine
vHTML = vHTML & "" & vRS("今回限り届先名前") & "&nbsp;様</dd>" & vbNewLine
vHTML = vHTML & "  <dt>配送方法</dt>" & vbNewLine
vHTML = vHTML & "  <dd>" & get_shipTypeWord(vRS("一括出荷フラグ")) & "</dd>" & vbNewLine
vHTML = vHTML & "  <dt>日時指定</dt>" & vbNewLine
If IsDate(vRS("最終指定納期")) Then
	vHTML = vHTML & "  <dd>" & formatDateYYYYMMDD_J(vRS("最終指定納期")) & "　" & vRS("最終時間指定") & "</dd>" & vbNewLine
Else
	vHTML = vHTML & "  <dd>&nbsp;</dd>" & vbNewLine
End If
vHTML = vHTML & "  <dt>領収書</dt>" & vbNewLine
If IsNull(vRS("領収書宛先")) = False And vRS("領収書宛先") <> "" Then
	vHTML = vHTML & "  <dd>領収書宛先：" & vRS("領収書宛先") & " 様 / 領収書但し書き：" & vRS("領収書但し書き") & "</dd>" & vbNewLine
Else
	vHTML = vHTML & "  <dd>&nbsp;</dd>" & vbNewLine
End If
vHTML = vHTML & "  <dt class='total_accounts'>合計金額</dt>" & vbNewLine
vHTML = vHTML & "  <dd class='total_accounts'>" & vbNewLine
vHTML = vHTML & "    <ul>" & vbNewLine
vHTML = vHTML & "      <li>商品合計(税込)：" & FormatNumber(get_detailTotalPrice(wOrderNo, wTaxRate), 0) & "円</li>" & vbNewLine
vHTML = vHTML & "      <li>送料(税込)：" & FormatNumber(calc_taxInclusivePrice(vRS("送料"), wTaxRate), 0) & "円</li>" & vbNewLine
If vRS("支払方法") = "代引き" Then
	vHTML = vHTML & "      <li>代引手数料(税込)：" & FormatNumber(calc_taxInclusivePrice(vRS("代引手数料"), wTaxRate), 0) & "円</li>" & vbNewLine
End If
vHTML = vHTML & "      <li>ご購入合計金額(税込)：" & FormatNumber(vRS("受注合計金額"), 0) & "円</li>" & vbNewLine
vHTML = vHTML & "    </ul>" & vbNewLine
vHTML = vHTML & "  </dd>" & vbNewLine
vHTML = vHTML & "</dl>" & vbNewLine


'--- 操作用のボタン表示
vCurrentTime = Time()
vStaffCd = LCase(vRS("担当者コード") & "")

vHTML = vHTML & "<ul id='order_modify'>" & vbNewline

If FIRST_STEP = False Then	' 2011/12/22 1st step 対処	2011/12/28 hn キャンセルもなし

If isUpdateableTime(vCurrentTime, wOrderUpdateStartTime, wOrderUpdateEndTime) _
And vStaffCd = LCase(UPDATEABLE_STAFF_CD) _
And IsNull(vRS("Web受注変更開始日")) Then

	' キャンセル可能
	vHTML = vHTML & "  <li><a href='javascript:void(0);' title='注文内容をキャンセル' class='showLayer_ordercancel'>注文内容をキャンセル</a></li>" & vbNewline
	wDeleteable = True
Else
	vHTML = vHTML & "  <li>注文内容をキャンセル</li>" & vbNewline
	wDeleteable = False
End If

'If FIRST_STEP = False Then	' 2011/12/22 1st step 対処

If isUpdateableTime(vCurrentTime, wOrderUpdateStartTime, wOrderUpdateEndTime) _
And vStaffCd = LCase(UPDATEABLE_STAFF_CD) _
And IsNull(vRS("Web受注変更開始日")) Then
	' 変更可能
	vHTML = vHTML & "  <li><a href='javascript:void(0);' title='注文内容を変更' class='showLayer_ordermodify'>注文内容を変更</a></li>" & vbNewline
	wUpdateable = True
Else
	vHTML = vHTML & "  <li>注文内容を変更</li>" & vbNewline
	wUpdateable = False
End If

vHTML = vHTML & "  <li><a href='javascript:void(0);' title='この注文を再注文' class='showLayer_reorder'>この注文を再注文</a></li>" & vbNewline

End If	' 2011/12/22 1st step 対処

vHTML = vHTML & "  <li><a href='Inquiry.asp' title='注文内容のお問合せ'>注文内容のお問合せ</a></li>" & vbNewline
vHTML = vHTML & "</ul>" & vbNewline

vRS.Close


'--- 未出荷データの情報取出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      b.受注明細番号 "
vSQL = vSQL & "    , b.メーカーコード "
vSQL = vSQL & "    , b.商品コード "
vSQL = vSQL & "    , b.色 "
vSQL = vSQL & "    , b.規格 "
vSQL = vSQL & "    , b.受注単価 "
vSQL = vSQL & "    , b.受注数量 "
vSQL = vSQL & "    , b.受注引当合計数量 "
vSQL = vSQL & "    , b.出荷合計数量 "
vSQL = vSQL & "    , c.メーカー名 "
vSQL = vSQL & "    , d.商品名 "
vSQL = vSQL & "    , d.商品概略Web "
vSQL = vSQL & "    , d.商品画像ファイル名_小 "
vSQL = vSQL & "    , d.Web商品フラグ "
vSQL = vSQL & "    , x.出荷予定日 "
vSQL = vSQL & "    , x.ソース "
vSQL = vSQL & "    , x.出荷予定テキスト "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注明細 b WITH (NOLOCK) "
vSQL = vSQL & "        LEFT JOIN " & gLinkServer & "受注明細出荷予定 x WITH (NOLOCK) "
vSQL = vSQL & "          ON     x.受注番号     = b.受注番号 "
vSQL = vSQL & "             AND x.受注明細番号 = b.受注明細番号 "
vSQL = vSQL & "             AND x.出荷予定連番 = 1 "
vSQL = vSQL & "             AND x.変更日       = (SELECT MAX(y.変更日) "
vSQL = vSQL & "                                   FROM   " & gLinkServer & "受注明細出荷予定 y WITH (NOLOCK) "
vSQL = vSQL & "                                   WHERE      y.受注番号     = b.受注番号 "
vSQL = vSQL & "                                          AND y.受注明細番号 = b.受注明細番号) "
vSQL = vSQL & "    , " & gLinkServer & "メーカー c WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "商品 d WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        c.メーカーコード = b.メーカーコード "
vSQL = vSQL & "    AND d.メーカーコード = b.メーカーコード "
vSQL = vSQL & "    AND d.商品コード = b.商品コード "
vSQL = vSQL & "    AND b.セット品親明細番号 = 0 "
vSQL = vSQL & "    AND b.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "    AND b.受注数量 > b.出荷合計数量 "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "      c.メーカー名 "
vSQL = vSQL & "    , d.商品名 "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	'--- データが存在する場合のみ「未出荷データ」表示
	vHTML = vHTML & "<div class='order_history_container'>" & vbNewline
	vHTML = vHTML & "<p>未出荷</p>" & vbNewline

	Do Until vRS.EOF = True

		vHTML = vHTML & make_orderDetailHTML(vRS, wTaxRate)

		vRS.MoveNext

	Loop

	vHTML = vHTML & "</div>" & vbNewline

End If

vRS.Close


'--- 出荷完了データの情報取出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      b.受注明細番号 "
vSQL = vSQL & "    , b.メーカーコード "
vSQL = vSQL & "    , b.商品コード "
vSQL = vSQL & "    , b.色 "
vSQL = vSQL & "    , b.規格 "
vSQL = vSQL & "    , b.受注単価 "
vSQL = vSQL & "    , b.受注数量 "
vSQL = vSQL & "    , b.受注引当合計数量 "
vSQL = vSQL & "    , b.出荷合計数量 "
vSQL = vSQL & "    , f.出荷数量 "
vSQL = vSQL & "    , c.メーカー名 "
vSQL = vSQL & "    , d.商品名 "
vSQL = vSQL & "    , d.商品概略Web "
vSQL = vSQL & "    , d.商品画像ファイル名_小 "
vSQL = vSQL & "    , d.Web商品フラグ "
vSQL = vSQL & "    , e.送り状番号 "
vSQL = vSQL & "    , NULL AS 出荷予定日 "
vSQL = vSQL & "    , NULL AS 出荷予定テキスト "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注明細     b WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "メーカー     c WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "商品         d WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "受注送り状   e WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "出荷明細View f WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        c.メーカーコード = b.メーカーコード "
vSQL = vSQL & "    AND d.メーカーコード = b.メーカーコード "
vSQL = vSQL & "    AND d.商品コード = b.商品コード "
vSQL = vSQL & "    AND e.受注番号 = b.受注番号 "
vSQL = vSQL & "    AND f.出荷番号 = e.出荷番号 "
vSQL = vSQL & "    AND f.受注番号 = b.受注番号 "
vSQL = vSQL & "    AND f.受注明細番号 = b.受注明細番号 "
vSQL = vSQL & "    AND f.セット品親明細番号 = 0 "
vSQL = vSQL & "    AND b.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "ORDER BY  "
vSQL = vSQL & "      e.送り状番号 "
vSQL = vSQL & "    , c.メーカー名 "
vSQL = vSQL & "    , d.商品名 "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	'--- データが存在する場合のみ「出荷完了データ」表示
	vHTML = vHTML & "<div class='order_history_container'>" & vbNewline
	vHTML = vHTML & "<p>出荷完了</p>" & vbNewline

	vTrackingNumber = ""

	'--- 運送会社名
	vTransporterName = get_transporterName(vTransporterCd)

	Do Until vRS.EOF = True

'2013/04/30 if-web del s
'		If (vRS("送り状番号") & "") <> vTrackingNumber Then
'
'			vTrackingNumber = vRS("送り状番号") & ""
'
'			vHTML = vHTML & "<dl class='modify_list'>" & vbNewline
'			vHTML = vHTML & "  <dt>送り状番号</dt>" & vbNewline
'
'			' 送り状番号の表示
'			If vTransporterName = "佐川" Then
'				vHTML = vHTML & "  <dd><a href='http://k2k.sagawa-exp.co.jp/cgi-bin/mole.mcgi?oku01=" & vTrackingNumber & "' target='_blank'>" & vTrackingNumber & "（" & vTransporterName & "）</a></dd>" & vbNewline
'			ElseIf vTransporterName = "西濃" Then
'				vHTML = vHTML & "  <dd><a href='http://track.seino.co.jp/kamotsu/KamotsuPrintServlet?GNPNO1=" & vTrackingNumber & "&ACTION=DETAIL&NUMBER=1' target='_blank'>" & vTrackingNumber & "（" & vTransporterName & "）</a></dd>" & vbNewline
'			Else
'				vHTML = vHTML & "  <dd>" & vTrackingNumber & "（" & vTransporterName & "）</dd>" & vbNewline
'			End If
'
'			vHTML = vHTML & "</dl>" & vbNewline
'
'		End If
'2013/04/30 if-web del e

		' 明細表示
		vHTML = vHTML & make_orderDetailHTML(vRS, wTaxRate)

		vRS.MoveNext

	Loop

	vHTML = vHTML & "</div>" & vbNewline

End If

vRS.Close

Set vRS = Nothing

wOrderDetailHTML = vHTML

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
'	Function	商品合計金額の取得 (税込み金額)
'	Note		受注明細.受注単価 に対し、税込み金額を計算後、受注明細.受注数量を掛け、その全明細総合計
'
'========================================================================
Function get_detailTotalPrice(plngOrderNo, plngTaxRate)

Dim vRS
Dim vSQL
Dim vTotalPrice

get_detailTotalPrice = 0

'---- 受注明細の「受注単価」と「受注数量」を取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.受注単価 "
vSQL = vSQL & "    , a.受注数量 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "受注明細 a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.受注番号 = " & wOrderNo & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	vRS.Close
	Set vRS = Nothing
	Exit Function
End If

vTotalPrice = 0

Do Until vRS.EOF = True

	If IsNumeric(vRS("受注単価")) And IsNumeric(vRS("受注数量")) Then

		' 単価の税込み金額計算後、数量を掛ける
		vTotalPrice = vTotalPrice + (calcPrice(vRS("受注単価"), plngTaxRate) * vRS("受注数量"))

	End If

	vRS.MoveNext

Loop

vRS.Close
Set vRS = Nothing

get_detailTotalPrice = vTotalPrice

End Function

'========================================================================
'
'	Function	内税額の計算
'
'========================================================================
Function calc_taxInclusivePrice(plngPrice, plngTaxRate)

calc_taxInclusivePrice = Fix(plngPrice * (100 + plngTaxRate) / 100)

End Function

'========================================================================
'
'	Function	注文明細用HTML生成 (データ部1行分)
'
'========================================================================
Function make_orderDetailHTML(pobjRS, plngTaxRate)

Dim vHTML
Dim vItemName
Dim vWebItem					' Web商品フラグ
Dim vParam
Dim vProductDetailLink			' ProductDetail.asp へのリンク URL
Dim vShippingStatus				' 明細の出荷状況
Dim vShippingComplete			' 出荷完了
Dim vPrice						' 明細金額

If pobjRS.EOF = True Then
    Exit Function
End If

vWebItem = pobjRS("Web商品フラグ") & ""
If vWebItem = "Y" Then
	vParam = Server.URLEncode(pobjRS("メーカーコード") & "^" & pobjRS("商品コード") & "^" & Trim(pobjRS("色")) & "^" & Trim(pobjRS("規格")))
	vProductDetailLink = g_HTTP & "shop/ProductDetail.asp?item=" & vParam
Else
	vProductDetailLink = ""
End If

If (Trim(pobjRS("色")) & Trim(pobjRS("規格"))) <> "" Then
	vItemName = pobjRS("商品名") & "/" & Trim(pobjRS("色")) & "/" & Trim(pobjRS("規格"))
Else
	vItemName = pobjRS("商品名") & ""
End If

' 明細の出荷状況
vShippingStatus = ""
vShippingComplete = False

If pobjRS("受注数量") = pobjRS("出荷合計数量") Then

	vShippingStatus = "出荷完了"
	vShippingComplete = True

ElseIf pobjRS("受注数量") = pobjRS("受注引当合計数量") Then

	vShippingStatus = "出荷準備中"

ElseIf pobjRS("受注数量") > pobjRS("受注引当合計数量") Then

	If IsNull(pobjRS("出荷予定日")) _
	And IsNull(pobjRS("出荷予定テキスト")) Then

		vShippingStatus = "取り寄せ中"

	ElseIf IsNull(pobjRS("出荷予定日")) = False Then

		vShippingStatus = formatDateMMDD_J(pobjRS("出荷予定日")) & "　入荷予定"

	ElseIf IsNull(pobjRS("出荷予定テキスト")) = False Then

		vShippingStatus = pobjRS("出荷予定テキスト") & "　入荷予定"

	End If

End If

' 明細の金額計算
If IsNumeric(pobjRS("受注単価")) And IsNumeric(pobjRS("受注数量")) Then
	' 受注単価(税込み) * 受注数量   (受注単価(税込み) : calcPrice(受注単価, 消費税率))
	vPrice = calcPrice(pobjRS("受注単価"), plngTaxRate) * pobjRS("受注数量")
Else
	vPrice = 0
End If

vHTML = ""

vHTML = vHTML & "<table class='order_history'>" & vbNewline
vHTML = vHTML & "  <tr>" & vbNewline
vHTML = vHTML & "    <td class='list_left'>" & vbNewline
If (pobjRS("商品画像ファイル名_小") & "") <> "" _
And vWebItem = "Y" Then
	vHTML = vHTML & "      <a href='" & vProductDetailLink & "'><img src='prod_img/" & pobjRS("商品画像ファイル名_小") & "' width='100' height='50' alt=''></a>" & vbNewline
Else
	vHTML = vHTML & "      <img src='prod_img/" & pobjRS("商品画像ファイル名_小") & "' width='100' height='50' alt=''>" & vbNewline
End If
vHTML = vHTML & "    </td>" & vbNewline
vHTML = vHTML & "    <td>" & vbNewline
vHTML = vHTML & "      " & pobjRS("メーカー名") & "<br>" & vbNewline
If vWebItem = "Y" Then
	vHTML = vHTML & "      <a href='" & vProductDetailLink & "'>" & vItemName & "</a><br>" & vbNewline
Else
	vHTML = vHTML & "      " & vItemName & "<br>" & vbNewline
End If
vHTML = vHTML & "      " & pobjRS("商品概略Web") & vbNewline
vHTML = vHTML & "    </td>" & vbNewline
vHTML = vHTML & "    <td class='contact'>" & vbNewline
vHTML = vHTML & "      <ul>" & vbNewline
vHTML = vHTML & "        <li>" & vShippingStatus & "</li>" & vbNewline
vHTML = vHTML & "        <li>" & pobjRS("受注数量") & "点：" & FormatNumber(vPrice, 0) & "円（税込）</li>" & vbNewline
vHTML = vHTML & "        <li><a href='Inquiry.asp?MakerNm=" & Server.URLEncode(pobjRS("メーカー名")) & "&ProductCd=" & Server.URLEncode(pobjRS("商品コード")) & "' class='tipBtn'>この商品のお問合せ</a></li>" & vbNewline
If vShippingComplete = True _
And vWebItem = "Y" Then
'2013/07/11 GV #1507 mod start
	If isReviewEntered(pobjRS("メーカーコード"), pobjRS("商品コード"), wUserID) = False Then
		' レビュー未記入の場合のみ
'		vHTML = vHTML & "        <li><a href='" & vProductDetailLink & "&WriteReview=Y#review' class='tipBtn'>レビューを書く</a></li>" & vbNewline
		vHTML = vHTML & "        <li><a href='" & g_HTTPS & "shop/ReviewWrite.asp?item=" & vParam & "' class='tipBtn'>レビューを書く</a></li>" & vbNewline
	Else
'		vHTML = vHTML & "        <li>&nbsp;</li>" & vbNewline
		vHTML = vHTML & "        <li><a href='" & g_HTTPS & "shop/ReviewWrite.asp?item=" & vParam & "' class='tipBtn'>レビューを編集</a></li>" & vbNewline
	End If
'2013/07/11 GV #1507 mod start
Else
	vHTML = vHTML & "        <li>&nbsp;</li>" & vbNewline
End If
vHTML = vHTML & "      </ul>" & vbNewline
vHTML = vHTML & "    </td>" & vbNewline
vHTML = vHTML & "  </tr>" & vbNewline
vHTML = vHTML & "</table>" & vbNewline

make_orderDetailHTML = vHTML

End Function

'========================================================================
'
'	Function	日付けのフォーマット (YYYY年MM月DD日)
'
'========================================================================
Function formatDateYYYYMMDD_J(pdatDate)

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

formatDateYYYYMMDD_J = vDate

End Function

'========================================================================
'
'	Function	日付けのフォーマット (MM月DD日)
'
'========================================================================
Function formatDateMMDD_J(pdatDate)

Dim vDate

If IsNull(pdatDate) = True Then
	' Null は計算不能
	Exit Function
End If

If IsDate(pdatDate) = False Then
	' 日付けでなければ計算不能
	Exit Function
End If

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

formatDateMMDD_J = vDate

End Function

'========================================================================
'
'	Function	日付けのフォーマット (YYYY/MM/DD)
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

vDate = DatePart("yyyy", pdatDate) & "/"

If DatePart("m", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("m", pdatDate)
Else
	vDate = vDate & DatePart("m", pdatDate)
End If

vDate = vDate & "/"

If DatePart("d", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("d", pdatDate)
Else
	vDate = vDate & DatePart("d", pdatDate)
End If

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
'	Function	運送会社名の取り出し
'
'========================================================================
Function get_transporterName(pstrTransporterCd)

Dim vRS
Dim vSQL
Dim vTransporterName

'---- 運送会社名取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "    a.運送会社略称 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "運送会社 a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.運送会社コード = " & pstrTransporterCd

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then
	vTransporterName = vRS("運送会社略称") & ""
Else
	vTransporterName = ""
End If

vRS.Close
Set vRS = Nothing

get_transporterName = vTransporterName

End Function

'========================================================================
'
'	Function	既にレビューを記入済みか？
'
'========================================================================
Function isReviewEntered(pstrMakerCd, pstrItemCd, plngCustNo)

Dim vRS
Dim vSQL
Dim vEntered

'---- 商品レビュー取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "    a.ID "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    商品レビュー a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.メーカーコード = '" & pstrMakerCd & "' "
vSQL = vSQL & "    AND a.商品コード = '" & escapeSingleQuote(pstrItemCd) & "' "
vSQL = vSQL & "    AND a.顧客番号 = " & plngCustNo & " "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	' レビューなし (レビュー未記入)
	vEntered = False
Else
	vEntered = True
End If

vRS.Close
Set vRS = Nothing

isReviewEntered = vEntered

End Function

'========================================================================
'
'	Function	受注変更可能時間帯の取得
'
'========================================================================
Function get_updateTimeSlot(pdatStartTime, pdatEndTime)

Dim vstrItemChar1
Dim vstrItemChar2
Dim vdblItemNum1
Dim vdblItemNum2
Dim vdatItemDate1
Dim vdatItemDate2

get_updateTimeSlot = False

'--- コントロールマスタ取り出し
Call getCntlMst("Web", "受注", "受注変更可能時間帯", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

If IsDate(vstrItemChar1) = False Then
	' システムの設定値不良
	Exit Function
End If

If IsDate(vstrItemChar2) = False Then
	' システムの設定値不良
	Exit Function
End If

'--- 開始終了時刻 返却
pdatStartTime = CDate(vstrItemChar1)
pdatEndTime = CDate(vstrItemChar2)

get_updateTimeSlot = True

End function

'========================================================================
'
'	Function	受注変更可能時間帯の判定
'	Note		現在時刻が、受注変更開始時間～受注変更終了時間 の間であるか判定する
'
'========================================================================
Function isUpdateableTime(pdatTargetTime, pdatStartTime, pdatEndTime)

isUpdateableTime = False

If pdatStartTime < pdatEndTime Then
	' 日を跨がない判定
	If pdatStartTime <= pdatTargetTime And pdatTargetTime <= pdatEndTime Then
		' 範囲内
		isUpdateableTime = True
	End If
Else
	' 日を跨ぐ判定 (終了時刻の方が早い時刻の場合 Start ～ 23:59:59 Or 0:00 ～ End)
	If pdatStartTime <= pdatTargetTime And pdatTargetTime <= CDate("23:59:59") _
	Or CDate("00:00:00") <= pdatTargetTime And pdatTargetTime <= pdatEndTime Then
		' 範囲内
		isUpdateableTime = True
	End If
End If

End function

'========================================================================
'
'	Function	出荷状況のタイトル文字生成
'
'	Note
'		下記の順番でチェックを行う
'		0. 変更中     : Web受注変更開始日 IS NOT NULL の時
'		1. 出荷完了   : 出荷完了日 IS NOT NULL の時
'		2. 出荷準備中 : 受注日 IS NOT NULL AND 出荷完了日 IS NULL の時
'		3. お見積り   : 受注日 IS NULL の時
'
'========================================================================
Function make_titleWord(pdatOrderDate, pdatShipCompleteDate, pdatWebOrderUpdateStartDate)

If IsNull(pdatWebOrderUpdateStartDate) = False Then
	'--- Web受注変更開始日がNullでない場合
	make_titleWord = "変更中"
	Exit Function
ElseIf IsNull(pdatShipCompleteDate) = False Then
	'--- 出荷完了日がNullでない場合
	make_titleWord = "出荷完了"
	Exit Function
ElseIf IsNull(pdatOrderDate) = False And IsNull(pdatShipCompleteDate) Then
	'--- 受注日がNullでなく、出荷完了日がNullの場合
	make_titleWord = "出荷準備中"
	Exit Function
ElseIf IsNull(pdatOrderDate) Then
	'--- 受注日がNullの場合
	make_titleWord = "お見積り"
	Exit Function
End If

End Function

'========================================================================
'
'	Function	表示用配送方法文字の生成
'
'========================================================================
Function get_shipTypeWord(pstrIkkatsuSyukkaFlg)

If IsNull(pstrIkkatsuSyukkaFlg) = True Then
	' Null は判定不能
	Exit Function
End If

If pstrIkkatsuSyukkaFlg = "Y" Then
	'--- 一括出荷の場合
	get_shipTypeWord = "一括出荷"
	Exit Function
Else
	get_shipTypeWord = "在庫商品から出荷"
	Exit Function
End If

End Function

'========================================================================
'
'	Function	SQLサーバ用、シングルクオーテーション抑止文字の付加
'
'========================================================================
Function escapeSingleQuote(pstrStringValue)

If IsNull(pstrStringValue) Then
	Exit Function
End If

escapeSingleQuote = Replace(pstrStringValue, "'", "''")

End Function

'========================================================================
%>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>ご注文情報｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel='stylesheet' href='../member/style/mypage.css?201309xx' type='text/css'>
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
        <li><a href="OrderHistory.asp">ご購入履歴</a></li>
        <li class="now">ご注文情報</li>
      </ul>
    </div></div></div>

    <h1 class="title">ご注文情報</h1>

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
  <% = wOrderDetailHTML %>
<% End If %>
</div>

<!-- #include file="../Navi/MyPageMenu.inc"-->

<% If wDeleteable Then %>
<% ' 注文キャンセルポップアップ %>
<div class="overContent" id="overContent_ordercancel">
  <h2>注文内容をキャンセル</h2>
  <p>この注文内容をキャンセルしますか？一度キャンセルされた注文を戻す事はできません。</p>

  <form name="f_cancel" method="post" action="OrderHistoryDelete.asp">
    <input type="submit" value="注文内容をキャンセルする" class="strong_btn">
    <input type='hidden' name='OrderNo' value='<% = wOrderNo %>'>
  </form>

  <ul class="back">
    <li><a href="javascript:void(0);" onClick="backclose();">←&nbsp;戻る</a></li>
  </ul>
</div>
<% End If %>

<% If wUpdateable Then %>
<% ' 注文変更ポップアップ %>
<div class="overContent" id="overContent_ordermodify">
  <h2>注文内容を変更</h2>
  <p>ショッピングカートページに戻ってご注文手続きをやり直すことができます。<br>
※現在カートに入っている商品は上書きされます。<br>
※注文内容の変更を途中でやめた場合には、今のご注文内容はキャンセルされません。</p>

  <form name="f_change" method="post" action="OrderHistoryChange.asp">
    <input type="submit" value="注文内容を変更する" class="strong_btn">
    <input type='hidden' name='OrderNo' value='<% = wOrderNo %>'>
  </form>

  <ul class="back">
    <li><a href="javascript:void(0);" onClick="backclose();">←&nbsp;戻る</a></li>
  </ul>

  <div id="ordermodify_flow">
    <p>ご注文内容の変更の流れ</p>
    <dl>
      <dt><img src="images/shopping_step1_off.gif" alt="ショッピングカート"></dt>
      <dd>ショッピングカートページで、商品の追加・削除を行います。</dd>
      <dt><img src="images/shopping_step2_off.gif" alt="お届け先、お支払方法の選択"></dt>
      <dd>お届け先、お支払方法等の変更ができます。</dd>
      <dt><img src="images/shopping_step3_off.gif" alt="ご注文内容の確認"></dt>
      <dd>変更いただいたご注文内容を確認します。</dd>
      <dt><img src="images/shopping_step4_off.gif" alt="ご注文完了"></dt>
      <dd id="off">ご注文内容の変更が完了します。<br>
変更前のご注文内容がキャンセルされ、変更いただきましたご注文にて承ります。<br>
ご登録のメールアドレス宛てに確認メールを送信いたしますので内容をご確認ください。</dd>
    </dl>
  </div>
</div>
<% End If %>

<% If FIRST_STEP = False Then	' 1st step 対処 %>
<% ' 再注文ポップアップ %>
<div class="overContent" id="overContent_reorder">
  <h2>この注文を再注文</h2>
  <p>このご注文内容と同じ内容で、再びご注文いただけます。<br>
現在カートにある商品に追加するか、カートの内容を上書きするかお選びください。</p>

  <form name="f_reorder" method="post" action="OrderHistoryCopy.asp">
    <ul id="reorder_select">
      <li><input type="button" value="カートに追加して再注文する" class="strong_btn" onClick="reorder_onClick('Y');"></li>
      <li><input type="button" value="カートを上書きして再注文する" class="strong_btn" onClick="reorder_onClick('N');"></li>
    </ul>
    <input type='hidden' name='OrderNo' value='<% = wOrderNo %>'>
    <input type='hidden' name='addItem' value='N'>
  </form>

  <ul class="back">
    <li><a href="javascript:void(0);" onClick="backclose();">←&nbsp;戻る</a></li>
  </ul>
</div>
<% End If	' 1st step 対処 %>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/showLayer.js"></script>
<script type="text/javascript">
function backclose(){
	$("#glayLayer").hide();
	$("#overLayer").fadeOut(500);
}
function reorder_onClick(p_additem){
	location.href = 'OrderHistoryCopy.asp?OrderNo=<% = wOrderNo %>&addItem=' + p_additem;
}
</script>
</body>
</html>