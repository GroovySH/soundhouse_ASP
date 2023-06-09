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
'	注文情報削除処理
'
'	OrderHistoryDetail.aspのキャンセルボタンから呼び出される。
'	概要：
'		該当の受注の削除日をセットする。
'		該当の受注明細の商品の在庫戻しを行う。
'		Web変更受注(Emax)に変更開始日、変更終了日をセットする。
'
'	HTTPSでないとエラー
'	ログインしていないとエラー
'	ログインしていれば、Session("userID")に顧客番号がセットされている。
'	Session("userID")が空文字の時はエラー　｢ログインしてください。｣
'	Session("userID")で顧客情報が取出せなければエラー　｢ログインしてください。｣
'	エラーメッセージをセットしLogin.aspへRedirect
'
'	・キャンセル可能時間帯、受注のチェック
'	・データ抽出
'	・Emaxの受注の更新
'	・Emaxの色規格別在庫の更新
'	・EmaxのWeb変更受注の登録
'	・OrderHistory.aspを呼び出し
'
'変更履歴
'2011/12/26 GV #1149 新規作成
'
'========================================================================
On Error Resume Next

Const THIS_PAGE_NAME = "OrderHistoryDelete.asp"
Const UPDATEABLE_STAFF_CD = "Internet"			' キャンセル・注文変更 可能な 受注.担当者コード

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' エラーメッセージ (他のページから渡されるメッセージ)
Dim wDispMsg					' 通常メッセージ(エラー以外) (他のページから渡されるメッセージ)
Dim wErrDesc
Dim wMsg						' エラーメッセージ (本ページで作成するメッセージ)
Dim wUserID

Dim wNotLogin					' ログインしていない
Dim wOrderUpdateStartTime		' 受注変更開始時間(コントロールマスタ)
Dim wOrderUpdateEndTime			' 受注変更終了時間(コントロールマスタ)

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

' パラメータ
wOrderNo = ReplaceInput(Trim(Request("OrderNo")))	' 受注番号

If wOrderNo = "" Or IsNumeric(wOrderNo) = False Then
	wOrderNo = 0				' main でエラーとして取り扱う
Else
	wOrderNo = CLng(wOrderNo)
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

If wMsg <> "" Then
	'--- ヘッダーが取り出せない,受注が見つからない,キャンセル不可等の場合はエラー　OrderHistoryDetail.aspへRedirect
	Session("ErrMsg") = wMsg
	Response.Redirect "OrderHistoryDetail.asp?OrderNo=" & wOrderNo
End If

'--- 正常終了の場合 OrderHistory.asp へメッセージ付きで戻る
Session("DispMsg") = "ご注文はキャンセルされました。"
Response.Redirect "OrderHistory.asp"


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
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      受注番号 "
vSQL = vSQL & "    , 受注日 "
vSQL = vSQL & "    , 支払方法 "
vSQL = vSQL & "    , Web受注変更開始日 "
vSQL = vSQL & "    , 担当者コード "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "受注 WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        受注番号 = " & wOrderNo & " "
vSQL = vSQL & "    AND 顧客番号 = " & wUserID & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	'--- ヘッダーが取り出せない場合はエラー｢該当の注文情報がありません。｣　OrderHistoryDetai.aspへRedirect
	vRS.Close
	Set vRS = Nothing
	wMsg = "該当の注文情報がありません。"
	Exit Function
End If

'--- キャンセル可能時間帯、受注のチェック
vCurrentTime = Time()
vStaffCd = LCase(vRS("担当者コード") & "")

If isUpdateableTime(vCurrentTime, wOrderUpdateStartTime, wOrderUpdateEndTime) = False _
Or vStaffCd <> LCase(UPDATEABLE_STAFF_CD) _
Or IsNull(vRS("Web受注変更開始日")) = False Then
	'--- キャンセル不可のオーダー の場合はエラー｢現在この注文のキャンセルは行えません。｣　OrderHistoryDetai.aspへRedirect
	vRS.Close
	Set vRS = Nothing
	wMsg = "現在この注文のキャンセルは行えません。"
	Exit Function
End If


'--- トランザクション開始
Connection.BeginTrans


'--- Emaxの受注の更新 (削除日の設定)
If update_order_deleteDate() = False Then
	Connection.RollbackTrans
	Exit Function
End If

'===========コメント 2011/12/22 hn
''--- Emaxの色規格別在庫の更新
'If update_Inventory(vRS("支払方法") & "") = False Then
'	Connection.RollbackTrans
'	Exit Function
'End If
'===========コメント

'--- EmaxのWeb変更受注の登録
If insert_WebUpdateOrder() = False Then
	Connection.RollbackTrans
	Exit Function
End If

'--- キャンセルメール送信
If send_cancelMail() = False Then
	Connection.RollbackTrans
	Exit Function
End If


vRS.Close

'--- トランザクション終了
Connection.CommitTrans

Set vRS = Nothing

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
vSQL = vSQL & "    , b.顧客郵便番号 "
vSQL = vSQL & "    , b.顧客都道府県 "
vSQL = vSQL & "    , b.顧客住所 "
vSQL = vSQL & "    , c.顧客電話番号 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      Web顧客     a WITH (NOLOCK) "
vSQL = vSQL & "    , Web顧客住所 b WITH (NOLOCK) "
vSQL = vSQL & "    , Web顧客住所電話番号 c WITH (NOLOCK)"
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.顧客番号 = " & wUserID & " "
vSQL = vSQL & "    AND a.Web不掲載フラグ <> 'Y' "
vSQL = vSQL & "    AND b.顧客番号 = a.顧客番号 "
vSQL = vSQL & "    AND b.住所連番 = 1 "
vSQL = vSQL & "    AND c.顧客番号 = a.顧客番号 "
vSQL = vSQL & "    AND c.住所連番 = 1 "
vSQL = vSQL & "    AND c.電話連番 = 1 "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

Set get_customer = vRS

End Function

'========================================================================
'
'	Function	受注情報の取り出し
'	Note		キャンセルメール送信用
'
'========================================================================
Function get_orderInfo()

Dim vRS
Dim vSQL

'---- 受注情報取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.受注番号 "
vSQL = vSQL & "    , a.見積日 "
vSQL = vSQL & "    , a.支払方法 "
vSQL = vSQL & "    , a.支払方法 "
vSQL = vSQL & "    , a.顧客E_mail "
vSQL = vSQL & "    , b.メーカーコード "
vSQL = vSQL & "    , b.商品コード "
vSQL = vSQL & "    , b.色 "
vSQL = vSQL & "    , b.規格 "
vSQL = vSQL & "    , b.受注単価 "
vSQL = vSQL & "    , b.受注数量 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "受注     a WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "受注明細 b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.受注番号 = " & wOrderNo & " "
vSQL = vSQL & "    AND a.受注番号 = b.受注番号 "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

Set get_orderInfo = vRS

End Function

'========================================================================
'
'	Function	商品情報の取り出し
'	Note		キャンセルメール送信用
'
'========================================================================
Function get_itemInfo(pstrMaketCd, pstrItemCd)

Dim vRS
Dim vSQL

'---- 商品情報取り出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.商品名 "
vSQL = vSQL & "    , a.Web商品フラグ "
vSQL = vSQL & "    , b.メーカー名 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      Web商品  a WITH (NOLOCK) "
vSQL = vSQL & "    , メーカー b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.メーカーコード = '" & pstrMaketCd & "' "
vSQL = vSQL & "    AND a.商品コード     = '" & escapeSingleQuote(pstrItemCd) & "' "
vSQL = vSQL & "    AND a.メーカーコード = b.メーカーコード "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

Set get_itemInfo = vRS

End Function

'========================================================================
'
'	Function	Emaxの受注の削除日設定
'
'========================================================================
Function update_order_deleteDate()

update_order_deleteDate = False

Dim vSQL
Dim vRS

vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      削除日 "
vSQL = vSQL & "    , 最終更新日 "
vSQL = vSQL & "    , 最終更新者コード "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "受注 "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    受注番号 = " & wOrderNo & " "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	wMsg = "該当の注文情報がありません。"
	vRS.Close
	Set vRS = Nothing
	Exit Function
End If

vRS("削除日") = Now()
vRS("最終更新日") = Now()
vRS("最終更新者コード") = UPDATEABLE_STAFF_CD

vRS.Update

vRS.Close
Set vRS = Nothing

update_order_deleteDate = True

End function

'========================================================================
'
'	Function	Emaxの色規格別在庫を更新
'
'========================================================================
Function update_Inventory(pstrPaymetMethod)

update_Inventory = False

Dim vSQL
Dim vRS
Dim vRS_Inventory
Dim vBItem				' B品フラグ
Dim vOrderNum			' 受注数量
Dim vInventoryReservNum	' 見積引当合計数量

'--- 受注明細の取出し
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      メーカーコード "
vSQL = vSQL & "    , 商品コード "
vSQL = vSQL & "    , 色 "
vSQL = vSQL & "    , 規格 "
vSQL = vSQL & "    , B品フラグ "
vSQL = vSQL & "    , 受注数量 "
vSQL = vSQL & "    , 受注引当合計数量 "
vSQL = vSQL & "    , 見積引当合計数量 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "受注明細 WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    受注番号 = " & wOrderNo & " "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	wMsg = "該当の注文情報がありません。"
	vRS.Close
	Set vRS = Nothing
	Exit Function
End If

Do Until vRS.EOF

	vBItem = vRS("B品フラグ") & ""

	'--- 更新する「色規格別在庫」の取出し
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "      引当可能数量 "
	vSQL = vSQL & "    , 受注引当数量 "
	vSQL = vSQL & "    , 受注残数量 "
	vSQL = vSQL & "    , 見積取置数量 "
	vSQL = vSQL & "    , B品引当可能数量 "
	vSQL = vSQL & "    , B品受注引当数量 "
	vSQL = vSQL & "    , B品見積取置数量 "
	vSQL = vSQL & "    , 最終更新日 "
	vSQL = vSQL & "    , 最終更新者コード "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "    " & gLinkServer & "色規格別在庫 "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "        メーカーコード = '" & vRS("メーカーコード") & "' "
	vSQL = vSQL & "    AND 商品コード     = '" & escapeSingleQuote(vRS("商品コード")) & "' "
	vSQL = vSQL & "    AND 色             = '" & vRS("色") & "' "
	vSQL = vSQL & "    AND 規格           = '" & vRS("規格") & "' "

'@@@@@@Response.Write(vSQL)

	Set vRS_Inventory = Server.CreateObject("ADODB.Recordset")
	vRS_Inventory.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	If vRS_Inventory.EOF = False Then

		If pstrPaymetMethod = "代引き" Then

			' 代引き

			If vBItem <> "Y" Then

				' Not B品
				vRS_Inventory("引当可能数量") = Nz(vRS_Inventory("引当可能数量"), 0) + Nz(vRS("受注数量"), 0)
				vRS_Inventory("受注引当数量") = Nz(vRS_Inventory("受注引当数量"), 0) - Nz(vRS("受注数量"), 0)
				vRS_Inventory("受注残数量")   = Nz(vRS_Inventory("受注残数量"), 0)   - (Nz(vRS("受注数量"), 0) - Nz(vRS("受注引当合計数量"), 0))

			Else

				' B品
				vRS_Inventory("B品引当可能数量") = Nz(vRS_Inventory("B品引当可能数量"), 0) + Nz(vRS("受注数量"), 0)
				vRS_Inventory("B品受注引当数量") = Nz(vRS_Inventory("B品受注引当数量"), 0) - Nz(vRS("受注数量"), 0)

			End If

		Else

			' Not 代引き

			If vBItem <> "Y" Then

				' Not B品
				vRS_Inventory("引当可能数量") = Nz(vRS_Inventory("引当可能数量"), 0) + Nz(vRS("見積引当合計数量"), 0)
				vRS_Inventory("見積取置数量") = Nz(vRS_Inventory("見積取置数量"), 0) - Nz(vRS("見積引当合計数量"), 0)

			Else

				' B品
				vRS_Inventory("B品引当可能数量") = Nz(vRS_Inventory("B品引当可能数量"), 0) + Nz(vRS("見積引当合計数量"), 0)
				vRS_Inventory("B品見積取置数量") = Nz(vRS_Inventory("B品見積取置数量"), 0) - Nz(vRS("見積引当合計数量"), 0)

			End If

		End If

		vRS_Inventory("最終更新日") = Now()
		vRS_Inventory("最終更新者コード") = UPDATEABLE_STAFF_CD

		vRS_Inventory.Update

	End If

	vRS_Inventory.Close

	vRS.MoveNext

Loop

vRS.close

Set vRS_Inventory = Nothing
Set vRS = Nothing

update_Inventory = True

End function

'========================================================================
'
'	Function	EmaxのWeb変更受注の登録
'	Note		レコードが存在すれば更新
'
'========================================================================
Function insert_WebUpdateOrder()

insert_WebUpdateOrder = False

Dim vSQL
Dim vRS

vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      受注番号 "
vSQL = vSQL & "    , 変更開始日 "
vSQL = vSQL & "    , 変更終了日 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "Web変更受注 "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    受注番号 = " & wOrderNo & " "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then

	' レコード無しの為、登録
	vRS.AddNew

	vRS("受注番号") = wOrderNo
	vRS("変更開始日") = Now()
	vRS("変更終了日") = Now()

Else

	' レコード有りの為、更新
	vRS("変更開始日") = Now()
	vRS("変更終了日") = Now()

End If

vRS.Update

vRS.Close
Set vRS = Nothing

insert_WebUpdateOrder = True

End function

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
call getCntlMst("Web", "受注", "受注変更可能時間帯", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

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
'	Function	Nz 関数
'
'========================================================================
Function Nz(pvarValue, pvarDefaultValue)

If IsNull(pvarValue) Then
	Nz = pvarDefaultValue
Else
	Nz = pvarValue
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
'
'	Function	キャンセルメール送信
'
'========================================================================
Function send_cancelMail()

send_cancelMail = False

Dim vstrItemChar1
Dim vstrItemChar2
Dim vdblItemNum1
Dim vdblItemNum2
Dim vdatItemDate1
Dim vdatItemDate2
Dim vEMailAddrFrom
Dim vEMailAddrTo
Dim vEMailAddrBCC
Dim vobjCBOMessage
Dim vSubject
Dim vBody
Dim vRS
Dim vRS_Item
Dim vItemName
Dim vUnitPrice
Dim vTaxRate


'--- コントロールマスタから消費税率取得
call getCntlMst("共通", "消費税率", "1", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

vTaxRate = Clng(vdblItemNum1)

'--- コントロールマスタから送信元(From)アドレス取得
call getCntlMst("共通", "送信先Email", "Web受注通知", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

'--- 送信元(From)
vEMailAddrFrom = vstrItemChar1

'--- コントロールマスタからBCCアドレス取得
call getCntlMst("共通", "送信先Email", "ShopBCC", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

'--- BCC
vEMailAddrBCC = vstrItemChar1

'--- subject
vSubject = "サウンドハウス　ご注文キャンセル確認メール（自動配信）[" & wUserID & "/Web-Emax/Web受注キャンセル]"

'--- 顧客情報取得
Set vRS = get_customer()
If vRS.EOF Then
	' 顧客情報無し
	Exit Function
End If

'--- body
vBody = ""
vBody = vBody & "サウンドハウス・オンラインショップをご利用頂き誠にありがとうございます。" & vbNewLine
vBody = vBody & "下記のご注文がキャンセルされました。" & vbNewLine
vBody = vBody & vbNewLine
vBody = vBody & "＝＝＝＝＝＝＝＝＝　ご注文（キャンセル）＝＝＝＝＝＝＝＝＝" & vbNewLine
vBody = vBody & vRS("顧客名") & " 様" & vbNewLine
vBody = vBody & "住所： 〒" & vRS("顧客郵便番号") & " " & vRS("顧客都道府県") & vRS("顧客住所") & vbNewLine
vBody = vBody & "電話番号： " & vRS("顧客電話番号") & vbNewLine
vBody = vBody & "お客様番号： " & wUserID & vbNewLine

vRS.Close

'--- 受注情報取得
Set vRS = get_orderInfo()
If vRS.EOF Then
	' 受注情報無し
	Exit Function
End If

'--- 送信先(To)
vEMailAddrTo = vRS("顧客E_mail")

vBody = vBody & "見積日付： " & formatDateYYYYMMDD(vRS("見積日")) & vbNewLine
vBody = vBody & "見積番号： " & wOrderNo & vbNewLine
vBody = vBody & "お支払方法： " & vRS("支払方法") & vbNewLine
vBody = vBody & vbNewLine

vBody = vBody & "－－－－－－－－－　詳　　細　－－－－－－－－－" & vbNewLine

Do Until vRS.EOF

	'--- 商品情報取出し
	Set vRS_Item = get_itemInfo(vRS("メーカーコード") & "", vRS("商品コード") & "")
	If vRS_Item.EOF = False Then
		vItemName = vRS_Item("メーカー名") & " " & vRS_Item("商品名")
	Else
		' 商品情報無し
		vItemName = ""
	End If
	vRS_Item.Close

	'--- 単価計算
	vUnitPrice = calcPrice(vRS("受注単価"), vTaxRate)

	vBody = vBody & "商品名： " & vItemName & vbNewLine
	vBody = vBody & "数量： " & FormatNumber(vRS("受注数量"), 0) & vbNewLine
	vBody = vBody & "単価(税込)： \" & FormatNumber(vUnitPrice, 0) & vbNewLine
	vBody = vBody & "金額(税込)： \" & FormatNumber(vUnitPrice * vRS("受注数量"), 0) & vbNewLine
	vBody = vBody & vbNewLine

	vRS.MoveNext

Loop

Set vRS_Item = Nothing

vRS.Close

Set vRS = Nothing

'--- コントロールマスタからトレーラー用文字列取得
call getCntlMst("Web", "Email", "トレーラ", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

vBody = vBody & vstrItemChar1


'--- メール送信
Set vobjCBOMessage = Server.CreateObject("CDO.Message")

vobjCBOMessage.From = vEMailAddrFrom
vobjCBOMessage.To = vEMailAddrTo
vobjCBOMessage.BCC = vEMailAddrBCC
vobjCBOMessage.Subject = vSubject
vobjCBOMessage.TextBody = vBody
vobjCBOMessage.BodyPart.Charset = "iso-2022-jp"

'--- メールサーバー指定 (不要であれば以下4行コメントアウト)
'vobjCBOMessage.Configuration.Fields.Item(g_ItemSMTPSendusing) = g_SMTPSendusing
'vobjCBOMessage.Configuration.Fields.Item(g_ItemSMTPServer) = g_SMTPServer
'vobjCBOMessage.Configuration.Fields.Item(g_ItemSMTPServerPort) = g_SMTPServerPort
'vobjCBOMessage.Configuration.Fields.Update

vobjCBOMessage.Send

Set vobjCBOMessage = Nothing

send_cancelMail = True

End function

'========================================================================
%>
