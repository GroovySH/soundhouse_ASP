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
'	商品レビュー 登録
'
'更新履歴
'2007/04/05 ログインさえしていれば誰でも登録できるように変更
'2007/04/25 ハンドルネームを顧客へ登録
'2007/04/27 メール送信メーカーコードをメーカー名に変更
'2007/08/23 出荷メールからのリンクでログインしていないときは、OrderNoからUserIDを取得するように変更
'2008/05/23 入力データチェック強化（LEFT, Numeric, EOF他)
'2009/04/30 エラー時にerror.aspへ移動
'2011/08/01 an #1087 Error.aspログ出力対応
'2012/01/20 an SELECT文へLACクエリー案を適用、メール送信を共通関数利用に変更
'2012/07/30 if-web リニューアルレイアウト調整
'2013/05/07 GV #1507 レビュー再編集機能
'
'========================================================================

On Error Resume Next

'Dim userID
Dim UserID
Dim msg

Dim maker_cd
Dim product_cd
Dim iro
Dim kikaku

Dim Rating
Dim Title
Dim HandleName
Dim Review
Dim OrderNo
Dim Item
Dim MakerCd
Dim ProductCd

Dim wMakerProduct
Dim wMakerName

Dim item_list()
Dim item_cnt

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim Connection
Dim RS

Dim w_sql
Dim w_html
Dim w_msg
Dim wErrDesc   '2011/08/01 an add
Dim wMsg

'2013/05/07 GV #1507 add start
Dim Mode				'処理モード(1...save,-1...delete)
Dim ReviewID			'レビューID
Dim oProductData		'商品情報
Dim oTotalReviewData	'レビュー（総評）
Dim oReviewData			'レビュー（個別）
Dim oCustomerData		'顧客情報
Dim oOrderData			'受注情報
Dim wReviewDate			'レビュー日付
Dim urlEncItem			'URLエンコードしたitem
Dim backUrl				'戻り先URL
'2013/05/07 GV #1507 add end
'========================================================================

Response.buffer = true
%>
<!--#include file="ReviewFunc.inc"-->
<%

'---- UserID 取り出し
'userID = Session("userID")
UserID = Session("userID")

'---- 呼び出し元からのデータ取り出し
Rating = ReplaceInput(Request("Rating"))
Title = ReplaceInput(Left(Request("Title"), 50))
HandleName = ReplaceInput(Left(Request("HandleName"), 30))
Review = ReplaceInput(Left(Request("Review"), 1000))
OrderNo = ReplaceInput(Request("OrderNo"))
Item = ReplaceInput(Request("Item"))
Mode = ReplaceInput(Request("Mode"))
ReviewID   = ReplaceInput(Request("ReviewID"))

if isNumeric(Rating) = false then
	Rating = 3
end if

if IsNumeric(OrderNo) = true then
	OrderNo = Clng(ReplaceInput(Request("OrderNo")))
else
	OrderNo = 0
end if

' 商品に関するクエリ文字列
If Item <> "" Then
	item_cnt = cf_unstring(Item, item_list, "^")
	maker_cd = item_list(0)
	product_cd = item_list(1)
	If item_cnt > 2 Then
		iro = item_list(2)
		If item_cnt > 3 Then
			kikaku = item_list(3)
		End If
	End If
End If

'=======================================================================
'	Execute main
'=======================================================================
'---- DB接続
Call ReviewFunc_ConnectDb()

Call main()

'2013/05/07 GV #1507 add start
'---- DB切断
Call ReviewFunc_CloseDb()
'2013/05/07 GV #1507 add end

'---- エラーメッセージをセッションデータに登録   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ReviewStore.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'========================================================================
'
'	Function	main proc
'
'========================================================================
'
Function main()

	'2013/05/07 GV #1507 modified start
	' セッションからユーザ情報を取得できなかった場合、ログインされていないとしてエラーとする
	If UserID = "" Then
		'---- オブジェクトの開放
		Call ReviewFunc_FreeObject()

		wNotLogin = True		' ログインされていない
		wMsg = "ログインしてください。"
		Exit Function
	End If

	' 顧客情報オブジェクトを取得
	Set oCustomerData = ReviewFunc_GetCustomer()

	If oCustomerData.EOF = True Then
		'---- オブジェクトの開放
		Call ReviewFunc_FreeObject()

		'--- Session("userID")で顧客情報が取出せなければエラー　｢ログインしてください。｣
		wNotLogin = True		' ログインされていない
		wMsg = "ログインしてください。"
		Exit Function
	End If

'2013/05/07 GV #1507 modified start
'if UserID = "" AND OrderNo <> "" then
	'---- UserID取り出し
'	w_sql = ""
'	w_sql = w_sql & "SELECT 顧客番号"
'	w_sql = w_sql & "  FROM Web受注 WITH (NOLOCK)"  '2012/01/20 an mod
'	w_sql = w_sql & " WHERE 受注番号 = " & OrderNo
'
'	Set RS = Server.CreateObject("ADODB.Recordset")
'	RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic
'
'	if RS.EOF = false then
'		UserID = RS("顧客番号")
'	else
'		w_msg = "<font color='#ff0000'>レビューは登録できません｡</font>"
'		exit function
'	end if
'
'	RS.Close
'end if
'2013/05/07 GV #1507 modified end

'2013/05/07 GV #1507 add start
	'---- 商品情報取り出し
	Set oProductData = ReviewFunc_GetProduct()

	' 商品情報が空の場合
'	If IsObject(oProductData) = false Then
	If (oProductData.EOF = true) Then
		'---- オブジェクトの開放
		Call ReviewFunc_FreeObject()

		wMsg = "商品情報が見つかりませんでした。"
		NgFlg = True
		Exit Function
	Else
		'---- 商品レビュー情報取り出し
		Set oTotalReviewData =  ReviewFunc_GetTotalReview()		' 総評
		Set oReviewData      =  ReviewFunc_GetReview(NULL)		' 個別

		'---- レビューデータがある場合
		If oReviewData.EOF = false Then
			wReviewDate = oReviewData("投稿日")
		Else
			wReviewDate = Now
		End If
	End If
'2013/05/07 GV #1507 add end

'2013/05/07 GV #1507 modified start
'商品情報の取得と同時にメーカー名を取得しているので、以下の処理をコメントアウト
'wMakerProduct = Split(Item, "^")
'
'---- メーカー名取り出し
'w_sql = ""
'w_sql = w_sql & "SELECT メーカー名"
'w_sql = w_sql & "  FROM メーカー WITH (NOLOCK)"  '2012/01/20 an mod
'w_sql = w_sql & " WHERE メーカーコード = '" & wMakerProduct(0) & "'"
'
'Set RS = Server.CreateObject("ADODB.Recordset")
'RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic
'
'if RS.EOF = false then
'	wMakerName = RS("メーカー名")
'else
'	wMakerName = wMakerProduct(0)
'end if
'RS.Close
'2013/05/07 GV #1507 modified end

'2013/05/07 GV #1507 modified start
'---- 商品レビュー結果登録
w_sql = ""
w_sql = w_sql & "SELECT *"
w_sql = w_sql & "  FROM 商品レビュー"
'w_sql = w_sql & " WHERE メーカーコード = '" & wMakerProduct(0) & "'"
'w_sql = w_sql & "   AND 商品コード = '" & wMakerProduct(1) & "'"
'w_sql = w_sql & "   AND 顧客番号 = " & UserID
w_sql = w_sql & " WHERE メーカーコード = '" & oProductData("メーカーコード") & "'"
w_sql = w_sql & "   AND 商品コード = '" & oProductData("商品コード") & "'"
w_sql = w_sql & "   AND 顧客番号 = " & UserID
'
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic
'
'if RS.EOF = false then
'	w_msg = "<font color='#ff0000'>レビューは1回のみ投稿できます｡　既にお客様からのレビューは投稿されていますので再投稿できません｡</font>"
'	exit function
'end if


'w_msg = "<b>商品レビュー登録ありがとうございました。</b>"

'---- insert 商品レビュー
'2013/05/07 GV #1507 modified start
'RS.AddNew
'RS("メーカーコード") = wMakerProduct(0)
'RS("商品コード") = wMakerProduct(1)
'RS("投稿日") = now()
'RS("顧客番号") = UserID
'RS("評価") = Rating
'RS("タイトル") = Title
'RS("名前") = HandleName
'RS("レビュー内容") = Review
'RS("参考数") = 0
'RS("不参考数") = 0

'RS.Update
'RS.close

	'DBに保存されているレビューデータが存在し、削除モードの場合は、削除
	If ((oReviewData.EOF = False) And (CInt(Mode) = -1)) Then
		w_msg = "<p>商品レビューを削除しました。</p>"
		RS.Delete
		RS.close
	Else
		'新規登録の場合
		If oReviewData.EOF = True Then
			RS.AddNew
		End If

		w_msg = "<p>商品レビューを登録しました。</p>"

		RS("メーカーコード") = oProductData("メーカーコード")
		RS("商品コード")     = oProductData("商品コード")
		RS("投稿日")         = wReviewDate
		RS("顧客番号")       = UserID
		RS("評価")           = Rating
		RS("タイトル")       = Title
		RS("名前")           = HandleName
		RS("レビュー内容")   = Review
		RS("参考数")         = 0
		RS("不参考数")       = 0

		RS.Update
		RS.close

		'---- ハンドルネーム登録
		w_sql = ""
		w_sql = w_sql & "SELECT ハンドルネーム"
		w_sql = w_sql & "     , 最終更新日"
		w_sql = w_sql & "  FROM Web顧客"
		'w_sql = w_sql & " WHERE 顧客番号 = " & UserID
		w_sql = w_sql & " WHERE 顧客番号 = " & UserID
		w_sql = w_sql & "   AND (ハンドルネーム = '' OR ハンドルネーム IS NULL)"

		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

		if RS.EOF = false then
			RS("ハンドルネーム") = HandleName
			RS("最終更新日") = Now()
			RS.update
			RS.close
		end if

'---- 登録メール送信
call sendMail()
	End If	'2013/05/07 GV #1507 modified end

	backUrl = g_HTTP & "shop/ProductDetail.asp?Item=" & item	'2013/05/07 GV #1507 add

End function

'========================================================================
'
'	Function	メール送信
'
'========================================================================
'
Function sendMail()

Dim v_body
'Dim OBJ_NewMail  '2012/01/20 an del

'2013/05/07 GV #1507 add start
Dim vSuffix
Dim vRS
Dim vSql
'2013/05/07 GV #1507 add end

'---- wItemChar1 = To, wItemChar2 = From
call getCntlMst("共通","送信先Email","商品レビュー", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

'2013/05/07 GV #1507 add start
vSuffix  = ""

'レビューが既にある場合、(編集)をつける
If (oReviewData.EOF = false) Then
	vSuffix = vSuffix & "(編集)"

	'ショップコメントが有る場合
	If (oReviewData("ショップコメント日") <> "") Then
		vSuffix = vSuffix & "ショップコメントあり"
	End If
End If

'v_body = "商品レビュー" & vbNewLine & vbNewLine
v_body = "商品レビュー" & vSuffix & vbNewLine & vbNewLine

'---- 商品レビュー結果登録
vSql = ""
vSql = vSql & "SELECT ID"
vSql = vSql & "  FROM 商品レビュー"
vSql = vSql & " WHERE メーカーコード = '" & oProductData("メーカーコード") & "'"
vSql = vSql & "   AND 商品コード = '" & oProductData("商品コード") & "'"
vSql = vSql & "   AND 顧客番号 = " & UserID

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSql, Connection, adOpenStatic, adLockOptimistic

If vRS.EOF = false Then
ReviewID = vRS("ID")
End If

vRS.close
'2013/05/07 GV #1507 add end

'2013/05/07 GV #1507 mod start
v_body = v_body & "ID　　　　　　：" & ReviewID & vbNewLine
'2013/05/07 GV #1507 mod end

v_body = v_body & "投稿日　　　　：" & now() & vbNewLine
v_body = v_body & "顧客番号　　　：" & UserID & vbNewLine & vbNewLine

'2013/05/07 GV #1507 mod start
'v_body = v_body & "メーカー名　　：" & wMakerName & vbNewLine
'v_body = v_body & "商品コード　　：" & wMakerProduct(1) & vbNewLine & vbNewLine
v_body = v_body & "メーカー名　　：" & oProductData("メーカー名") & vbNewLine
v_body = v_body & "商品コード　　：" & oProductData("商品コード") & vbNewLine & vbNewLine
'2013/05/07 GV #1507 mod end

v_body = v_body & "評価　　　　　：" & Rating & vbNewLine
v_body = v_body & "タイトル　　　：" & Title & vbNewLine
v_body = v_body & "ハンドルネーム：" & HandleName & vbNewLine
v_body = v_body & "レビュー内容　：" & vbNewLine & Review & vbNewLine

'2013/05/ GV #1507 mod start
'Call fSendEmail(wItemChar2, wItemChar1, "商品レビュー", v_body, "")    '2012/01/20 an add
Call fSendEmail(wItemChar2, wItemChar1, "商品レビュー" & vSuffix, v_body, "")
'2013/05/ GV #1507 mod end

'Set OBJ_NewMail = Server.CreateObject("CDO.Message") '2012/01/20 an del s
'
'OBJ_NewMail.from = wItemChar2
'OBJ_NewMail.to = wItemChar1
'
'OBJ_NewMail.subject = "商品レビュー"
'OBJ_NewMail.TextBody = v_body
'OBJ_NewMail.BodyPart.Charset = "iso-2022-jp"
'
'OBJ_NewMail.Send
'
'Set OBJ_NewMail = Nothing                             '2012/01/20 an del e

End function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>商品レビュー登録ありがとうございました｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">
  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now">商品レビュー</li>
    </ul>
  </div></div></div>

  <h1 class="title">商品レビュー</h1>
  <%=w_msg%>
  <p class="btnBox"><a href="<%= backUrl %>" class="opover">商品ページへ戻る</a></p>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>