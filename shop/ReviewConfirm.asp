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
'	レビュー内容確認ページ
'
'	レビュー編集ページの｢内容を確認する」ボタン押下から遷移される。
'	顧客番号をSession("userID")から取出し、引数として用いる。
'
'	クエリ文字列 ?item=628^NT5^^&WriteReview=Y
'	628^ ... メーカーコード
'	NT5^ ... 商品コード
'	^    ... 色
'	^    ... 規格
'	WriteReview ... レビュー記述
'
'	HTTPSでないとエラー
'	ログインしていないとエラー
'	ログインしていれば、Session("userID")に顧客番号がセットされている。
'	Session("userID")が空文字の時はエラー　｢ログインしてください。｣
'	Session("userID")で顧客情報が取出せなければエラー　｢ログインしてください。｣
'	エラーメッセージをセットしLogin.aspへRedirect
'
'変更履歴
'2013/05/07 GV #1507 新規作成(ReviewWrite.aspをベースに新規作成)
'
'========================================================================
'On Error Resume Next
Const THIS_PAGE_NAME = "ReviewWrite.asp"

Dim maker_cd
Dim product_cd
Dim iro
Dim kikaku
Dim item
Dim item_list()
Dim item_cnt

Dim ReviewAll
Dim WriteReview
Dim OrderNo

Dim wMakerName
Dim wMakerNameNoKana
Dim wProductName
Dim wPrefecture
Dim wHandleName
Dim wReviewBody				'2013/05/07 GV Add レビュー文章

Dim Connection
Dim RS

Dim wTitle
Dim wMainProdPic        '2011/11/22 an add

Dim wSQL
Dim wHTML
Dim wMsg
Dim wErrDesc   '2011/08/01 an add

Dim wDispMsg
Dim wErrMsg
Dim UserID
Dim wNotLogin
Dim vEditMode
Dim oProductData		'商品情報
Dim oTotalReviewData	'レビュー（総評）
Dim oReviewData			'レビュー（個別）
Dim oCustomerData		'顧客情報
Dim oOrderData			'受注情報
Dim wReviewProduct		'レビュー中の商品
Dim wEvaluteOnpu		'総評価音符
Dim wBreadCrumbs		'パン屑リスト
Dim isExistReview		'レビューが存在している
Dim wRating				'評価
Dim wReviewTitle		'レビュータイトル
Dim wReviewDate			'レビュー投稿日
Dim wReviewBodyBr		'BRタグで改行したレビュー内容
Dim wReviewSelected		'選択した評価にselectedをつける
Dim wInvalidate			'バリデーションエラー
Dim NgFlg				'NGフラグ
Dim Mode				'処理モード(1...save,-1...delete)
'=======================================================================
'	受け渡し情報取り出し & 初期設定
'=======================================================================
Response.buffer = true
%>
<!--#include file="ReviewFunc.inc"-->
<%

'---- Session変数
wDispMsg = Session("DispMsg")
Session("DispMsg") = ""
wErrMsg = Session("ErrMsg")
Session("ErrMsg") = ""

UserID = Session("userID")
wNotLogin = False				' 初期状態はログインしている事を前提とする


'---- Get input data
maker_cd    = ReplaceInput(Trim(Request("maker_cd")))
product_cd  = ReplaceInput(Trim(Request("product_cd")))
iro         = ReplaceInput(Trim(Request("iro")))
kikaku      = ReplaceInput(Trim(Request("kikaku")))
item        = ReplaceInput(Trim(Request("item")))

wHandleName   = ReplaceInput(Trim(Request("HandleName")))
wReviewTitle  = ReplaceInput(Trim(Request("Title")))
wReviewBody   = ReplaceInput(Trim(Request("Review")))
wReviewBodyBr = Replace(wReviewBody, vbCrLf, "<BR>")
wRating       = ReplaceInput(Trim(Request("Rating")))
Mode          = ReplaceInput(Request("Mode"))

wPrefecture = ""
vEditMode    = ""


If Trim(Request("parm")) <> "" Then
	item = ReplaceInput(Trim(Request("parm")))
End If

' 商品に関するクエリ文字列
If item <> "" Then
	item_cnt = cf_unstring(item, item_list, "^")
	maker_cd = item_list(0)
	product_cd = item_list(1)
	If item_cnt > 2 Then
		iro = item_list(2)
		If item_cnt > 3 Then
			kikaku = item_list(3)
		End If
	End If
End If

'----商品レビュー用パラメータ
ReviewAll = ReplaceInput(Request("ReviewAll"))
WriteReview = ReplaceInput(UCase(Request("WriteReview")))

OrderNo = ReplaceInput(Request("OrderNo"))
If (OrderNo <> "" and isNumeric(OrderNo) = false) OR OrderNo = "" Then
	OrderNo = 0
End If

NgFlg = false


'=======================================================================
'	Execute main
'=======================================================================
'---- DB接続
Call ReviewFunc_ConnectDb()

'---- メイン処理
Call main()

'---- DB切断
Call ReviewFunc_CloseDb()

'---- エラーメッセージをセッションデータに登録   ' member系の他のページ処理にならう
If Err.Description <> "" Then
	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- ログインしていない場合はログインページへ
If wNotLogin = True Then
	Session("msg") = wMsg
	Server.Transfer "../shop/Login.asp"
End If

'---- データなし等のエラーがある場合、エラーページへ
If NgFlg = True Then
	Session("msg") = wMsg
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

	' セッションからユーザ情報を取得できなかった場合、ログインされていないとしてエラーとする
	If UserID = "" Then
		'---- オブジェクトの開放
		Call DeallocObject()

		wNotLogin = True		' ログインされていない
		wMsg = "ログインしてください。"
		Exit Function
	End If

	' 顧客情報オブジェクトを取得
	Set oCustomerData = ReviewFunc_GetCustomer()

	If oCustomerData.EOF = True Then
		'---- オブジェクトの開放
		Call DeallocObject()

		'--- Session("userID")で顧客情報が取出せなければエラー　｢ログインしてください。｣
		wNotLogin = True		' ログインされていない
		wMsg = "ログインしてください。"
		Exit Function
	ElseIf (CLng(oCustomerData("購入回数")) < 1) Then
		'---- オブジェクトの開放
		Call ReviewFunc_FreeObject()

		wMsg = "レビューは購入された方が投稿できます。"
		NgFlg = True
		Exit Function
	End If

	If (IsNumeric(Mode) = false) Then
		'---- オブジェクトの開放
		Call ReviewFunc_FreeObject()

		wMsg = ""
		NgFlg = True
		Exit Function
	End If

	'---- 商品情報取り出し
	Set oProductData = ReviewFunc_GetProduct()

	' 商品情報が空の場合
	If IsObject(oProductData) = false Then
		'---- オブジェクトの開放
		Call ReviewFunc_FreeObject()

		wMsg = "商品情報が見つかりませんでした。"
		NgFlg = True
		Exit Function
	Else
		'---- 受注情報の取得
		oOrderData = ReviewFunc_GetOrder(UserID, maker_cd, product_cd)

		'受注情報がない場合、レビューを投稿させない
		If (IsObject(oOrderData) = false) Then
			'---- オブジェクトの開放
			Call ReviewFunc_FreeObject()

			wMsg = "レビューを投稿することができません。"
			NgFlg = True
			Exit Function
		End If

		'---- 商品レビュー情報取り出し
		Set oTotalReviewData =  ReviewFunc_GetTotalReview()		' 総評
		Set oReviewData      =  ReviewFunc_GetReview(null)		' 個別

		wPrefecture = oCustomerData("顧客都道府県")

		'---- レビューデータがある場合
		If oReviewData.EOF = false Then
			If (CInt(Mode) <> -1) Then
				vEditMode = "編集"
				isExistReview = true
			Else
				vEditMode = "削除"
			End If

			wReviewDate = FormatDateTime(oReviewData("投稿日"), 1)
			'DebugEcho("レビューデータあり")
		Else
			vEditMode = "投稿"
			wReviewDate = FormatDateTime(Now, 1)
			'DebugEcho("レビューデータなし")
		End If
	End If

	'---- パン屑リスト
	'wBreadCrumbs = ReviewFunc_CreateBreadCrumbsHTML()
	wBreadCrumbs = "商品レビュー"

	'---- 商品画像
	wReviewProduct = ReviewFunc_CreateReviewProductPictureHTML()

	'---- ユーザがつけた評価の音符アイコン
	wEvaluteOnpu = ReviewFunc_CreateEvaluteOnpu(wRating, false)

	'---- ユーザがつけた評価の音符selected
	wReviewSelected = ReviewFunc_EvaluteSelectedArray(6, wRating)

	'---- バリデーション
	If (CInt(Mode) <> -1) Then
		wInvalidate = ReviewFunc_Validate()
	End If

	'---- オブジェクトの開放
	Call ReviewFunc_FreeObject()

End Function	' End of main()
'========================================================================
%>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>商品レビュー｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/review.css?201309xx" type="text/css">
<% Call ReviewFunc_JsDelete_onClick()%>
<% Call ReviewFunc_JsReview_onClick()%>
<script type="text/javascript">
//
// ====== 	Function:	review_onSubmit
//
function review_back(mode){
	if (mode == 1) {
		document.f_review.action = "ReviewWrite.asp?Item=<%= Server.URLEncode(item)%>";
		document.f_review.Mode.value = 9;
	}
	document.f_review.submit();
}
</script>
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
      <li class="now"><%=wBreadCrumbs%></li>
    </ul>
  </div></div></div>
  <h1 class="title">商品レビュー</h1>

  <div id="review_main">
<%
	If (CInt(Mode) <> -1) Then
		'エラーがある場合
		If (IsNull(wInvalidate) <> True) Then
			'---- 入力フォームの呼び出し
			Call ReviewFunc_CreateReviewForm()
		Else
			Call ReviewFunc_CreateReviewConfirmForm()
		End If
	Else
		Call ReviewFunc_CreateReviewConfirmForm()
	End If
%>
  </div>

  <div id="review_side">
    <div class='review_side_inner01'><div class='review_side_inner02'>
      <div class='review_side_inner_box'>
        <h4 class='review_sub'>レビュー中の商品</h4>
        <%=wReviewProduct%>
      </div>
    </div></div>
  </div>

<!--/#contents --></div>

  <div id="globalSide">
  <!--#include file="../Navi/NaviSide.inc"-->
  <!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>