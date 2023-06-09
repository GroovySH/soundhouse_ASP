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
'	レビュー編集ページ
'
'	商品詳細ページの｢この商品のレビューを書く」リンクから呼び出される。
'	顧客番号をSession("userID")から取出し、引数として用いる。
'
'	クエリ文字列 ?item=628^NT5^^
'	628^ ... メーカーコード
'	NT5^ ... 商品コード
'	^    ... 色
'	^    ... 規格
'
'	HTTPSでないとエラー
'	ログインしていないとエラー
'	ログインしていれば、Session("userID")に顧客番号がセットされている。
'	Session("userID")が空文字の時はログインページにリダイレクト
'	エラーメッセージをセットしLogin.aspへRedirect
'
'変更履歴
'2013/05/07 GV #1507 新規作成(ProductDetail.aspをベースに新規作成)
'
'========================================================================
On Error Resume Next

'キャッシュなし
Response.Expires = -1
Response.AddHeader "Cache-Control", "No-Cache"
Response.AddHeader "Pragma", "No-Cache"

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
Dim iShop

'Dim wMakerName
'Dim wMakerNameNoKana
Dim wProductName
'Dim wCategoryCode
'Dim wTitleWithLink
'Dim wKoukeiMakerCd
'Dim wKoukeiProductCd
Dim wLargeCategoryCd
Dim wMidCategoryCd
Dim wCanWriteReviewFl
Dim wPrefecture
Dim wHandleName
Dim wIroKikakuSelectMsg
Dim wLargeCategoryName		'2010/08/23 an add
Dim wMidCategoryName		'2010/08/23 an add
Dim wCategoryName			'2010/08/23 an add
Dim wTokucho				'2010/08/23 an add
Dim wFreeShippingFlag		' 2011/02/18 GV Add
Dim s_category_cd        	'2011/09/09 an add For NaviLeftShop
Dim wOptionPartsTitleFlag	'2012/08/29 ok Add
Dim wReviewBody				'2013/03/26 GV Add レビュー文章
Dim wMode

Dim wIroKikakuCombo

Dim wPictureHTML
Dim wKanrenLinkHTML
Dim wTokuchoHTML
Dim wSpecHTML
Dim wOptionHtml
Dim wPartsHtml
Dim wReviewHTML
Dim wCampaignHTML	' 2013/01/30 GV Add

Dim wProductHTML
Dim wHyoukaHTML
Dim wCartHTML
Dim wSeriesHTML
Dim wRecommendHTML
Dim wRecommendBuyHTML	' 2012/04/10 GV Add
Dim wSubItemHTML        ' 2012/05/01 GV Add
Dim wViewHTML		' 2012/07/10 GV Add

Dim Connection
Dim RS

Dim wTitle
Dim wSalesTaxRate
Dim wProdTermFl
Dim wPrice
Dim wOptionPartsFl
Dim wIroKikakuSelectedFl
Dim wNoData
Dim wIroKikakuFl       '2010/11/26 an add
Dim wIroKikakuZaikoFl  '2010/11/26 an add
Dim wIroKikakuHacchuuFl  '2011/06/09 an add
Dim wMainProdPic        '2011/11/22 an add

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim wNum		'数字画像

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

Dim wSeriesCd		'2012/10/30 nt Add
Dim wContentsHTML	'2012/10/30 nt Add
Dim wDispMsg
Dim wErrMsg
Dim UserID
Dim wNotLogin
Dim vEditMode
Dim wBHinFl				' B品フラグ
Dim oProductData		'商品情報
Dim oTotalReviewData	'レビュー（総評）
Dim oReviewData			'レビュー（個別）
Dim oCustomerData		'顧客情報
Dim wTotalEvaluteOnpu	'評価音符
Dim wBreadCrumbs		'パン屑リスト
Dim isExistReview		'レビューが存在している
Dim wReviewTitle		'レビュータイトル
Dim wReviewSelected		'選択した評価にselectedをつける
Dim wRating				'評価
Dim wInvalidate			'バリデーション
Dim NgFlg				'NGフラグ


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

isExistReview = false

'---- Get input data
maker_cd    = ReplaceInput(Trim(Request("maker_cd")))
product_cd  = ReplaceInput(Trim(Request("product_cd")))
iro         = ReplaceInput(Trim(Request("iro")))
kikaku      = ReplaceInput(Trim(Request("kikaku")))
item        = ReplaceInput(Trim(Request("item")))

wHandleName   = ReplaceInput(Trim(Request("HandleName")))
wReviewTitle  = ReplaceInput(Trim(Request("Title")))
wReviewBody   = ReplaceInput(Trim(Request("Review")))
wRating       = ReplaceInput(Trim(Request("Rating")))
wMode         = ReplaceInput(Trim(Request("Mode")))

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
Call ReviewFunc_ConnectDb()

Call main()

Call ReviewFunc_CloseDb()

'---- 該当商品なしのとき
if wNoData = "Y" then
	Response.Redirect "SearchNotFound.asp"
end if

'---- エラーメッセージをセッションデータに登録   ' member系の他のページ処理にならう
If Err.Description <> "" Then
	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- ログインしていない場合はログインページへ
If wNotLogin = True Then
	If  gPhoneType = "SP" Then
		Response.Redirect g_HTTPS & "sp/shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "sp/shop/ReviewWrite.asp?Item=" & Server.URLEncode(item)
	Else
		Response.Redirect g_HTTPS & "shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "shop/ReviewWrite.asp?Item=" & Server.URLEncode(item)
	End If
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
	ElseIf (oCustomerData("購入回数") < 1) Then
		'---- オブジェクトの開放
		Call ReviewFunc_FreeObject()
		wMsg = "購入履歴がない為、書き込めません。"
		NgFlg = True
		Exit Function
	End If

	'---- 商品情報取り出し
	Set oProductData = ReviewFunc_GetProduct()

	' 商品情報が空の場合
'	If IsObject(oProductData) = false Then
	If oProductData.EOF = true Then
		'---- オブジェクトの開放
		Call ReviewFunc_FreeObject()

		wNoData = "Y"
		wMsg = "商品情報が見つかりませんでした。"
		NgFlg = True
		Exit Function
	Else
		'---- 商品レビュー情報取り出し
		Set oTotalReviewData =  ReviewFunc_GetTotalReview()		' 総評

		wPrefecture = oCustomerData("顧客都道府県")

		Set oReviewData      =  ReviewFunc_GetReview(null)		' 個別

		If (wMode = "9") Then
			vEditMode = "編集"
			wReviewSelected = ReviewFunc_EvaluteSelectedArray(6, wRating)
			isExistReview = true

			'---- レビューデータがある場合
			If oReviewData.EOF = false Then
				isExistReview = true
			End If
		Else
			'---- HTML で使う変数の調整
			wPrefecture = ""
			wHandleName = ""
			wReviewTitle = ""
			wReviewBody  = ""
			vEditMode    = ""

'			Set oReviewData      =  ReviewFunc_GetReview(null)		' 個別

			'---- レビューデータがある場合
			If oReviewData.EOF = false Then
				vEditMode = "編集"
				isExistReview = true

				'レビューデータに名前が入っている場合、そちらを優先
'				If (IsNull(oReviewData("名前")) = false) Then
				If (Trim(oReviewData("名前")) <> "") Then
					wHandleName  = Trim(oReviewData("名前"))
				Else
					wHandleName = Trim(oCustomerData("ハンドルネーム"))
				End If

				wReviewTitle = Trim(oReviewData("タイトル"))
				wReviewBody  = Trim(oReviewData("レビュー内容"))
				wReviewSelected = ReviewFunc_EvaluteSelectedArray(6, CInt(oReviewData("評価")))

			Else
				vEditMode = "投稿"
				wHandleName = Trim(oCustomerData("ハンドルネーム"))
				wReviewSelected = ReviewFunc_EvaluteSelectedArray(6, 0)
			End If
		End If
	End If

	'---- パン屑リスト
	wBreadCrumbs = "商品レビュー"

	'---- 商品画像
	wTotalEvaluteOnpu = ReviewFunc_CreateReviewProductPictureHTML()

	'---- オブジェクトの開放
	Call ReviewFunc_FreeObject()

End Function	' End of main()
'========================================================================
%>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="ShIft_JIS">
<title>商品レビュー｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/review.css?201309xx" type="text/css">
<% Call ReviewFunc_JsDelete_onClick()%>
<% Call ReviewFunc_JsReview_onClick()%>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gIf" alt="ここから本文です"></a></span>
  <!-- コンテンツstart -->
  <div id="globalContents">
  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gIf" alt="HOME"></a></p>
    <ul id="path">
      <li class="now"><%=wBreadCrumbs%></li>
    </ul>
  </div></div></div>
  <h1 class="title">商品レビュー</h1>

  <div id="review_main">
<%
	'---- 入力フォームの呼び出し
	Call ReviewFunc_CreateReviewForm()
%>
  </div>

  <div id="review_side">
    <div class="review_side_inner01"><div class="review_side_inner02">
      <div class="review_side_inner_box">
        <h4 class="review_sub">レビュー中の商品</h4>
        <%=wTotalEvaluteOnpu%>
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
