<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
'�T�E���h�n�E�X
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
'	���r���[�ҏW�y�[�W
'
'	���i�ڍ׃y�[�W�̢���̏��i�̃��r���[�������v�����N����Ăяo�����B
'	�ڋq�ԍ���Session("userID")�����o���A�����Ƃ��ėp����B
'
'	�N�G�������� ?item=628^NT5^^
'	628^ ... ���[�J�[�R�[�h
'	NT5^ ... ���i�R�[�h
'	^    ... �F
'	^    ... �K�i
'
'	HTTPS�łȂ��ƃG���[
'	���O�C�����Ă��Ȃ��ƃG���[
'	���O�C�����Ă���΁ASession("userID")�Ɍڋq�ԍ����Z�b�g����Ă���B
'	Session("userID")���󕶎��̎��̓��O�C���y�[�W�Ƀ��_�C���N�g
'	�G���[���b�Z�[�W���Z�b�g��Login.asp��Redirect
'
'�ύX����
'2013/05/07 GV #1507 �V�K�쐬(ProductDetail.asp���x�[�X�ɐV�K�쐬)
'
'========================================================================
On Error Resume Next

'�L���b�V���Ȃ�
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
Dim wReviewBody				'2013/03/26 GV Add ���r���[����
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

Dim wNum		'�����摜

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
Dim wBHinFl				' B�i�t���O
Dim oProductData		'���i���
Dim oTotalReviewData	'���r���[�i���]�j
Dim oReviewData			'���r���[�i�ʁj
Dim oCustomerData		'�ڋq���
Dim wTotalEvaluteOnpu	'�]������
Dim wBreadCrumbs		'�p�������X�g
Dim isExistReview		'���r���[�����݂��Ă���
Dim wReviewTitle		'���r���[�^�C�g��
Dim wReviewSelected		'�I�������]����selected������
Dim wRating				'�]��
Dim wInvalidate			'�o���f�[�V����
Dim NgFlg				'NG�t���O


'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
Response.buffer = true
%>
<!--#include file="ReviewFunc.inc"-->
<%

'---- Session�ϐ�
wDispMsg = Session("DispMsg")
Session("DispMsg") = ""
wErrMsg = Session("ErrMsg")
Session("ErrMsg") = ""

UserID = Session("userID")
wNotLogin = False				' ������Ԃ̓��O�C�����Ă��鎖��O��Ƃ���

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

' ���i�Ɋւ���N�G��������
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

'----���i���r���[�p�p�����[�^
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

'---- �Y�����i�Ȃ��̂Ƃ�
if wNoData = "Y" then
	Response.Redirect "SearchNotFound.asp"
end if

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   ' member�n�̑��̃y�[�W�����ɂȂ炤
If Err.Description <> "" Then
	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- ���O�C�����Ă��Ȃ��ꍇ�̓��O�C���y�[�W��
If wNotLogin = True Then
	If  gPhoneType = "SP" Then
		Response.Redirect g_HTTPS & "sp/shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "sp/shop/ReviewWrite.asp?Item=" & Server.URLEncode(item)
	Else
		Response.Redirect g_HTTPS & "shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "shop/ReviewWrite.asp?Item=" & Server.URLEncode(item)
	End If
End If

'---- �f�[�^�Ȃ����̃G���[������ꍇ�A�G���[�y�[�W��
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

	' �Z�b�V�������烆�[�U�����擾�ł��Ȃ������ꍇ�A���O�C������Ă��Ȃ��Ƃ��ăG���[�Ƃ���
	If UserID = "" Then
		'---- �I�u�W�F�N�g�̊J��
		Call ReviewFunc_FreeObject()

		wNotLogin = True		' ���O�C������Ă��Ȃ�
		wMsg = "���O�C�����Ă��������B"
		Exit Function
	End If

	' �ڋq���I�u�W�F�N�g���擾
	Set oCustomerData = ReviewFunc_GetCustomer()

	If oCustomerData.EOF = True Then
		'---- �I�u�W�F�N�g�̊J��
		Call ReviewFunc_FreeObject()

		'--- Session("userID")�Ōڋq��񂪎�o���Ȃ���΃G���[�@����O�C�����Ă��������B�
		wNotLogin = True		' ���O�C������Ă��Ȃ�
		wMsg = "���O�C�����Ă��������B"
		Exit Function
	ElseIf (oCustomerData("�w����") < 1) Then
		'---- �I�u�W�F�N�g�̊J��
		Call ReviewFunc_FreeObject()
		wMsg = "�w���������Ȃ��ׁA�������߂܂���B"
		NgFlg = True
		Exit Function
	End If

	'---- ���i�����o��
	Set oProductData = ReviewFunc_GetProduct()

	' ���i��񂪋�̏ꍇ
'	If IsObject(oProductData) = false Then
	If oProductData.EOF = true Then
		'---- �I�u�W�F�N�g�̊J��
		Call ReviewFunc_FreeObject()

		wNoData = "Y"
		wMsg = "���i��񂪌�����܂���ł����B"
		NgFlg = True
		Exit Function
	Else
		'---- ���i���r���[�����o��
		Set oTotalReviewData =  ReviewFunc_GetTotalReview()		' ���]

		wPrefecture = oCustomerData("�ڋq�s���{��")

		Set oReviewData      =  ReviewFunc_GetReview(null)		' ��

		If (wMode = "9") Then
			vEditMode = "�ҏW"
			wReviewSelected = ReviewFunc_EvaluteSelectedArray(6, wRating)
			isExistReview = true

			'---- ���r���[�f�[�^������ꍇ
			If oReviewData.EOF = false Then
				isExistReview = true
			End If
		Else
			'---- HTML �Ŏg���ϐ��̒���
			wPrefecture = ""
			wHandleName = ""
			wReviewTitle = ""
			wReviewBody  = ""
			vEditMode    = ""

'			Set oReviewData      =  ReviewFunc_GetReview(null)		' ��

			'---- ���r���[�f�[�^������ꍇ
			If oReviewData.EOF = false Then
				vEditMode = "�ҏW"
				isExistReview = true

				'���r���[�f�[�^�ɖ��O�������Ă���ꍇ�A�������D��
'				If (IsNull(oReviewData("���O")) = false) Then
				If (Trim(oReviewData("���O")) <> "") Then
					wHandleName  = Trim(oReviewData("���O"))
				Else
					wHandleName = Trim(oCustomerData("�n���h���l�[��"))
				End If

				wReviewTitle = Trim(oReviewData("�^�C�g��"))
				wReviewBody  = Trim(oReviewData("���r���[���e"))
				wReviewSelected = ReviewFunc_EvaluteSelectedArray(6, CInt(oReviewData("�]��")))

			Else
				vEditMode = "���e"
				wHandleName = Trim(oCustomerData("�n���h���l�[��"))
				wReviewSelected = ReviewFunc_EvaluteSelectedArray(6, 0)
			End If
		End If
	End If

	'---- �p�������X�g
	wBreadCrumbs = "���i���r���["

	'---- ���i�摜
	wTotalEvaluteOnpu = ReviewFunc_CreateReviewProductPictureHTML()

	'---- �I�u�W�F�N�g�̊J��
	Call ReviewFunc_FreeObject()

End Function	' End of main()
'========================================================================
%>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="ShIft_JIS">
<title>���i���r���[�b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/review.css?201309xx" type="text/css">
<% Call ReviewFunc_JsDelete_onClick()%>
<% Call ReviewFunc_JsReview_onClick()%>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->
<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gIf" alt="��������{���ł�"></a></span>
  <!-- �R���e���cstart -->
  <div id="globalContents">
  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gIf" alt="HOME"></a></p>
    <ul id="path">
      <li class="now"><%=wBreadCrumbs%></li>
    </ul>
  </div></div></div>
  <h1 class="title">���i���r���[</h1>

  <div id="review_main">
<%
	'---- ���̓t�H�[���̌Ăяo��
	Call ReviewFunc_CreateReviewForm()
%>
  </div>

  <div id="review_side">
    <div class="review_side_inner01"><div class="review_side_inner02">
      <div class="review_side_inner_box">
        <h4 class="review_sub">���r���[���̏��i</h4>
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
