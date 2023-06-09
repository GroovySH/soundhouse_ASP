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
'	���r���[���e�m�F�y�[�W
'
'	���r���[�ҏW�y�[�W�̢���e���m�F����v�{�^����������J�ڂ����B
'	�ڋq�ԍ���Session("userID")�����o���A�����Ƃ��ėp����B
'
'	�N�G�������� ?item=628^NT5^^&WriteReview=Y
'	628^ ... ���[�J�[�R�[�h
'	NT5^ ... ���i�R�[�h
'	^    ... �F
'	^    ... �K�i
'	WriteReview ... ���r���[�L�q
'
'	HTTPS�łȂ��ƃG���[
'	���O�C�����Ă��Ȃ��ƃG���[
'	���O�C�����Ă���΁ASession("userID")�Ɍڋq�ԍ����Z�b�g����Ă���B
'	Session("userID")���󕶎��̎��̓G���[�@����O�C�����Ă��������B�
'	Session("userID")�Ōڋq��񂪎�o���Ȃ���΃G���[�@����O�C�����Ă��������B�
'	�G���[���b�Z�[�W���Z�b�g��Login.asp��Redirect
'
'�ύX����
'2013/05/07 GV #1507 �V�K�쐬(ReviewWrite.asp���x�[�X�ɐV�K�쐬)
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
Dim wReviewBody				'2013/05/07 GV Add ���r���[����

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
Dim oProductData		'���i���
Dim oTotalReviewData	'���r���[�i���]�j
Dim oReviewData			'���r���[�i�ʁj
Dim oCustomerData		'�ڋq���
Dim oOrderData			'�󒍏��
Dim wReviewProduct		'���r���[���̏��i
Dim wEvaluteOnpu		'���]������
Dim wBreadCrumbs		'�p�������X�g
Dim isExistReview		'���r���[�����݂��Ă���
Dim wRating				'�]��
Dim wReviewTitle		'���r���[�^�C�g��
Dim wReviewDate			'���r���[���e��
Dim wReviewBodyBr		'BR�^�O�ŉ��s�������r���[���e
Dim wReviewSelected		'�I�������]����selected������
Dim wInvalidate			'�o���f�[�V�����G���[
Dim NgFlg				'NG�t���O
Dim Mode				'�������[�h(1...save,-1...delete)
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
'---- DB�ڑ�
Call ReviewFunc_ConnectDb()

'---- ���C������
Call main()

'---- DB�ؒf
Call ReviewFunc_CloseDb()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   ' member�n�̑��̃y�[�W�����ɂȂ炤
If Err.Description <> "" Then
	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- ���O�C�����Ă��Ȃ��ꍇ�̓��O�C���y�[�W��
If wNotLogin = True Then
	Session("msg") = wMsg
	Server.Transfer "../shop/Login.asp"
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
		Call DeallocObject()

		wNotLogin = True		' ���O�C������Ă��Ȃ�
		wMsg = "���O�C�����Ă��������B"
		Exit Function
	End If

	' �ڋq���I�u�W�F�N�g���擾
	Set oCustomerData = ReviewFunc_GetCustomer()

	If oCustomerData.EOF = True Then
		'---- �I�u�W�F�N�g�̊J��
		Call DeallocObject()

		'--- Session("userID")�Ōڋq��񂪎�o���Ȃ���΃G���[�@����O�C�����Ă��������B�
		wNotLogin = True		' ���O�C������Ă��Ȃ�
		wMsg = "���O�C�����Ă��������B"
		Exit Function
	ElseIf (CLng(oCustomerData("�w����")) < 1) Then
		'---- �I�u�W�F�N�g�̊J��
		Call ReviewFunc_FreeObject()

		wMsg = "���r���[�͍w�����ꂽ�������e�ł��܂��B"
		NgFlg = True
		Exit Function
	End If

	If (IsNumeric(Mode) = false) Then
		'---- �I�u�W�F�N�g�̊J��
		Call ReviewFunc_FreeObject()

		wMsg = ""
		NgFlg = True
		Exit Function
	End If

	'---- ���i�����o��
	Set oProductData = ReviewFunc_GetProduct()

	' ���i��񂪋�̏ꍇ
	If IsObject(oProductData) = false Then
		'---- �I�u�W�F�N�g�̊J��
		Call ReviewFunc_FreeObject()

		wMsg = "���i��񂪌�����܂���ł����B"
		NgFlg = True
		Exit Function
	Else
		'---- �󒍏��̎擾
		oOrderData = ReviewFunc_GetOrder(UserID, maker_cd, product_cd)

		'�󒍏�񂪂Ȃ��ꍇ�A���r���[�𓊍e�����Ȃ�
		If (IsObject(oOrderData) = false) Then
			'---- �I�u�W�F�N�g�̊J��
			Call ReviewFunc_FreeObject()

			wMsg = "���r���[�𓊍e���邱�Ƃ��ł��܂���B"
			NgFlg = True
			Exit Function
		End If

		'---- ���i���r���[�����o��
		Set oTotalReviewData =  ReviewFunc_GetTotalReview()		' ���]
		Set oReviewData      =  ReviewFunc_GetReview(null)		' ��

		wPrefecture = oCustomerData("�ڋq�s���{��")

		'---- ���r���[�f�[�^������ꍇ
		If oReviewData.EOF = false Then
			If (CInt(Mode) <> -1) Then
				vEditMode = "�ҏW"
				isExistReview = true
			Else
				vEditMode = "�폜"
			End If

			wReviewDate = FormatDateTime(oReviewData("���e��"), 1)
			'DebugEcho("���r���[�f�[�^����")
		Else
			vEditMode = "���e"
			wReviewDate = FormatDateTime(Now, 1)
			'DebugEcho("���r���[�f�[�^�Ȃ�")
		End If
	End If

	'---- �p�������X�g
	'wBreadCrumbs = ReviewFunc_CreateBreadCrumbsHTML()
	wBreadCrumbs = "���i���r���["

	'---- ���i�摜
	wReviewProduct = ReviewFunc_CreateReviewProductPictureHTML()

	'---- ���[�U�������]���̉����A�C�R��
	wEvaluteOnpu = ReviewFunc_CreateEvaluteOnpu(wRating, false)

	'---- ���[�U�������]���̉���selected
	wReviewSelected = ReviewFunc_EvaluteSelectedArray(6, wRating)

	'---- �o���f�[�V����
	If (CInt(Mode) <> -1) Then
		wInvalidate = ReviewFunc_Validate()
	End If

	'---- �I�u�W�F�N�g�̊J��
	Call ReviewFunc_FreeObject()

End Function	' End of main()
'========================================================================
%>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>���i���r���[�b�T�E���h�n�E�X</title>
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
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  <!-- �R���e���cstart -->
  <div id="globalContents">
  <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
    <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
    <ul id="path">
      <li class="now"><%=wBreadCrumbs%></li>
    </ul>
  </div></div></div>
  <h1 class="title">���i���r���[</h1>

  <div id="review_main">
<%
	If (CInt(Mode) <> -1) Then
		'�G���[������ꍇ
		If (IsNull(wInvalidate) <> True) Then
			'---- ���̓t�H�[���̌Ăяo��
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
        <h4 class='review_sub'>���r���[���̏��i</h4>
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