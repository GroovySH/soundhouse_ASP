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
'	���i���r���[ �o�^
'
'�X�V����
'2007/04/05 ���O�C���������Ă���ΒN�ł��o�^�ł���悤�ɕύX
'2007/04/25 �n���h���l�[�����ڋq�֓o�^
'2007/04/27 ���[�����M���[�J�[�R�[�h�����[�J�[���ɕύX
'2007/08/23 �o�׃��[������̃����N�Ń��O�C�����Ă��Ȃ��Ƃ��́AOrderNo����UserID���擾����悤�ɕύX
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/20 an SELECT����LAC�N�G���[�Ă�K�p�A���[�����M�����ʊ֐����p�ɕύX
'2012/07/30 if-web ���j���[�A�����C�A�E�g����
'2013/05/07 GV #1507 ���r���[�ĕҏW�@�\
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
Dim Mode				'�������[�h(1...save,-1...delete)
Dim ReviewID			'���r���[ID
Dim oProductData		'���i���
Dim oTotalReviewData	'���r���[�i���]�j
Dim oReviewData			'���r���[�i�ʁj
Dim oCustomerData		'�ڋq���
Dim oOrderData			'�󒍏��
Dim wReviewDate			'���r���[���t
Dim urlEncItem			'URL�G���R�[�h����item
Dim backUrl				'�߂��URL
'2013/05/07 GV #1507 add end
'========================================================================

Response.buffer = true
%>
<!--#include file="ReviewFunc.inc"-->
<%

'---- UserID ���o��
'userID = Session("userID")
UserID = Session("userID")

'---- �Ăяo��������̃f�[�^���o��
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

' ���i�Ɋւ���N�G��������
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
'---- DB�ڑ�
Call ReviewFunc_ConnectDb()

Call main()

'2013/05/07 GV #1507 add start
'---- DB�ؒf
Call ReviewFunc_CloseDb()
'2013/05/07 GV #1507 add end

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
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
	End If

'2013/05/07 GV #1507 modified start
'if UserID = "" AND OrderNo <> "" then
	'---- UserID���o��
'	w_sql = ""
'	w_sql = w_sql & "SELECT �ڋq�ԍ�"
'	w_sql = w_sql & "  FROM Web�� WITH (NOLOCK)"  '2012/01/20 an mod
'	w_sql = w_sql & " WHERE �󒍔ԍ� = " & OrderNo
'
'	Set RS = Server.CreateObject("ADODB.Recordset")
'	RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic
'
'	if RS.EOF = false then
'		UserID = RS("�ڋq�ԍ�")
'	else
'		w_msg = "<font color='#ff0000'>���r���[�͓o�^�ł��܂���</font>"
'		exit function
'	end if
'
'	RS.Close
'end if
'2013/05/07 GV #1507 modified end

'2013/05/07 GV #1507 add start
	'---- ���i�����o��
	Set oProductData = ReviewFunc_GetProduct()

	' ���i��񂪋�̏ꍇ
'	If IsObject(oProductData) = false Then
	If (oProductData.EOF = true) Then
		'---- �I�u�W�F�N�g�̊J��
		Call ReviewFunc_FreeObject()

		wMsg = "���i��񂪌�����܂���ł����B"
		NgFlg = True
		Exit Function
	Else
		'---- ���i���r���[�����o��
		Set oTotalReviewData =  ReviewFunc_GetTotalReview()		' ���]
		Set oReviewData      =  ReviewFunc_GetReview(NULL)		' ��

		'---- ���r���[�f�[�^������ꍇ
		If oReviewData.EOF = false Then
			wReviewDate = oReviewData("���e��")
		Else
			wReviewDate = Now
		End If
	End If
'2013/05/07 GV #1507 add end

'2013/05/07 GV #1507 modified start
'���i���̎擾�Ɠ����Ƀ��[�J�[�����擾���Ă���̂ŁA�ȉ��̏������R�����g�A�E�g
'wMakerProduct = Split(Item, "^")
'
'---- ���[�J�[�����o��
'w_sql = ""
'w_sql = w_sql & "SELECT ���[�J�[��"
'w_sql = w_sql & "  FROM ���[�J�[ WITH (NOLOCK)"  '2012/01/20 an mod
'w_sql = w_sql & " WHERE ���[�J�[�R�[�h = '" & wMakerProduct(0) & "'"
'
'Set RS = Server.CreateObject("ADODB.Recordset")
'RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic
'
'if RS.EOF = false then
'	wMakerName = RS("���[�J�[��")
'else
'	wMakerName = wMakerProduct(0)
'end if
'RS.Close
'2013/05/07 GV #1507 modified end

'2013/05/07 GV #1507 modified start
'---- ���i���r���[���ʓo�^
w_sql = ""
w_sql = w_sql & "SELECT *"
w_sql = w_sql & "  FROM ���i���r���["
'w_sql = w_sql & " WHERE ���[�J�[�R�[�h = '" & wMakerProduct(0) & "'"
'w_sql = w_sql & "   AND ���i�R�[�h = '" & wMakerProduct(1) & "'"
'w_sql = w_sql & "   AND �ڋq�ԍ� = " & UserID
w_sql = w_sql & " WHERE ���[�J�[�R�[�h = '" & oProductData("���[�J�[�R�[�h") & "'"
w_sql = w_sql & "   AND ���i�R�[�h = '" & oProductData("���i�R�[�h") & "'"
w_sql = w_sql & "   AND �ڋq�ԍ� = " & UserID
'
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic
'
'if RS.EOF = false then
'	w_msg = "<font color='#ff0000'>���r���[��1��̂ݓ��e�ł��܂���@���ɂ��q�l����̃��r���[�͓��e����Ă��܂��̂ōē��e�ł��܂���</font>"
'	exit function
'end if


'w_msg = "<b>���i���r���[�o�^���肪�Ƃ��������܂����B</b>"

'---- insert ���i���r���[
'2013/05/07 GV #1507 modified start
'RS.AddNew
'RS("���[�J�[�R�[�h") = wMakerProduct(0)
'RS("���i�R�[�h") = wMakerProduct(1)
'RS("���e��") = now()
'RS("�ڋq�ԍ�") = UserID
'RS("�]��") = Rating
'RS("�^�C�g��") = Title
'RS("���O") = HandleName
'RS("���r���[���e") = Review
'RS("�Q�l��") = 0
'RS("�s�Q�l��") = 0

'RS.Update
'RS.close

	'DB�ɕۑ�����Ă��郌�r���[�f�[�^�����݂��A�폜���[�h�̏ꍇ�́A�폜
	If ((oReviewData.EOF = False) And (CInt(Mode) = -1)) Then
		w_msg = "<p>���i���r���[���폜���܂����B</p>"
		RS.Delete
		RS.close
	Else
		'�V�K�o�^�̏ꍇ
		If oReviewData.EOF = True Then
			RS.AddNew
		End If

		w_msg = "<p>���i���r���[��o�^���܂����B</p>"

		RS("���[�J�[�R�[�h") = oProductData("���[�J�[�R�[�h")
		RS("���i�R�[�h")     = oProductData("���i�R�[�h")
		RS("���e��")         = wReviewDate
		RS("�ڋq�ԍ�")       = UserID
		RS("�]��")           = Rating
		RS("�^�C�g��")       = Title
		RS("���O")           = HandleName
		RS("���r���[���e")   = Review
		RS("�Q�l��")         = 0
		RS("�s�Q�l��")       = 0

		RS.Update
		RS.close

		'---- �n���h���l�[���o�^
		w_sql = ""
		w_sql = w_sql & "SELECT �n���h���l�[��"
		w_sql = w_sql & "     , �ŏI�X�V��"
		w_sql = w_sql & "  FROM Web�ڋq"
		'w_sql = w_sql & " WHERE �ڋq�ԍ� = " & UserID
		w_sql = w_sql & " WHERE �ڋq�ԍ� = " & UserID
		w_sql = w_sql & "   AND (�n���h���l�[�� = '' OR �n���h���l�[�� IS NULL)"

		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.Open w_sql, Connection, adOpenStatic, adLockOptimistic

		if RS.EOF = false then
			RS("�n���h���l�[��") = HandleName
			RS("�ŏI�X�V��") = Now()
			RS.update
			RS.close
		end if

'---- �o�^���[�����M
call sendMail()
	End If	'2013/05/07 GV #1507 modified end

	backUrl = g_HTTP & "shop/ProductDetail.asp?Item=" & item	'2013/05/07 GV #1507 add

End function

'========================================================================
'
'	Function	���[�����M
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
call getCntlMst("����","���M��Email","���i���r���[", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)

'2013/05/07 GV #1507 add start
vSuffix  = ""

'���r���[�����ɂ���ꍇ�A(�ҏW)������
If (oReviewData.EOF = false) Then
	vSuffix = vSuffix & "(�ҏW)"

	'�V���b�v�R�����g���L��ꍇ
	If (oReviewData("�V���b�v�R�����g��") <> "") Then
		vSuffix = vSuffix & "�V���b�v�R�����g����"
	End If
End If

'v_body = "���i���r���[" & vbNewLine & vbNewLine
v_body = "���i���r���[" & vSuffix & vbNewLine & vbNewLine

'---- ���i���r���[���ʓo�^
vSql = ""
vSql = vSql & "SELECT ID"
vSql = vSql & "  FROM ���i���r���["
vSql = vSql & " WHERE ���[�J�[�R�[�h = '" & oProductData("���[�J�[�R�[�h") & "'"
vSql = vSql & "   AND ���i�R�[�h = '" & oProductData("���i�R�[�h") & "'"
vSql = vSql & "   AND �ڋq�ԍ� = " & UserID

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSql, Connection, adOpenStatic, adLockOptimistic

If vRS.EOF = false Then
ReviewID = vRS("ID")
End If

vRS.close
'2013/05/07 GV #1507 add end

'2013/05/07 GV #1507 mod start
v_body = v_body & "ID�@�@�@�@�@�@�F" & ReviewID & vbNewLine
'2013/05/07 GV #1507 mod end

v_body = v_body & "���e���@�@�@�@�F" & now() & vbNewLine
v_body = v_body & "�ڋq�ԍ��@�@�@�F" & UserID & vbNewLine & vbNewLine

'2013/05/07 GV #1507 mod start
'v_body = v_body & "���[�J�[���@�@�F" & wMakerName & vbNewLine
'v_body = v_body & "���i�R�[�h�@�@�F" & wMakerProduct(1) & vbNewLine & vbNewLine
v_body = v_body & "���[�J�[���@�@�F" & oProductData("���[�J�[��") & vbNewLine
v_body = v_body & "���i�R�[�h�@�@�F" & oProductData("���i�R�[�h") & vbNewLine & vbNewLine
'2013/05/07 GV #1507 mod end

v_body = v_body & "�]���@�@�@�@�@�F" & Rating & vbNewLine
v_body = v_body & "�^�C�g���@�@�@�F" & Title & vbNewLine
v_body = v_body & "�n���h���l�[���F" & HandleName & vbNewLine
v_body = v_body & "���r���[���e�@�F" & vbNewLine & Review & vbNewLine

'2013/05/ GV #1507 mod start
'Call fSendEmail(wItemChar2, wItemChar1, "���i���r���[", v_body, "")    '2012/01/20 an add
Call fSendEmail(wItemChar2, wItemChar1, "���i���r���[" & vSuffix, v_body, "")
'2013/05/ GV #1507 mod end

'Set OBJ_NewMail = Server.CreateObject("CDO.Message") '2012/01/20 an del s
'
'OBJ_NewMail.from = wItemChar2
'OBJ_NewMail.to = wItemChar1
'
'OBJ_NewMail.subject = "���i���r���["
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
<title>���i���r���[�o�^���肪�Ƃ��������܂����b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
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
      <li class="now">���i���r���[</li>
    </ul>
  </div></div></div>

  <h1 class="title">���i���r���[</h1>
  <%=w_msg%>
  <p class="btnBox"><a href="<%= backUrl %>" class="opover">���i�y�[�W�֖߂�</a></p>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>