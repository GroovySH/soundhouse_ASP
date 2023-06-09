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
'	�w�������ꗗ�y�[�W
'
'	������j���[�̢���w�����𣂨��сA�y�[�W�w�肳�ꂽ�ꍇ�́A��������Ăяo�����B
'	�Y���ڋq�̍w���ꗗ��\������B
'	�ڋq�ԍ���Session("userID")�����o���A�����Ƃ��ėp����B
'
'	HTTPS�łȂ��ƃG���[
'	���O�C�����Ă��Ȃ��ƃG���[
'	���O�C�����Ă���΁ASession("userID")�Ɍڋq�ԍ����Z�b�g����Ă���B
'	Session("userID")���󕶎��̎��̓G���[�@����O�C�����Ă��������B�
'	Session("userID")�Ōڋq��񂪎�o���Ȃ���΃G���[�@����O�C�����Ă��������B�
'	�G���[���b�Z�[�W���Z�b�g��Login.asp��Redirect
'
'	�E�Y���ڋq�̎󒍏�����������
'	�E�󒍏���EmaxDB���g�p����B(WebDB�ł͂Ȃ��B)
'	�E�w�������̏ꍇ1�y�[�W�֕\�����錏���́A20���i�v���O�������Ŏw��j
'	�E���ϒ��Əo�׏�������1SQL�ŁA�w���������SQL�ō쐬����
'	�E�e�G���A���̕\�����͌��ϓ��~��
'
'�ύX����
'2011/12/22 GV #1149 �V�K�쐬
'2012/08/11 if-web ���j���[�A�����C�A�E�g����
'2012/11/24 ok �󒍌`�ԂɃX�}�[�g�t�H����ǉ�
'
'========================================================================
'On Error Resume Next

Const THIS_PAGE_NAME = "OrderHistory.asp"
Const PAGE_SIZE = 20						' �w����������1�y�[�W������̕\���s��

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wUserID

Dim wNotLogin					' ���O�C�����Ă��Ȃ�

Dim wOrderHistryListHTML

Dim wIPage						' �\������y�[�W�ʒu (�p�����[�^)

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
'---- Session�ϐ�
wDispMsg = Session("DispMsg")
Session("DispMsg") = ""
wErrMsg = Session("ErrMsg")
Session("ErrMsg") = ""

wUserID = Session("userID")

' Get�p�����[�^
wIPage = ReplaceInput(Trim(Request("IPage")))	' �y�[�W�ʒu

If wIPage = "" Or IsNumeric(wIPage) = False Then
	wIPage = 1
Else
	wIPage = CLng(wIPage)
End If

wNotLogin = False				' ������Ԃ̓��O�C�����Ă��鎖��O��Ƃ���

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   ' member�n�̑��̃y�[�W�����ɂȂ炤
If Err.Description <> "" Then
	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

If wNotLogin = True Then
	'---- ���O�C�����Ă��Ȃ��ꍇ�̓��O�C���y�[�W��
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
	'--- ���O�C�����Ă��Ȃ���΃G���[�@����O�C�����Ă��������B�
	wNotLogin = True		' ���O�C������Ă��Ȃ�
	wMsg = "���O�C�����Ă��������B"
	Exit Function
End If

' �ڋq���擾
Set vRS_Cust = get_customer()

If vRS_Cust.EOF = True Then
	'--- Session("userID")�Ōڋq��񂪎�o���Ȃ���΃G���[�@����O�C�����Ă��������B�
	wNotLogin = True		' ���O�C������Ă��Ȃ�
	wMsg = "���O�C�����Ă��������B"
	Exit Function
End If

vRS_Cust.Close

Set vRS_Cust = Nothing

' �������������̏�����
vHistoryCount = 0

'--- �Y���ڋq�̎󒍈ꗗ���o��1 (���ϒ��E�o�׏�����)
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.�󒍔ԍ� "
vSQL = vSQL & "    , a.�󒍓� "
vSQL = vSQL & "    , a.���ϓ� "
vSQL = vSQL & "    , a.�o�׊����� "
vSQL = vSQL & "    , a.�󒍌`�� "
vSQL = vSQL & "    , a.�x�����@ "
vSQL = vSQL & "    , a.Web�󒍕ύX�J�n�� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "�� a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.�폜��     IS NULL "
vSQL = vSQL & "    AND a.�o�׊����� IS NULL "
'vSQL = vSQL & "    AND a.�󒍌`�� in ('E-mail','FAX','�C���^�[�l�b�g','�g��','�d�b','�X��','���X')"	'2012/11/24 ok Del
vSQL = vSQL & "    AND a.�󒍌`�� in ('E-mail','FAX','�C���^�[�l�b�g','�g��','�d�b','�X��','���X','�X�}�[�g�t�H��')"	'2012/11/24 ok Add
vSQL = vSQL & "    AND a.�ڋq�ԍ�   = " & wUserID & " "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "      CASE WHEN a.�󒍓� IS NULL "
vSQL = vSQL & "          THEN 1 "
vSQL = vSQL & "          ELSE 2 "
vSQL = vSQL & "      END "
vSQL = vSQL & "    , ���ϓ� DESC "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

vHTML = ""

If vRS.EOF = False Then

	vTitleWordSave = ""

	Do Until vRS.EOF = True

		'--- �o�׏�(�^�C�g��) �̔���
		vTitleWord = make_titleWord(vRS("�󒍓�"), vRS("�o�׊�����"))

		If vTitleWord <> vTitleWordSave Then

			If vTitleWordSave <> "" Then
				vHTML = vHTML & "</table>" & vbNewLine
			End If

			' ���ݏ������̃^�C�g����Ҕ�
			vTitleWordSave = vTitleWord

			'--- ��������̃^�C�g�����x������
			If vTitleWord = "������" Then
				vOrderDateLabel = "�����ϓ�"
			ElseIf vTitleWord = "�o�׏�����" Then
				vOrderDateLabel = "��������"
			ElseIf vTitleWord = "���w������" Then
				vOrderDateLabel = "��������"
			Else
				vOrderDateLabel = "��������"
			End If

			'--- �^�C�g������
			vHTML = vHTML & "<p class='table_bar'>" & vTitleWord & "</p>" & vbNewLine

			vHTML = vHTML & "<table class='order_history_list'>" & vbNewLine
			vHTML = vHTML & "  <tr>" & vbNewLine
			vHTML = vHTML & "    <th>" & vOrderDateLabel & "</th>" & vbNewLine
			vHTML = vHTML & "    <th>�������ԍ�</th>" & vbNewLine
			vHTML = vHTML & "    <th>���������@</th>" & vbNewLine
			vHTML = vHTML & "    <th>���x�����@</th>" & vbNewLine
			vHTML = vHTML & "  </tr>" & vbNewLine

		End If

		'--- ���׍s����
		vHTML = vHTML & make_orderHistoryHTML(vRS)

		vRS.MoveNext

	Loop

	vHTML = vHTML & "</table>" & vbNewLine

	' ���������������m�F�p�ɑҔ�
	vHistoryCount = vRS.RecordCount

End If

vRS.Close

'--- �Y���ڋq�̎󒍈ꗗ���o��2 (���w������)
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.�󒍔ԍ� "
vSQL = vSQL & "    , a.�󒍓� "
vSQL = vSQL & "    , a.���ϓ� "
vSQL = vSQL & "    , a.�o�׊����� "
vSQL = vSQL & "    , a.�󒍌`�� "
vSQL = vSQL & "    , a.�x�����@ "
vSQL = vSQL & "    , a.Web�󒍕ύX�J�n�� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "�� a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.�폜��     IS NULL "
vSQL = vSQL & "    AND a.�o�׊����� IS NOT NULL "
'vSQL = vSQL & "    AND a.�󒍌`�� in ('E-mail','FAX','�C���^�[�l�b�g','�g��','�d�b','�X��','���X')"	'2012/11/24 ok Del
vSQL = vSQL & "    AND a.�󒍌`�� in ('E-mail','FAX','�C���^�[�l�b�g','�g��','�d�b','�X��','���X','�X�}�[�g�t�H��')"	'2012/11/24 ok Add
vSQL = vSQL & "    AND a.�ڋq�ԍ�   = " & wUserID & " "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "    ���ϓ� DESC "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	'--- �o�׏�(�^�C�g��) ��������
	vTitleWord = "���w������"
	vOrderDateLabel = "��������"

	'--- �^�C�g������
	vHTML = vHTML & "<p class='table_bar'>" & vTitleWord & "</p>" & vbNewLine

	vHTML = vHTML & "<table class='order_history_list'>" & vbNewLine
	vHTML = vHTML & "  <tr>" & vbNewLine
	vHTML = vHTML & "    <th>" & vOrderDateLabel & "</th>" & vbNewLine
	vHTML = vHTML & "    <th>�������ԍ�</th>" & vbNewLine
	vHTML = vHTML & "    <th>���������@</th>" & vbNewLine
	vHTML = vHTML & "    <th>���x�����@</th>" & vbNewLine
	vHTML = vHTML & "  </tr>" & vbNewLine

	'--- �w��y�[�W��\������ׂ̃��R�[�h�ʒu�t��(SearchList�̏����ɕ키)
	vRS.PageSize = PAGE_SIZE
	If wIPage > ((vRS.RecordCount + (PAGE_SIZE - 1)) / PAGE_SIZE) Then		'MAX�y�[�W�𒴂���ꍇ�͍ŏI�y�[�W��
		wIPage = Fix(vRS.RecordCount / PAGE_SIZE)
	End If

	' ���R�[�h�ʒu�̈ʒu�t��
	vRS.AbsolutePage = wIPage

	For i = 0 To (vRS.PageSize - 1)

		'--- ���׍s����
		vHTML = vHTML & make_orderHistoryHTML(vRS)

		vRS.MoveNext

		If vRS.EOF Then
			Exit For
		End If

	Next

	vHTML = vHTML & "</table>" & vbNewLine

	' ���������������m�F�p�ɑҔ�
	vHistoryCount = vHistoryCount + vRS.RecordCount

	'--- �y�[�W�J�ڕ�HTML����
	vHTML = vHTML & make_pageNaviHTML(vRS, wIPage)

End If

vRS.Close

Set vRS = Nothing

'--- �w�������̑��݊m�F
If vHistoryCount <= 0 Then

	wMsg = "�w������������܂���B"
	Exit Function

End If

wOrderHistryListHTML = vHTML

End function

'========================================================================
'
'	Function	�ڋq���̎��o��
'
'========================================================================
Function get_customer()

Dim vRS
Dim vSQL

'---- �ڋq�����o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.�ڋq�ԍ� "
vSQL = vSQL & "    , a.���[�U�[ID "
vSQL = vSQL & "    , a.�ڋq�� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web�ڋq a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.�ڋq�ԍ� = " & wUserID
vSQL = vSQL & "    AND a.Web�s�f�ڃt���O <> 'Y'"

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

Set get_customer = vRS

End Function

'========================================================================
'
'	Function	��������pHTML���� (�f�[�^��1�s��)
'
'========================================================================
Function make_orderHistoryHTML(pobjRS)

Dim vHTML

If pobjRS.EOF = True Then
    Exit Function
End If

vHTML = ""

vHTML = vHTML & "  <tr>" & vbNewLine

vHTML = vHTML & "    <td>" & formatDateYYYYMMDD(pobjRS("���ϓ�")) & "</td>" & vbNewLine
vHTML = vHTML & "    <td><a href='OrderHistoryDetail.asp?OrderNo=" & pobjRS("�󒍔ԍ�") & "'>" & pobjRS("�󒍔ԍ�") & "</a></td>" & vbNewLine
vHTML = vHTML & "    <td>" & pobjRS("�󒍌`��") & "</td>" & vbNewLine
vHTML = vHTML & "    <td>" & get_paymetMethodWord(pobjRS("�x�����@")) & "</td>" & vbNewLine

vHTML = vHTML & "  </tr>" & vbNewLine

make_orderHistoryHTML = vHTML

End Function

'========================================================================
'
'	Function	�y�[�W�J�ڕ�HTML����
'
'========================================================================
Function make_pageNaviHTML(pobjRS, plngPage)

Dim vHTML
Dim i

vHTML = ""
vHTML = vHTML & "  <ol id='pagenavi'>" & vbNewLine

If plngPage <> 1 Then
	' �O�̃y�[�W
	vHTML = vHTML & "    <li id='go'><a href='JavaScript:page_onClick(" & plngPage - 1 & ");' title='�O�̃y�[�W�ɖ߂�'><span>&laquo;</span></a></li>" & vbNewline
End If

For i = 1 To pobjRS.PageCount
	If i = plngPage Then
		' ���݂̃y�[�W
		vHTML = vHTML & "    <li id='now'><a href='JavaScript:void(0);'>" & i & "</a></li>" & vbNewLine
	Else
		' �y�[�W�ԍ��w��
		vHTML = vHTML & "    <li><a href='JavaScript:page_onClick(" & i & ");'>" & i & "</a></li>" & vbNewLine
	End If
next

If plngPage <> pobjRS.PageCount Then
	' ���̃y�[�W
	vHTML = vHTML & "    <li id='go'><a href='JavaScript:page_onClick(" & plngPage + 1 & ");' title='���̃y�[�W�ɐi��'><span>&raquo;</span></a></li>" & vbNewline
End If

vHTML = vHTML & "  </ol>" & vbNewLine

make_pageNaviHTML = vHTML

End Function

'========================================================================
'
'	Function	���t���̃t�H�[�}�b�g (YYYY�NMM��DD��)
'
'========================================================================
Function formatDateYYYYMMDD(pdatDate)

Dim vDate

If IsNull(pdatDate) = True Then
	' Null �͌v�Z�s�\
	Exit Function
End If

If IsDate(pdatDate) = False Then
	' ���t���łȂ���Όv�Z�s�\
	Exit Function
End If

vDate = DatePart("yyyy", pdatDate) & "�N"

If DatePart("m", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("m", pdatDate)
Else
	vDate = vDate & DatePart("m", pdatDate)
End If

vDate = vDate & "��"

If DatePart("d", pdatDate) <= 9 Then
	vDate = vDate & "0" & DatePart("d", pdatDate)
Else
	vDate = vDate & DatePart("d", pdatDate)
End If

vDate = vDate & "��"

formatDateYYYYMMDD = vDate

End Function

'========================================================================
'
'	Function	�\���p�x�������@�����̐���
'
'	Note
'	  �x�����@              �\������
'	��������������������������������������������
'	  �R���r�j�x��       �� "�R���r�j����"
'	  �l�b�g�o���L���O   �� "�R���r�j����"
'	  �䂤����           �� "�R���r�j����"
'	  ���[��(��������)   �� "���[��"
'	  ���[��(�����Ȃ�)   �� "���[��"
'	  ���[��(��������)   �� "���[��"
'	  ��s�U��           �� "��s�U��"
'	  �����             �� "�������"
'	  ����               �� (�x�����@���̂܂�)
'	  ���|               �� (�x�����@���̂܂�)
'	  �A�}�]��           �� (�x�����@���̂܂�)
'	  �N���W�b�g�J�[�h   �� (�x�����@���̂܂�)
'
'========================================================================
Function get_paymetMethodWord(pstrPaymetMethod)

Dim vDisplayWord

If IsNull(pstrPaymetMethod) = True Then
	' Null �͔���s�\
	Exit Function
End If

If pstrPaymetMethod = "�����" Then
	vDisplayWord = "�������"
ElseIf pstrPaymetMethod = "�R���r�j�x��" Then
	vDisplayWord = "�R���r�j����"
ElseIf pstrPaymetMethod = "�l�b�g�o���L���O" Then
	vDisplayWord = "�R���r�j����"
ElseIf pstrPaymetMethod = "�䂤����" Then
	vDisplayWord = "�R���r�j����"
ElseIf pstrPaymetMethod = "��s�U��" Then
	vDisplayWord = "��s�U��"
ElseIf InStr(pstrPaymetMethod, "���[��") > 0 Then
	vDisplayWord = "���[��"
Else
	vDisplayWord = pstrPaymetMethod
End If

get_paymetMethodWord = vDisplayWord

End Function

'========================================================================
'
'	Function	�w�������̃^�C�g����������
'
'========================================================================
Function make_titleWord(pdatOrderDate, pdatShipCompleteDate)

Dim vTitleWord

If IsNull(pdatOrderDate) Then
	'--- �󒍓���Null�̏ꍇ
	vTitleWord = "������"
ElseIf IsNull(pdatOrderDate) = False And IsNull(pdatShipCompleteDate) Then
	'--- �󒍓���Null�łȂ��A�o�׊�������Null�̏ꍇ
	vTitleWord = "�o�׏�����"
ElseIf IsNull(pdatShipCompleteDate) = False Then
	'--- �o�׊�������Null�̏ꍇ
	vTitleWord = "���w������"
Else
	vTitleWord = "���w������"
End If

make_TitleWord = vTitleWord

End Function

'========================================================================
%>
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="Shift_JIS">
<title>���w�������b�T�E���h�n�E�X</title>
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
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
      <p class="home"><a href="<%=g_HTTP%>"><img src="../images/icon_home.gif" alt="HOME"></a></p>
      <ul id="path">
        <li><a href="../member/Mypage.asp">�}�C�y�[�W</a></li>
        <li class="now">���w������</li>
      </ul>
    </div></div></div>

    <h1 class="title">���w������</h1>

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