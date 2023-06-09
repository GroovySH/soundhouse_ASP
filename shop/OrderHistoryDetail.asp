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
'	�ʒ������y�[�W
'
'	OrderHistory.asp�̎󒍔ԍ������N����Ăяo�����
'	�w�����̏ڍ׏���\������B
'	�R���g���[���}�X�^����o�����ύX�\���ԑтŁA����.�S���҃R�[�h = 'internet' �̏ꍇ�A�L�����Z���ύX���\�Ƃ���B
'
'	HTTPS�łȂ��ƃG���[
'	���O�C�����Ă��Ȃ��ƃG���[
'	���O�C�����Ă���΁ASession("userID")�Ɍڋq�ԍ����Z�b�g����Ă���B
'	Session("userID")���󕶎��̎��̓G���[�@����O�C�����Ă��������B�
'	Session("userID")�Ōڋq��񂪎�o���Ȃ���΃G���[�@����O�C�����Ă��������B�
'	�G���[���b�Z�[�W���Z�b�g��Login.asp��Redirect
'
'	�E�Y���ڋq�̎󒍏�����������B�w�b�_�����A���o�ו����A�o�׊���������ʁX�Ɏ�o���B
'	�E�w�b�_�[�����o���Ȃ��ꍇ�̓G���[��Y���̒�����񂪂���܂���B� OrderHistory.asp �� Redirect
'	�E�󒍏���EmaxDB���g�p����B(WebDB�ł͂Ȃ��B)
'
'�ύX����
'2011/12/27 GV #1149 �V�K�쐬
'2012/08/11 if-web ���j���[�A�����C�A�E�g����
'2013/04/30 if-web �����ԍ��\�����R�����g�A�E�g
'2013/07/11 GV #1507 ���r���[�ҏW�@�\
'
'========================================================================
On Error Resume Next

Const THIS_PAGE_NAME = "OrderHistoryDetail.asp"
Const UPDATEABLE_STAFF_CD = "Internet"			' �L�����Z���E�����ύX �\�� ��.�S���҃R�[�h

Const FIRST_STEP = True			' 1st step �Ώ�

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wUserID

Dim wNotLogin					' ���O�C�����Ă��Ȃ�
Dim wUpdateable					' �I�[�_�[�̕ύX���\
Dim wDeleteable					' �I�[�_�[�̃L�����Z�����\
Dim wTaxRate					' ����ŗ�
Dim wOrderUpdateStartTime		' �󒍕ύX�J�n����(�R���g���[���}�X�^)
Dim wOrderUpdateEndTime			' �󒍕ύX�I������(�R���g���[���}�X�^)

Dim wOrderDetailHTML

Dim wOrderNo					' �󒍔ԍ� (�p�����[�^)

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
wOrderNo = ReplaceInput(Trim(Request("OrderNo")))	' �󒍔ԍ�

If wOrderNo = "" Or IsNumeric(wOrderNo) = False Then
	wOrderNo = 0				' main �ŃG���[�Ƃ��Ď�舵��
Else
	wOrderNo = CLng(wOrderNo)
End If

wNotLogin = False				' ������Ԃ̓��O�C�����Ă��鎖��O��Ƃ���

wUpdateable = False				' �I�[�_�[�̕ύX�͕s��
wDeleteable = False				' �I�[�_�[�̃L�����Z���͕s��

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

If wMsg <> "" Then
	'--- �w�b�_�[�����o���Ȃ�,�󒍂�������Ȃ����̏ꍇ�̓G���[�@OrderHistory.asp��Redirect
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
Dim vCurrentTime			' ���݂̎���
Dim vStaffCd				' �S���҃R�[�h
Dim vParam
Dim vTitleWord
Dim vTitleWordSave
Dim vOrderDateLabel
Dim vTrackingNumber			' �����ԍ�
Dim vTransporterCd			' �^����ЃR�[�h
Dim vTransporterName		' �^����Ж�
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

' �p�����[�^�̃`�F�b�N (�󒍔ԍ�)
If wOrderNo <= 0 Then
	'--- �s���Ȏ󒍔ԍ��̏ꍇ�@��Y���̒�����񂪂���܂���B��@OrderHistory.asp��Redirect
	wMsg = "�Y���̒�����񂪂���܂���B"
	Exit Function
End If

'--- �R���g���[���}�X�^���u�󒍕ύX�J�n���ԁv�u�󒍕ύX�I�����ԁv��o��
If get_updateTimeSlot(wOrderUpdateStartTime, wOrderUpdateEndTime) = False Then
	'--- �R���g���[���}�X�^�ɒ�`����
	wMsg = "�G���[���������܂����B"
	Exit Function
End If

'--- �w�b�_�����̏���o��
vSQL = ""
vSQL = vSQL & "SELECT TOP 1 "
vSQL = vSQL & "      a.�󒍔ԍ� "
vSQL = vSQL & "    , a.���ϓ� "
vSQL = vSQL & "    , a.�󒍓� "
vSQL = vSQL & "    , a.�o�׊����� "
vSQL = vSQL & "    , a.�󒍌`�� "
vSQL = vSQL & "    , a.�x�����@ "
vSQL = vSQL & "    , a.���i���v���z "
vSQL = vSQL & "    , a.���� "
vSQL = vSQL & "    , a.����萔�� "
vSQL = vSQL & "    , a.�󒍍��v���z "
vSQL = vSQL & "    , a.�ꊇ�o�׃t���O "
vSQL = vSQL & "    , a.�̎������� "
vSQL = vSQL & "    , a.�̎����A������ "
vSQL = vSQL & "    , a.Web�󒍕ύX�J�n�� "
vSQL = vSQL & "    , a.����ŗ� "
vSQL = vSQL & "    , a.�^����ЃR�[�h "
vSQL = vSQL & "    , a.�S���҃R�[�h "
vSQL = vSQL & "    , b.�������͐�X�֔ԍ� "
vSQL = vSQL & "    , b.�������͐�s���{�� "
vSQL = vSQL & "    , b.�������͐�Z�� "
vSQL = vSQL & "    , b.�������͐於�O "
vSQL = vSQL & "    , b.�ŏI�w��[�� "
vSQL = vSQL & "    , b.�ŏI���Ԏw�� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "��     a WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "�󒍖��� b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        b.�󒍔ԍ� = a.�󒍔ԍ� "
vSQL = vSQL & "    AND a.�󒍔ԍ� = " & wOrderNo & " "
vSQL = vSQL & "    AND a.�ڋq�ԍ� = " & wUserID & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	'--- �w�b�_�[�����o���Ȃ��ꍇ�̓G���[��Y���̒�����񂪂���܂���B��@OrderHistory.asp��Redirect
	vRS.Close
	Set vRS = Nothing
	wMsg = "�Y���̒�����񂪂���܂���B"
	Exit Function
End If

'--- ����ŗ���o��
wTaxRate = CLng(vRS("����ŗ�"))

'--- �^����ЃR�[�h��o��
vTransporterCd = vRS("�^����ЃR�[�h")

vHTML = ""

'--- �o�׏�(�^�C�g��) �̔���
vTitleWord = make_titleWord(vRS("�󒍓�"), vRS("�o�׊�����"), vRS("Web�󒍕ύX�J�n��"))

vHTML = vHTML & "<p class='table_bar'>" & vTitleWord & "</p>" & vbNewLine

'--- �������� �` ���x�����@
vHTML = vHTML & "<table class='order_history_list'>" & vbNewLine
vHTML = vHTML & "  <tr>" & vbNewLine
vHTML = vHTML & "    <th>��������</th>" & vbNewLine
vHTML = vHTML & "    <th>�������ԍ�</th>" & vbNewLine
vHTML = vHTML & "    <th>���������@</th>" & vbNewLine
vHTML = vHTML & "    <th>���x�����@</th>" & vbNewLine
vHTML = vHTML & "  </tr>    " & vbNewLine
vHTML = vHTML & "  <tr>" & vbNewLine
vHTML = vHTML & "    <td>" & formatDateYYYYMMDD_J(vRS("���ϓ�")) & "</td>" & vbNewLine
vHTML = vHTML & "    <td class='number'>" & vRS("�󒍔ԍ�") & "</td>" & vbNewLine
vHTML = vHTML & "    <td>" & vRS("�󒍌`��") & "</td>" & vbNewLine
vHTML = vHTML & "    <td>" & get_paymetMethodWord(vRS("�x�����@")) & "</td>" & vbNewLine
vHTML = vHTML & "  </tr>" & vbNewLine
vHTML = vHTML & "</table>" & vbNewLine

'--- ���͂���E�z�����@�E�����w��E�̎发�E���v���z(���i���v,����,����萔��,���w�����v���z)
vHTML = vHTML & "<dl class='modify_list'>" & vbNewLine
vHTML = vHTML & "  <dt class='address'>���͂���</dt>" & vbNewLine
vHTML = vHTML & "  <dd class='address'>" & vbNewLine
vHTML = vHTML & "��" & vRS("�������͐�X�֔ԍ�") & "<br>" & vbNewLine
vHTML = vHTML & "" & vRS("�������͐�s���{��") & vRS("�������͐�Z��") & "<br>" & vbNewLine
vHTML = vHTML & "" & vRS("�������͐於�O") & "&nbsp;�l</dd>" & vbNewLine
vHTML = vHTML & "  <dt>�z�����@</dt>" & vbNewLine
vHTML = vHTML & "  <dd>" & get_shipTypeWord(vRS("�ꊇ�o�׃t���O")) & "</dd>" & vbNewLine
vHTML = vHTML & "  <dt>�����w��</dt>" & vbNewLine
If IsDate(vRS("�ŏI�w��[��")) Then
	vHTML = vHTML & "  <dd>" & formatDateYYYYMMDD_J(vRS("�ŏI�w��[��")) & "�@" & vRS("�ŏI���Ԏw��") & "</dd>" & vbNewLine
Else
	vHTML = vHTML & "  <dd>&nbsp;</dd>" & vbNewLine
End If
vHTML = vHTML & "  <dt>�̎���</dt>" & vbNewLine
If IsNull(vRS("�̎�������")) = False And vRS("�̎�������") <> "" Then
	vHTML = vHTML & "  <dd>�̎�������F" & vRS("�̎�������") & " �l / �̎����A�������F" & vRS("�̎����A������") & "</dd>" & vbNewLine
Else
	vHTML = vHTML & "  <dd>&nbsp;</dd>" & vbNewLine
End If
vHTML = vHTML & "  <dt class='total_accounts'>���v���z</dt>" & vbNewLine
vHTML = vHTML & "  <dd class='total_accounts'>" & vbNewLine
vHTML = vHTML & "    <ul>" & vbNewLine
vHTML = vHTML & "      <li>���i���v(�ō�)�F" & FormatNumber(get_detailTotalPrice(wOrderNo, wTaxRate), 0) & "�~</li>" & vbNewLine
vHTML = vHTML & "      <li>����(�ō�)�F" & FormatNumber(calc_taxInclusivePrice(vRS("����"), wTaxRate), 0) & "�~</li>" & vbNewLine
If vRS("�x�����@") = "�����" Then
	vHTML = vHTML & "      <li>����萔��(�ō�)�F" & FormatNumber(calc_taxInclusivePrice(vRS("����萔��"), wTaxRate), 0) & "�~</li>" & vbNewLine
End If
vHTML = vHTML & "      <li>���w�����v���z(�ō�)�F" & FormatNumber(vRS("�󒍍��v���z"), 0) & "�~</li>" & vbNewLine
vHTML = vHTML & "    </ul>" & vbNewLine
vHTML = vHTML & "  </dd>" & vbNewLine
vHTML = vHTML & "</dl>" & vbNewLine


'--- ����p�̃{�^���\��
vCurrentTime = Time()
vStaffCd = LCase(vRS("�S���҃R�[�h") & "")

vHTML = vHTML & "<ul id='order_modify'>" & vbNewline

If FIRST_STEP = False Then	' 2011/12/22 1st step �Ώ�	2011/12/28 hn �L�����Z�����Ȃ�

If isUpdateableTime(vCurrentTime, wOrderUpdateStartTime, wOrderUpdateEndTime) _
And vStaffCd = LCase(UPDATEABLE_STAFF_CD) _
And IsNull(vRS("Web�󒍕ύX�J�n��")) Then

	' �L�����Z���\
	vHTML = vHTML & "  <li><a href='javascript:void(0);' title='�������e���L�����Z��' class='showLayer_ordercancel'>�������e���L�����Z��</a></li>" & vbNewline
	wDeleteable = True
Else
	vHTML = vHTML & "  <li>�������e���L�����Z��</li>" & vbNewline
	wDeleteable = False
End If

'If FIRST_STEP = False Then	' 2011/12/22 1st step �Ώ�

If isUpdateableTime(vCurrentTime, wOrderUpdateStartTime, wOrderUpdateEndTime) _
And vStaffCd = LCase(UPDATEABLE_STAFF_CD) _
And IsNull(vRS("Web�󒍕ύX�J�n��")) Then
	' �ύX�\
	vHTML = vHTML & "  <li><a href='javascript:void(0);' title='�������e��ύX' class='showLayer_ordermodify'>�������e��ύX</a></li>" & vbNewline
	wUpdateable = True
Else
	vHTML = vHTML & "  <li>�������e��ύX</li>" & vbNewline
	wUpdateable = False
End If

vHTML = vHTML & "  <li><a href='javascript:void(0);' title='���̒������Ē���' class='showLayer_reorder'>���̒������Ē���</a></li>" & vbNewline

End If	' 2011/12/22 1st step �Ώ�

vHTML = vHTML & "  <li><a href='Inquiry.asp' title='�������e�̂��⍇��'>�������e�̂��⍇��</a></li>" & vbNewline
vHTML = vHTML & "</ul>" & vbNewline

vRS.Close


'--- ���o�׃f�[�^�̏���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      b.�󒍖��הԍ� "
vSQL = vSQL & "    , b.���[�J�[�R�[�h "
vSQL = vSQL & "    , b.���i�R�[�h "
vSQL = vSQL & "    , b.�F "
vSQL = vSQL & "    , b.�K�i "
vSQL = vSQL & "    , b.�󒍒P�� "
vSQL = vSQL & "    , b.�󒍐��� "
vSQL = vSQL & "    , b.�󒍈������v���� "
vSQL = vSQL & "    , b.�o�׍��v���� "
vSQL = vSQL & "    , c.���[�J�[�� "
vSQL = vSQL & "    , d.���i�� "
vSQL = vSQL & "    , d.���i�T��Web "
vSQL = vSQL & "    , d.���i�摜�t�@�C����_�� "
vSQL = vSQL & "    , d.Web���i�t���O "
vSQL = vSQL & "    , x.�o�ח\��� "
vSQL = vSQL & "    , x.�\�[�X "
vSQL = vSQL & "    , x.�o�ח\��e�L�X�g "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "�󒍖��� b WITH (NOLOCK) "
vSQL = vSQL & "        LEFT JOIN " & gLinkServer & "�󒍖��׏o�ח\�� x WITH (NOLOCK) "
vSQL = vSQL & "          ON     x.�󒍔ԍ�     = b.�󒍔ԍ� "
vSQL = vSQL & "             AND x.�󒍖��הԍ� = b.�󒍖��הԍ� "
vSQL = vSQL & "             AND x.�o�ח\��A�� = 1 "
vSQL = vSQL & "             AND x.�ύX��       = (SELECT MAX(y.�ύX��) "
vSQL = vSQL & "                                   FROM   " & gLinkServer & "�󒍖��׏o�ח\�� y WITH (NOLOCK) "
vSQL = vSQL & "                                   WHERE      y.�󒍔ԍ�     = b.�󒍔ԍ� "
vSQL = vSQL & "                                          AND y.�󒍖��הԍ� = b.�󒍖��הԍ�) "
vSQL = vSQL & "    , " & gLinkServer & "���[�J�[ c WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "���i d WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        c.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
vSQL = vSQL & "    AND d.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
vSQL = vSQL & "    AND d.���i�R�[�h = b.���i�R�[�h "
vSQL = vSQL & "    AND b.�Z�b�g�i�e���הԍ� = 0 "
vSQL = vSQL & "    AND b.�󒍔ԍ� = " & wOrderNo & " "
vSQL = vSQL & "    AND b.�󒍐��� > b.�o�׍��v���� "
vSQL = vSQL & "ORDER BY "
vSQL = vSQL & "      c.���[�J�[�� "
vSQL = vSQL & "    , d.���i�� "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	'--- �f�[�^�����݂���ꍇ�̂݁u���o�׃f�[�^�v�\��
	vHTML = vHTML & "<div class='order_history_container'>" & vbNewline
	vHTML = vHTML & "<p>���o��</p>" & vbNewline

	Do Until vRS.EOF = True

		vHTML = vHTML & make_orderDetailHTML(vRS, wTaxRate)

		vRS.MoveNext

	Loop

	vHTML = vHTML & "</div>" & vbNewline

End If

vRS.Close


'--- �o�׊����f�[�^�̏���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      b.�󒍖��הԍ� "
vSQL = vSQL & "    , b.���[�J�[�R�[�h "
vSQL = vSQL & "    , b.���i�R�[�h "
vSQL = vSQL & "    , b.�F "
vSQL = vSQL & "    , b.�K�i "
vSQL = vSQL & "    , b.�󒍒P�� "
vSQL = vSQL & "    , b.�󒍐��� "
vSQL = vSQL & "    , b.�󒍈������v���� "
vSQL = vSQL & "    , b.�o�׍��v���� "
vSQL = vSQL & "    , f.�o�א��� "
vSQL = vSQL & "    , c.���[�J�[�� "
vSQL = vSQL & "    , d.���i�� "
vSQL = vSQL & "    , d.���i�T��Web "
vSQL = vSQL & "    , d.���i�摜�t�@�C����_�� "
vSQL = vSQL & "    , d.Web���i�t���O "
vSQL = vSQL & "    , e.�����ԍ� "
vSQL = vSQL & "    , NULL AS �o�ח\��� "
vSQL = vSQL & "    , NULL AS �o�ח\��e�L�X�g "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "�󒍖���     b WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "���[�J�[     c WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "���i         d WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "�󒍑����   e WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "�o�ז���View f WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        c.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
vSQL = vSQL & "    AND d.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
vSQL = vSQL & "    AND d.���i�R�[�h = b.���i�R�[�h "
vSQL = vSQL & "    AND e.�󒍔ԍ� = b.�󒍔ԍ� "
vSQL = vSQL & "    AND f.�o�הԍ� = e.�o�הԍ� "
vSQL = vSQL & "    AND f.�󒍔ԍ� = b.�󒍔ԍ� "
vSQL = vSQL & "    AND f.�󒍖��הԍ� = b.�󒍖��הԍ� "
vSQL = vSQL & "    AND f.�Z�b�g�i�e���הԍ� = 0 "
vSQL = vSQL & "    AND b.�󒍔ԍ� = " & wOrderNo & " "
vSQL = vSQL & "ORDER BY  "
vSQL = vSQL & "      e.�����ԍ� "
vSQL = vSQL & "    , c.���[�J�[�� "
vSQL = vSQL & "    , d.���i�� "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	'--- �f�[�^�����݂���ꍇ�̂݁u�o�׊����f�[�^�v�\��
	vHTML = vHTML & "<div class='order_history_container'>" & vbNewline
	vHTML = vHTML & "<p>�o�׊���</p>" & vbNewline

	vTrackingNumber = ""

	'--- �^����Ж�
	vTransporterName = get_transporterName(vTransporterCd)

	Do Until vRS.EOF = True

'2013/04/30 if-web del s
'		If (vRS("�����ԍ�") & "") <> vTrackingNumber Then
'
'			vTrackingNumber = vRS("�����ԍ�") & ""
'
'			vHTML = vHTML & "<dl class='modify_list'>" & vbNewline
'			vHTML = vHTML & "  <dt>�����ԍ�</dt>" & vbNewline
'
'			' �����ԍ��̕\��
'			If vTransporterName = "����" Then
'				vHTML = vHTML & "  <dd><a href='http://k2k.sagawa-exp.co.jp/cgi-bin/mole.mcgi?oku01=" & vTrackingNumber & "' target='_blank'>" & vTrackingNumber & "�i" & vTransporterName & "�j</a></dd>" & vbNewline
'			ElseIf vTransporterName = "���Z" Then
'				vHTML = vHTML & "  <dd><a href='http://track.seino.co.jp/kamotsu/KamotsuPrintServlet?GNPNO1=" & vTrackingNumber & "&ACTION=DETAIL&NUMBER=1' target='_blank'>" & vTrackingNumber & "�i" & vTransporterName & "�j</a></dd>" & vbNewline
'			Else
'				vHTML = vHTML & "  <dd>" & vTrackingNumber & "�i" & vTransporterName & "�j</dd>" & vbNewline
'			End If
'
'			vHTML = vHTML & "</dl>" & vbNewline
'
'		End If
'2013/04/30 if-web del e

		' ���ו\��
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
'	Function	���i���v���z�̎擾 (�ō��݋��z)
'	Note		�󒍖���.�󒍒P�� �ɑ΂��A�ō��݋��z���v�Z��A�󒍖���.�󒍐��ʂ��|���A���̑S���ב����v
'
'========================================================================
Function get_detailTotalPrice(plngOrderNo, plngTaxRate)

Dim vRS
Dim vSQL
Dim vTotalPrice

get_detailTotalPrice = 0

'---- �󒍖��ׂ́u�󒍒P���v�Ɓu�󒍐��ʁv�����o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.�󒍒P�� "
vSQL = vSQL & "    , a.�󒍐��� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "�󒍖��� a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.�󒍔ԍ� = " & wOrderNo & " "

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

	If IsNumeric(vRS("�󒍒P��")) And IsNumeric(vRS("�󒍐���")) Then

		' �P���̐ō��݋��z�v�Z��A���ʂ��|����
		vTotalPrice = vTotalPrice + (calcPrice(vRS("�󒍒P��"), plngTaxRate) * vRS("�󒍐���"))

	End If

	vRS.MoveNext

Loop

vRS.Close
Set vRS = Nothing

get_detailTotalPrice = vTotalPrice

End Function

'========================================================================
'
'	Function	���Ŋz�̌v�Z
'
'========================================================================
Function calc_taxInclusivePrice(plngPrice, plngTaxRate)

calc_taxInclusivePrice = Fix(plngPrice * (100 + plngTaxRate) / 100)

End Function

'========================================================================
'
'	Function	�������חpHTML���� (�f�[�^��1�s��)
'
'========================================================================
Function make_orderDetailHTML(pobjRS, plngTaxRate)

Dim vHTML
Dim vItemName
Dim vWebItem					' Web���i�t���O
Dim vParam
Dim vProductDetailLink			' ProductDetail.asp �ւ̃����N URL
Dim vShippingStatus				' ���ׂ̏o�׏�
Dim vShippingComplete			' �o�׊���
Dim vPrice						' ���׋��z

If pobjRS.EOF = True Then
    Exit Function
End If

vWebItem = pobjRS("Web���i�t���O") & ""
If vWebItem = "Y" Then
	vParam = Server.URLEncode(pobjRS("���[�J�[�R�[�h") & "^" & pobjRS("���i�R�[�h") & "^" & Trim(pobjRS("�F")) & "^" & Trim(pobjRS("�K�i")))
	vProductDetailLink = g_HTTP & "shop/ProductDetail.asp?item=" & vParam
Else
	vProductDetailLink = ""
End If

If (Trim(pobjRS("�F")) & Trim(pobjRS("�K�i"))) <> "" Then
	vItemName = pobjRS("���i��") & "/" & Trim(pobjRS("�F")) & "/" & Trim(pobjRS("�K�i"))
Else
	vItemName = pobjRS("���i��") & ""
End If

' ���ׂ̏o�׏�
vShippingStatus = ""
vShippingComplete = False

If pobjRS("�󒍐���") = pobjRS("�o�׍��v����") Then

	vShippingStatus = "�o�׊���"
	vShippingComplete = True

ElseIf pobjRS("�󒍐���") = pobjRS("�󒍈������v����") Then

	vShippingStatus = "�o�׏�����"

ElseIf pobjRS("�󒍐���") > pobjRS("�󒍈������v����") Then

	If IsNull(pobjRS("�o�ח\���")) _
	And IsNull(pobjRS("�o�ח\��e�L�X�g")) Then

		vShippingStatus = "���񂹒�"

	ElseIf IsNull(pobjRS("�o�ח\���")) = False Then

		vShippingStatus = formatDateMMDD_J(pobjRS("�o�ח\���")) & "�@���ח\��"

	ElseIf IsNull(pobjRS("�o�ח\��e�L�X�g")) = False Then

		vShippingStatus = pobjRS("�o�ח\��e�L�X�g") & "�@���ח\��"

	End If

End If

' ���ׂ̋��z�v�Z
If IsNumeric(pobjRS("�󒍒P��")) And IsNumeric(pobjRS("�󒍐���")) Then
	' �󒍒P��(�ō���) * �󒍐���   (�󒍒P��(�ō���) : calcPrice(�󒍒P��, ����ŗ�))
	vPrice = calcPrice(pobjRS("�󒍒P��"), plngTaxRate) * pobjRS("�󒍐���")
Else
	vPrice = 0
End If

vHTML = ""

vHTML = vHTML & "<table class='order_history'>" & vbNewline
vHTML = vHTML & "  <tr>" & vbNewline
vHTML = vHTML & "    <td class='list_left'>" & vbNewline
If (pobjRS("���i�摜�t�@�C����_��") & "") <> "" _
And vWebItem = "Y" Then
	vHTML = vHTML & "      <a href='" & vProductDetailLink & "'><img src='prod_img/" & pobjRS("���i�摜�t�@�C����_��") & "' width='100' height='50' alt=''></a>" & vbNewline
Else
	vHTML = vHTML & "      <img src='prod_img/" & pobjRS("���i�摜�t�@�C����_��") & "' width='100' height='50' alt=''>" & vbNewline
End If
vHTML = vHTML & "    </td>" & vbNewline
vHTML = vHTML & "    <td>" & vbNewline
vHTML = vHTML & "      " & pobjRS("���[�J�[��") & "<br>" & vbNewline
If vWebItem = "Y" Then
	vHTML = vHTML & "      <a href='" & vProductDetailLink & "'>" & vItemName & "</a><br>" & vbNewline
Else
	vHTML = vHTML & "      " & vItemName & "<br>" & vbNewline
End If
vHTML = vHTML & "      " & pobjRS("���i�T��Web") & vbNewline
vHTML = vHTML & "    </td>" & vbNewline
vHTML = vHTML & "    <td class='contact'>" & vbNewline
vHTML = vHTML & "      <ul>" & vbNewline
vHTML = vHTML & "        <li>" & vShippingStatus & "</li>" & vbNewline
vHTML = vHTML & "        <li>" & pobjRS("�󒍐���") & "�_�F" & FormatNumber(vPrice, 0) & "�~�i�ō��j</li>" & vbNewline
vHTML = vHTML & "        <li><a href='Inquiry.asp?MakerNm=" & Server.URLEncode(pobjRS("���[�J�[��")) & "&ProductCd=" & Server.URLEncode(pobjRS("���i�R�[�h")) & "' class='tipBtn'>���̏��i�̂��⍇��</a></li>" & vbNewline
If vShippingComplete = True _
And vWebItem = "Y" Then
'2013/07/11 GV #1507 mod start
	If isReviewEntered(pobjRS("���[�J�[�R�[�h"), pobjRS("���i�R�[�h"), wUserID) = False Then
		' ���r���[���L���̏ꍇ�̂�
'		vHTML = vHTML & "        <li><a href='" & vProductDetailLink & "&WriteReview=Y#review' class='tipBtn'>���r���[������</a></li>" & vbNewline
		vHTML = vHTML & "        <li><a href='" & g_HTTPS & "shop/ReviewWrite.asp?item=" & vParam & "' class='tipBtn'>���r���[������</a></li>" & vbNewline
	Else
'		vHTML = vHTML & "        <li>&nbsp;</li>" & vbNewline
		vHTML = vHTML & "        <li><a href='" & g_HTTPS & "shop/ReviewWrite.asp?item=" & vParam & "' class='tipBtn'>���r���[��ҏW</a></li>" & vbNewline
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
'	Function	���t���̃t�H�[�}�b�g (YYYY�NMM��DD��)
'
'========================================================================
Function formatDateYYYYMMDD_J(pdatDate)

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

formatDateYYYYMMDD_J = vDate

End Function

'========================================================================
'
'	Function	���t���̃t�H�[�}�b�g (MM��DD��)
'
'========================================================================
Function formatDateMMDD_J(pdatDate)

Dim vDate

If IsNull(pdatDate) = True Then
	' Null �͌v�Z�s�\
	Exit Function
End If

If IsDate(pdatDate) = False Then
	' ���t���łȂ���Όv�Z�s�\
	Exit Function
End If

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

formatDateMMDD_J = vDate

End Function

'========================================================================
'
'	Function	���t���̃t�H�[�}�b�g (YYYY/MM/DD)
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
'	Function	�^����Ж��̎��o��
'
'========================================================================
Function get_transporterName(pstrTransporterCd)

Dim vRS
Dim vSQL
Dim vTransporterName

'---- �^����Ж����o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "    a.�^����З��� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "�^����� a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.�^����ЃR�[�h = " & pstrTransporterCd

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then
	vTransporterName = vRS("�^����З���") & ""
Else
	vTransporterName = ""
End If

vRS.Close
Set vRS = Nothing

get_transporterName = vTransporterName

End Function

'========================================================================
'
'	Function	���Ƀ��r���[���L���ς݂��H
'
'========================================================================
Function isReviewEntered(pstrMakerCd, pstrItemCd, plngCustNo)

Dim vRS
Dim vSQL
Dim vEntered

'---- ���i���r���[���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "    a.ID "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    ���i���r���[ a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.���[�J�[�R�[�h = '" & pstrMakerCd & "' "
vSQL = vSQL & "    AND a.���i�R�[�h = '" & escapeSingleQuote(pstrItemCd) & "' "
vSQL = vSQL & "    AND a.�ڋq�ԍ� = " & plngCustNo & " "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	' ���r���[�Ȃ� (���r���[���L��)
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
'	Function	�󒍕ύX�\���ԑт̎擾
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

'--- �R���g���[���}�X�^���o��
Call getCntlMst("Web", "��", "�󒍕ύX�\���ԑ�", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

If IsDate(vstrItemChar1) = False Then
	' �V�X�e���̐ݒ�l�s��
	Exit Function
End If

If IsDate(vstrItemChar2) = False Then
	' �V�X�e���̐ݒ�l�s��
	Exit Function
End If

'--- �J�n�I������ �ԋp
pdatStartTime = CDate(vstrItemChar1)
pdatEndTime = CDate(vstrItemChar2)

get_updateTimeSlot = True

End function

'========================================================================
'
'	Function	�󒍕ύX�\���ԑт̔���
'	Note		���ݎ������A�󒍕ύX�J�n���ԁ`�󒍕ύX�I������ �̊Ԃł��邩���肷��
'
'========================================================================
Function isUpdateableTime(pdatTargetTime, pdatStartTime, pdatEndTime)

isUpdateableTime = False

If pdatStartTime < pdatEndTime Then
	' �����ׂ��Ȃ�����
	If pdatStartTime <= pdatTargetTime And pdatTargetTime <= pdatEndTime Then
		' �͈͓�
		isUpdateableTime = True
	End If
Else
	' �����ׂ����� (�I�������̕������������̏ꍇ Start �` 23:59:59 Or 0:00 �` End)
	If pdatStartTime <= pdatTargetTime And pdatTargetTime <= CDate("23:59:59") _
	Or CDate("00:00:00") <= pdatTargetTime And pdatTargetTime <= pdatEndTime Then
		' �͈͓�
		isUpdateableTime = True
	End If
End If

End function

'========================================================================
'
'	Function	�o�׏󋵂̃^�C�g����������
'
'	Note
'		���L�̏��ԂŃ`�F�b�N���s��
'		0. �ύX��     : Web�󒍕ύX�J�n�� IS NOT NULL �̎�
'		1. �o�׊���   : �o�׊����� IS NOT NULL �̎�
'		2. �o�׏����� : �󒍓� IS NOT NULL AND �o�׊����� IS NULL �̎�
'		3. �����ς�   : �󒍓� IS NULL �̎�
'
'========================================================================
Function make_titleWord(pdatOrderDate, pdatShipCompleteDate, pdatWebOrderUpdateStartDate)

If IsNull(pdatWebOrderUpdateStartDate) = False Then
	'--- Web�󒍕ύX�J�n����Null�łȂ��ꍇ
	make_titleWord = "�ύX��"
	Exit Function
ElseIf IsNull(pdatShipCompleteDate) = False Then
	'--- �o�׊�������Null�łȂ��ꍇ
	make_titleWord = "�o�׊���"
	Exit Function
ElseIf IsNull(pdatOrderDate) = False And IsNull(pdatShipCompleteDate) Then
	'--- �󒍓���Null�łȂ��A�o�׊�������Null�̏ꍇ
	make_titleWord = "�o�׏�����"
	Exit Function
ElseIf IsNull(pdatOrderDate) Then
	'--- �󒍓���Null�̏ꍇ
	make_titleWord = "�����ς�"
	Exit Function
End If

End Function

'========================================================================
'
'	Function	�\���p�z�����@�����̐���
'
'========================================================================
Function get_shipTypeWord(pstrIkkatsuSyukkaFlg)

If IsNull(pstrIkkatsuSyukkaFlg) = True Then
	' Null �͔���s�\
	Exit Function
End If

If pstrIkkatsuSyukkaFlg = "Y" Then
	'--- �ꊇ�o�ׂ̏ꍇ
	get_shipTypeWord = "�ꊇ�o��"
	Exit Function
Else
	get_shipTypeWord = "�݌ɏ��i����o��"
	Exit Function
End If

End Function

'========================================================================
'
'	Function	SQL�T�[�o�p�A�V���O���N�I�[�e�[�V�����}�~�����̕t��
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
<title>���������b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel='stylesheet' href='../member/style/mypage.css?201309xx' type='text/css'>
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
        <li><a href="OrderHistory.asp">���w������</a></li>
        <li class="now">���������</li>
      </ul>
    </div></div></div>

    <h1 class="title">���������</h1>

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
<% ' �����L�����Z���|�b�v�A�b�v %>
<div class="overContent" id="overContent_ordercancel">
  <h2>�������e���L�����Z��</h2>
  <p>���̒������e���L�����Z�����܂����H��x�L�����Z�����ꂽ������߂����͂ł��܂���B</p>

  <form name="f_cancel" method="post" action="OrderHistoryDelete.asp">
    <input type="submit" value="�������e���L�����Z������" class="strong_btn">
    <input type='hidden' name='OrderNo' value='<% = wOrderNo %>'>
  </form>

  <ul class="back">
    <li><a href="javascript:void(0);" onClick="backclose();">��&nbsp;�߂�</a></li>
  </ul>
</div>
<% End If %>

<% If wUpdateable Then %>
<% ' �����ύX�|�b�v�A�b�v %>
<div class="overContent" id="overContent_ordermodify">
  <h2>�������e��ύX</h2>
  <p>�V���b�s���O�J�[�g�y�[�W�ɖ߂��Ă������葱������蒼�����Ƃ��ł��܂��B<br>
�����݃J�[�g�ɓ����Ă��鏤�i�͏㏑������܂��B<br>
���������e�̕ύX��r���ł�߂��ꍇ�ɂ́A���̂��������e�̓L�����Z������܂���B</p>

  <form name="f_change" method="post" action="OrderHistoryChange.asp">
    <input type="submit" value="�������e��ύX����" class="strong_btn">
    <input type='hidden' name='OrderNo' value='<% = wOrderNo %>'>
  </form>

  <ul class="back">
    <li><a href="javascript:void(0);" onClick="backclose();">��&nbsp;�߂�</a></li>
  </ul>

  <div id="ordermodify_flow">
    <p>���������e�̕ύX�̗���</p>
    <dl>
      <dt><img src="images/shopping_step1_off.gif" alt="�V���b�s���O�J�[�g"></dt>
      <dd>�V���b�s���O�J�[�g�y�[�W�ŁA���i�̒ǉ��E�폜���s���܂��B</dd>
      <dt><img src="images/shopping_step2_off.gif" alt="���͂���A���x�����@�̑I��"></dt>
      <dd>���͂���A���x�����@���̕ύX���ł��܂��B</dd>
      <dt><img src="images/shopping_step3_off.gif" alt="���������e�̊m�F"></dt>
      <dd>�ύX�������������������e���m�F���܂��B</dd>
      <dt><img src="images/shopping_step4_off.gif" alt="����������"></dt>
      <dd id="off">���������e�̕ύX���������܂��B<br>
�ύX�O�̂��������e���L�����Z������A�ύX���������܂����������ɂď���܂��B<br>
���o�^�̃��[���A�h���X���ĂɊm�F���[���𑗐M�������܂��̂œ��e�����m�F���������B</dd>
    </dl>
  </div>
</div>
<% End If %>

<% If FIRST_STEP = False Then	' 1st step �Ώ� %>
<% ' �Ē����|�b�v�A�b�v %>
<div class="overContent" id="overContent_reorder">
  <h2>���̒������Ē���</h2>
  <p>���̂��������e�Ɠ������e�ŁA�Ăт��������������܂��B<br>
���݃J�[�g�ɂ��鏤�i�ɒǉ����邩�A�J�[�g�̓��e���㏑�����邩���I�т��������B</p>

  <form name="f_reorder" method="post" action="OrderHistoryCopy.asp">
    <ul id="reorder_select">
      <li><input type="button" value="�J�[�g�ɒǉ����čĒ�������" class="strong_btn" onClick="reorder_onClick('Y');"></li>
      <li><input type="button" value="�J�[�g���㏑�����čĒ�������" class="strong_btn" onClick="reorder_onClick('N');"></li>
    </ul>
    <input type='hidden' name='OrderNo' value='<% = wOrderNo %>'>
    <input type='hidden' name='addItem' value='N'>
  </form>

  <ul class="back">
    <li><a href="javascript:void(0);" onClick="backclose();">��&nbsp;�߂�</a></li>
  </ul>
</div>
<% End If	' 1st step �Ώ� %>

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