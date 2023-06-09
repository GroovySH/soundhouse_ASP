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
'	�������폜����
'
'	OrderHistoryDetail.asp�̃L�����Z���{�^������Ăяo�����B
'	�T�v�F
'		�Y���̎󒍂̍폜�����Z�b�g����B
'		�Y���̎󒍖��ׂ̏��i�̍݌ɖ߂����s���B
'		Web�ύX��(Emax)�ɕύX�J�n���A�ύX�I�������Z�b�g����B
'
'	HTTPS�łȂ��ƃG���[
'	���O�C�����Ă��Ȃ��ƃG���[
'	���O�C�����Ă���΁ASession("userID")�Ɍڋq�ԍ����Z�b�g����Ă���B
'	Session("userID")���󕶎��̎��̓G���[�@����O�C�����Ă��������B�
'	Session("userID")�Ōڋq��񂪎�o���Ȃ���΃G���[�@����O�C�����Ă��������B�
'	�G���[���b�Z�[�W���Z�b�g��Login.asp��Redirect
'
'	�E�L�����Z���\���ԑсA�󒍂̃`�F�b�N
'	�E�f�[�^���o
'	�EEmax�̎󒍂̍X�V
'	�EEmax�̐F�K�i�ʍ݌ɂ̍X�V
'	�EEmax��Web�ύX�󒍂̓o�^
'	�EOrderHistory.asp���Ăяo��
'
'�ύX����
'2011/12/26 GV #1149 �V�K�쐬
'
'========================================================================
On Error Resume Next

Const THIS_PAGE_NAME = "OrderHistoryDelete.asp"
Const UPDATEABLE_STAFF_CD = "Internet"			' �L�����Z���E�����ύX �\�� ��.�S���҃R�[�h

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wUserID

Dim wNotLogin					' ���O�C�����Ă��Ȃ�
Dim wOrderUpdateStartTime		' �󒍕ύX�J�n����(�R���g���[���}�X�^)
Dim wOrderUpdateEndTime			' �󒍕ύX�I������(�R���g���[���}�X�^)

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

' �p�����[�^
wOrderNo = ReplaceInput(Trim(Request("OrderNo")))	' �󒍔ԍ�

If wOrderNo = "" Or IsNumeric(wOrderNo) = False Then
	wOrderNo = 0				' main �ŃG���[�Ƃ��Ď�舵��
Else
	wOrderNo = CLng(wOrderNo)
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

If wMsg <> "" Then
	'--- �w�b�_�[�����o���Ȃ�,�󒍂�������Ȃ�,�L�����Z���s���̏ꍇ�̓G���[�@OrderHistoryDetail.asp��Redirect
	Session("ErrMsg") = wMsg
	Response.Redirect "OrderHistoryDetail.asp?OrderNo=" & wOrderNo
End If

'--- ����I���̏ꍇ OrderHistory.asp �փ��b�Z�[�W�t���Ŗ߂�
Session("DispMsg") = "�������̓L�����Z������܂����B"
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
Dim vCurrentTime			' ���݂̎���
Dim vStaffCd				' �S���҃R�[�h

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
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      �󒍔ԍ� "
vSQL = vSQL & "    , �󒍓� "
vSQL = vSQL & "    , �x�����@ "
vSQL = vSQL & "    , Web�󒍕ύX�J�n�� "
vSQL = vSQL & "    , �S���҃R�[�h "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "�� WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        �󒍔ԍ� = " & wOrderNo & " "
vSQL = vSQL & "    AND �ڋq�ԍ� = " & wUserID & " "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	'--- �w�b�_�[�����o���Ȃ��ꍇ�̓G���[��Y���̒�����񂪂���܂���B��@OrderHistoryDetai.asp��Redirect
	vRS.Close
	Set vRS = Nothing
	wMsg = "�Y���̒�����񂪂���܂���B"
	Exit Function
End If

'--- �L�����Z���\���ԑсA�󒍂̃`�F�b�N
vCurrentTime = Time()
vStaffCd = LCase(vRS("�S���҃R�[�h") & "")

If isUpdateableTime(vCurrentTime, wOrderUpdateStartTime, wOrderUpdateEndTime) = False _
Or vStaffCd <> LCase(UPDATEABLE_STAFF_CD) _
Or IsNull(vRS("Web�󒍕ύX�J�n��")) = False Then
	'--- �L�����Z���s�̃I�[�_�[ �̏ꍇ�̓G���[����݂��̒����̃L�����Z���͍s���܂���B��@OrderHistoryDetai.asp��Redirect
	vRS.Close
	Set vRS = Nothing
	wMsg = "���݂��̒����̃L�����Z���͍s���܂���B"
	Exit Function
End If


'--- �g�����U�N�V�����J�n
Connection.BeginTrans


'--- Emax�̎󒍂̍X�V (�폜���̐ݒ�)
If update_order_deleteDate() = False Then
	Connection.RollbackTrans
	Exit Function
End If

'===========�R�����g 2011/12/22 hn
''--- Emax�̐F�K�i�ʍ݌ɂ̍X�V
'If update_Inventory(vRS("�x�����@") & "") = False Then
'	Connection.RollbackTrans
'	Exit Function
'End If
'===========�R�����g

'--- Emax��Web�ύX�󒍂̓o�^
If insert_WebUpdateOrder() = False Then
	Connection.RollbackTrans
	Exit Function
End If

'--- �L�����Z�����[�����M
If send_cancelMail() = False Then
	Connection.RollbackTrans
	Exit Function
End If


vRS.Close

'--- �g�����U�N�V�����I��
Connection.CommitTrans

Set vRS = Nothing

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
vSQL = vSQL & "    , b.�ڋq�X�֔ԍ� "
vSQL = vSQL & "    , b.�ڋq�s���{�� "
vSQL = vSQL & "    , b.�ڋq�Z�� "
vSQL = vSQL & "    , c.�ڋq�d�b�ԍ� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      Web�ڋq     a WITH (NOLOCK) "
vSQL = vSQL & "    , Web�ڋq�Z�� b WITH (NOLOCK) "
vSQL = vSQL & "    , Web�ڋq�Z���d�b�ԍ� c WITH (NOLOCK)"
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.�ڋq�ԍ� = " & wUserID & " "
vSQL = vSQL & "    AND a.Web�s�f�ڃt���O <> 'Y' "
vSQL = vSQL & "    AND b.�ڋq�ԍ� = a.�ڋq�ԍ� "
vSQL = vSQL & "    AND b.�Z���A�� = 1 "
vSQL = vSQL & "    AND c.�ڋq�ԍ� = a.�ڋq�ԍ� "
vSQL = vSQL & "    AND c.�Z���A�� = 1 "
vSQL = vSQL & "    AND c.�d�b�A�� = 1 "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

Set get_customer = vRS

End Function

'========================================================================
'
'	Function	�󒍏��̎��o��
'	Note		�L�����Z�����[�����M�p
'
'========================================================================
Function get_orderInfo()

Dim vRS
Dim vSQL

'---- �󒍏����o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.�󒍔ԍ� "
vSQL = vSQL & "    , a.���ϓ� "
vSQL = vSQL & "    , a.�x�����@ "
vSQL = vSQL & "    , a.�x�����@ "
vSQL = vSQL & "    , a.�ڋqE_mail "
vSQL = vSQL & "    , b.���[�J�[�R�[�h "
vSQL = vSQL & "    , b.���i�R�[�h "
vSQL = vSQL & "    , b.�F "
vSQL = vSQL & "    , b.�K�i "
vSQL = vSQL & "    , b.�󒍒P�� "
vSQL = vSQL & "    , b.�󒍐��� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "��     a WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "�󒍖��� b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.�󒍔ԍ� = " & wOrderNo & " "
vSQL = vSQL & "    AND a.�󒍔ԍ� = b.�󒍔ԍ� "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

Set get_orderInfo = vRS

End Function

'========================================================================
'
'	Function	���i���̎��o��
'	Note		�L�����Z�����[�����M�p
'
'========================================================================
Function get_itemInfo(pstrMaketCd, pstrItemCd)

Dim vRS
Dim vSQL

'---- ���i�����o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.���i�� "
vSQL = vSQL & "    , a.Web���i�t���O "
vSQL = vSQL & "    , b.���[�J�[�� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      Web���i  a WITH (NOLOCK) "
vSQL = vSQL & "    , ���[�J�[ b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.���[�J�[�R�[�h = '" & pstrMaketCd & "' "
vSQL = vSQL & "    AND a.���i�R�[�h     = '" & escapeSingleQuote(pstrItemCd) & "' "
vSQL = vSQL & "    AND a.���[�J�[�R�[�h = b.���[�J�[�R�[�h "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

Set get_itemInfo = vRS

End Function

'========================================================================
'
'	Function	Emax�̎󒍂̍폜���ݒ�
'
'========================================================================
Function update_order_deleteDate()

update_order_deleteDate = False

Dim vSQL
Dim vRS

vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      �폜�� "
vSQL = vSQL & "    , �ŏI�X�V�� "
vSQL = vSQL & "    , �ŏI�X�V�҃R�[�h "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "�� "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    �󒍔ԍ� = " & wOrderNo & " "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	wMsg = "�Y���̒�����񂪂���܂���B"
	vRS.Close
	Set vRS = Nothing
	Exit Function
End If

vRS("�폜��") = Now()
vRS("�ŏI�X�V��") = Now()
vRS("�ŏI�X�V�҃R�[�h") = UPDATEABLE_STAFF_CD

vRS.Update

vRS.Close
Set vRS = Nothing

update_order_deleteDate = True

End function

'========================================================================
'
'	Function	Emax�̐F�K�i�ʍ݌ɂ��X�V
'
'========================================================================
Function update_Inventory(pstrPaymetMethod)

update_Inventory = False

Dim vSQL
Dim vRS
Dim vRS_Inventory
Dim vBItem				' B�i�t���O
Dim vOrderNum			' �󒍐���
Dim vInventoryReservNum	' ���ψ������v����

'--- �󒍖��ׂ̎�o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      ���[�J�[�R�[�h "
vSQL = vSQL & "    , ���i�R�[�h "
vSQL = vSQL & "    , �F "
vSQL = vSQL & "    , �K�i "
vSQL = vSQL & "    , B�i�t���O "
vSQL = vSQL & "    , �󒍐��� "
vSQL = vSQL & "    , �󒍈������v���� "
vSQL = vSQL & "    , ���ψ������v���� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "�󒍖��� WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    �󒍔ԍ� = " & wOrderNo & " "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then
	wMsg = "�Y���̒�����񂪂���܂���B"
	vRS.Close
	Set vRS = Nothing
	Exit Function
End If

Do Until vRS.EOF

	vBItem = vRS("B�i�t���O") & ""

	'--- �X�V����u�F�K�i�ʍ݌Ɂv�̎�o��
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "      �����\���� "
	vSQL = vSQL & "    , �󒍈������� "
	vSQL = vSQL & "    , �󒍎c���� "
	vSQL = vSQL & "    , ���ώ�u���� "
	vSQL = vSQL & "    , B�i�����\���� "
	vSQL = vSQL & "    , B�i�󒍈������� "
	vSQL = vSQL & "    , B�i���ώ�u���� "
	vSQL = vSQL & "    , �ŏI�X�V�� "
	vSQL = vSQL & "    , �ŏI�X�V�҃R�[�h "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & "    " & gLinkServer & "�F�K�i�ʍ݌� "
	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "        ���[�J�[�R�[�h = '" & vRS("���[�J�[�R�[�h") & "' "
	vSQL = vSQL & "    AND ���i�R�[�h     = '" & escapeSingleQuote(vRS("���i�R�[�h")) & "' "
	vSQL = vSQL & "    AND �F             = '" & vRS("�F") & "' "
	vSQL = vSQL & "    AND �K�i           = '" & vRS("�K�i") & "' "

'@@@@@@Response.Write(vSQL)

	Set vRS_Inventory = Server.CreateObject("ADODB.Recordset")
	vRS_Inventory.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	If vRS_Inventory.EOF = False Then

		If pstrPaymetMethod = "�����" Then

			' �����

			If vBItem <> "Y" Then

				' Not B�i
				vRS_Inventory("�����\����") = Nz(vRS_Inventory("�����\����"), 0) + Nz(vRS("�󒍐���"), 0)
				vRS_Inventory("�󒍈�������") = Nz(vRS_Inventory("�󒍈�������"), 0) - Nz(vRS("�󒍐���"), 0)
				vRS_Inventory("�󒍎c����")   = Nz(vRS_Inventory("�󒍎c����"), 0)   - (Nz(vRS("�󒍐���"), 0) - Nz(vRS("�󒍈������v����"), 0))

			Else

				' B�i
				vRS_Inventory("B�i�����\����") = Nz(vRS_Inventory("B�i�����\����"), 0) + Nz(vRS("�󒍐���"), 0)
				vRS_Inventory("B�i�󒍈�������") = Nz(vRS_Inventory("B�i�󒍈�������"), 0) - Nz(vRS("�󒍐���"), 0)

			End If

		Else

			' Not �����

			If vBItem <> "Y" Then

				' Not B�i
				vRS_Inventory("�����\����") = Nz(vRS_Inventory("�����\����"), 0) + Nz(vRS("���ψ������v����"), 0)
				vRS_Inventory("���ώ�u����") = Nz(vRS_Inventory("���ώ�u����"), 0) - Nz(vRS("���ψ������v����"), 0)

			Else

				' B�i
				vRS_Inventory("B�i�����\����") = Nz(vRS_Inventory("B�i�����\����"), 0) + Nz(vRS("���ψ������v����"), 0)
				vRS_Inventory("B�i���ώ�u����") = Nz(vRS_Inventory("B�i���ώ�u����"), 0) - Nz(vRS("���ψ������v����"), 0)

			End If

		End If

		vRS_Inventory("�ŏI�X�V��") = Now()
		vRS_Inventory("�ŏI�X�V�҃R�[�h") = UPDATEABLE_STAFF_CD

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
'	Function	Emax��Web�ύX�󒍂̓o�^
'	Note		���R�[�h�����݂���΍X�V
'
'========================================================================
Function insert_WebUpdateOrder()

insert_WebUpdateOrder = False

Dim vSQL
Dim vRS

vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      �󒍔ԍ� "
vSQL = vSQL & "    , �ύX�J�n�� "
vSQL = vSQL & "    , �ύX�I���� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    " & gLinkServer & "Web�ύX�� "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    �󒍔ԍ� = " & wOrderNo & " "

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF Then

	' ���R�[�h�����ׁ̈A�o�^
	vRS.AddNew

	vRS("�󒍔ԍ�") = wOrderNo
	vRS("�ύX�J�n��") = Now()
	vRS("�ύX�I����") = Now()

Else

	' ���R�[�h�L��ׁ̈A�X�V
	vRS("�ύX�J�n��") = Now()
	vRS("�ύX�I����") = Now()

End If

vRS.Update

vRS.Close
Set vRS = Nothing

insert_WebUpdateOrder = True

End function

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
call getCntlMst("Web", "��", "�󒍕ύX�\���ԑ�", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

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
'	Function	Nz �֐�
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
'
'	Function	�L�����Z�����[�����M
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


'--- �R���g���[���}�X�^�������ŗ��擾
call getCntlMst("����", "����ŗ�", "1", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

vTaxRate = Clng(vdblItemNum1)

'--- �R���g���[���}�X�^���瑗�M��(From)�A�h���X�擾
call getCntlMst("����", "���M��Email", "Web�󒍒ʒm", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

'--- ���M��(From)
vEMailAddrFrom = vstrItemChar1

'--- �R���g���[���}�X�^����BCC�A�h���X�擾
call getCntlMst("����", "���M��Email", "ShopBCC", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

'--- BCC
vEMailAddrBCC = vstrItemChar1

'--- subject
vSubject = "�T�E���h�n�E�X�@�������L�����Z���m�F���[���i�����z�M�j[" & wUserID & "/Web-Emax/Web�󒍃L�����Z��]"

'--- �ڋq���擾
Set vRS = get_customer()
If vRS.EOF Then
	' �ڋq��񖳂�
	Exit Function
End If

'--- body
vBody = ""
vBody = vBody & "�T�E���h�n�E�X�E�I�����C���V���b�v�������p�������ɂ��肪�Ƃ��������܂��B" & vbNewLine
vBody = vBody & "���L�̂��������L�����Z������܂����B" & vbNewLine
vBody = vBody & vbNewLine
vBody = vBody & "�������������������@�������i�L�����Z���j������������������" & vbNewLine
vBody = vBody & vRS("�ڋq��") & " �l" & vbNewLine
vBody = vBody & "�Z���F ��" & vRS("�ڋq�X�֔ԍ�") & " " & vRS("�ڋq�s���{��") & vRS("�ڋq�Z��") & vbNewLine
vBody = vBody & "�d�b�ԍ��F " & vRS("�ڋq�d�b�ԍ�") & vbNewLine
vBody = vBody & "���q�l�ԍ��F " & wUserID & vbNewLine

vRS.Close

'--- �󒍏��擾
Set vRS = get_orderInfo()
If vRS.EOF Then
	' �󒍏�񖳂�
	Exit Function
End If

'--- ���M��(To)
vEMailAddrTo = vRS("�ڋqE_mail")

vBody = vBody & "���ϓ��t�F " & formatDateYYYYMMDD(vRS("���ϓ�")) & vbNewLine
vBody = vBody & "���ϔԍ��F " & wOrderNo & vbNewLine
vBody = vBody & "���x�����@�F " & vRS("�x�����@") & vbNewLine
vBody = vBody & vbNewLine

vBody = vBody & "�|�|�|�|�|�|�|�|�|�@�ځ@�@�ׁ@�|�|�|�|�|�|�|�|�|" & vbNewLine

Do Until vRS.EOF

	'--- ���i����o��
	Set vRS_Item = get_itemInfo(vRS("���[�J�[�R�[�h") & "", vRS("���i�R�[�h") & "")
	If vRS_Item.EOF = False Then
		vItemName = vRS_Item("���[�J�[��") & " " & vRS_Item("���i��")
	Else
		' ���i��񖳂�
		vItemName = ""
	End If
	vRS_Item.Close

	'--- �P���v�Z
	vUnitPrice = calcPrice(vRS("�󒍒P��"), vTaxRate)

	vBody = vBody & "���i���F " & vItemName & vbNewLine
	vBody = vBody & "���ʁF " & FormatNumber(vRS("�󒍐���"), 0) & vbNewLine
	vBody = vBody & "�P��(�ō�)�F \" & FormatNumber(vUnitPrice, 0) & vbNewLine
	vBody = vBody & "���z(�ō�)�F \" & FormatNumber(vUnitPrice * vRS("�󒍐���"), 0) & vbNewLine
	vBody = vBody & vbNewLine

	vRS.MoveNext

Loop

Set vRS_Item = Nothing

vRS.Close

Set vRS = Nothing

'--- �R���g���[���}�X�^����g���[���[�p������擾
call getCntlMst("Web", "Email", "�g���[��", vstrItemChar1, vstrItemChar2, vdblItemNum1, vdblItemNum2, vdatItemDate1, vdatItemDate2)

vBody = vBody & vstrItemChar1


'--- ���[�����M
Set vobjCBOMessage = Server.CreateObject("CDO.Message")

vobjCBOMessage.From = vEMailAddrFrom
vobjCBOMessage.To = vEMailAddrTo
vobjCBOMessage.BCC = vEMailAddrBCC
vobjCBOMessage.Subject = vSubject
vobjCBOMessage.TextBody = vBody
vobjCBOMessage.BodyPart.Charset = "iso-2022-jp"

'--- ���[���T�[�o�[�w�� (�s�v�ł���Έȉ�4�s�R�����g�A�E�g)
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
