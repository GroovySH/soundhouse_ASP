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

<%
'========================================================================
'
'	�E�B�b�V�����X�g�֒ǉ�
'
'�X�V����
'2007/08/23 ���i�A�N�Z�X�����o�^�i�E�B�b�V�����X�g�j
'2007/09/10 ���i�A�N�Z�X�����o�^�i�E�B�b�V�����X�g�j�����ʂɕύX
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2009/04/30 �G���[����error.asp�ֈړ�
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'
'========================================================================

On Error Resume Next

Dim userID

Dim OrderDetailNo
Dim Item
Dim ItemCnt
Dim ItemList()
Dim MakerCd
Dim ProductCd
Dim Iro
Dim Kikaku

Dim Connection
Dim RS

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

'========================================================================

Response.buffer = true

'---- UserID ���o��
userID = Session("userID")

'---- �Ăяo��������̃f�[�^���o��
OrderDetailNo = ReplaceInput(Request("OrderDetailNo"))
Item = ReplaceInput(Request("Item"))

if Item <> "" then
	ItemCnt = cf_unstring(Item, ItemList, "^")
	MakerCd = Left(ItemList(0), 8)
	ProductCd = Left(ItemList(1), 20)
	Iro = Left(ItemList(2), 20)
	Kikaku = Left(ItemList(3), 20)
end if

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "WishListAdd.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then	
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

Response.Redirect "WishList.asp?msg=" & wMSG

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()
Dim i

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Main
'
'========================================================================
'
Function main()

Dim RSv
Dim vWishListAdded
Dim vYYYYMM

'---- �E�B�b�V�����X�g�o�^
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM �E�B�b�V�����X�g"
wSQL = wSQL & " WHERE �ڋq�ԍ� = " & userID
wSQL = wSQL & "   AND ���[�J�[�R�[�h = '" & MakerCd & "'"
wSQL = wSQL & "   AND ���i�R�[�h = '" & ProductCd & "'"
wSQL = wSQL & "   AND �F = '" & Iro & "'"
wSQL = wSQL & "   AND �K�i = '" & Kikaku & "'"
	  
Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	RSv.AddNew

	RSv("�ڋq�ԍ�") = userID
	RSv("���[�J�[�R�[�h") = MakerCd
	RSv("���i�R�[�h") = ProductCd
	RSv("�F") = Iro
	RSv("�K�i") = Kikaku

	vWishListAdded = "Y"
end if

RSv("�o�^��") = now()

RSv.Update
RSv.close

'---- ���󒍖��׍폜
if OrderDetailNo <> "" and isNumeric(OrderDetailNo) = true then
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM ���󒍖���"
	wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	wSQL = wSQL & "   AND �󒍖��הԍ� = " & OrderDetailNo
		  
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	if RSv.EOF = false then
		RSv.Delete
	end if

	RSv.close
end if

'---- ���i�A�N�Z�X�����o�^�i�E�B�b�V�����X�g�j
if vWishListAdded = "Y" then
	vYYYYMM = Year(Now()) & Right("0" & Month(Now()),2)
	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM ���i�A�N�Z�X����"
	wSQL = wSQL & " WHERE ���[�J�[�R�[�h = '" & MakerCd & "'"
	wSQL = wSQL & "   AND ���i�R�[�h = '" & ProductCd & "'"
	wSQL = wSQL & "   AND �N�� = '" & vYYYYMM & "'"
		  
	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	if RSv.EOF = true then
		RSv.AddNew

		RSv("���[�J�[�R�[�h") = MakerCd
		RSv("���i�R�[�h") = ProductCd
		RSv("�N��") = vYYYYMM
		RSv("�E�B�b�V�����X�g����") = 1
	else
		RSv("�E�B�b�V�����X�g����") = RSv("�E�B�b�V�����X�g����") + 1
	end if

	RSv.Update
	RSv.close
end if

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.close
Set Connection= Nothing    '2011/08/01 an add

End function

'========================================================================
%>
