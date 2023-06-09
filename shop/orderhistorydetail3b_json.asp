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
<!--#include file="../3rdParty/aspJSON1.17.asp"-->
<%
'========================================================================
'
'	�w�������ꗗ�y�[�W
'
'
'�ύX����
'2014/09/16 GV �V�K�쐬
'2015.03.09 GV �o�ׂ��ʂꂽ�ꍇ�̕\���s��Ή��B
'2016.05.09 GV �Z�b�g�i�ŏo�ׂ��ʂꂽ�ꍇ�̕\���s��Ή��B
'========================================================================
'On Error Resume Next

Dim Connection
Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wUserID

Dim oJSON						' JSON�I�u�W�F�N�g
Dim wOrderNo					' �󒍔ԍ�

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
' Get�p�����[�^
'wUserID = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("order_no")))

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()

Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   ' member�n�̑��̃y�[�W�����ɂȂ炤
If Err.Description <> "" Then
'	wErrDesc = THIS_PAGE_NAME & " " & Replace(Replace(Err.Description, vbCR, " "), vbLF, " ")
'	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If

Call close_db()

If Err.Description <> "" Then

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
Dim vRS

Dim orderDate
Dim shippingDate
Dim estimateDate
Dim one_time_todokesaki
Dim final_nouki_date_time
Dim receiptName
Dim receiptNote
Dim webOutline
Dim source
Dim shippingText
Dim itemPicSmall
Dim slipNo
Dim i
Dim invoiceNo
Dim shippingNo
Dim makerName
Dim itemName
Dim setItem '2016.05.09 GV add
Dim shipSuu '2016.05.09 GV add

Set oJSON = New aspJSON


one_time_todokesaki = ""
final_nouki_date_time = ""
receiptName = ""
receiptNote = ""
webOutline = ""
source = ""
shippingText = ""
itemPicSmall = ""
slipNo = ""
invoiceNo = ""
shippingNo = ""
makerName = ""
itemName = ""
setItem = "" '2016.05.09 GV add
i = 0

'--- �o�׊����f�[�^�̏���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      b.�󒍖��הԍ� "
vSQL = vSQL & "    , b.���[�J�[�R�[�h "
vSQL = vSQL & "    , b.���i�R�[�h "
vSQL = vSQL & "    , b.�F "
vSQL = vSQL & "    , b.�K�i "
vSQL = vSQL & "    , z.���iID "		'2015.03.09 GV add
vSQL = vSQL & "    , b.�󒍒P�� "
vSQL = vSQL & "    , b.�󒍐��� "
vSQL = vSQL & "    , b.�󒍈������v���� "
vSQL = vSQL & "    , b.�o�׍��v���� "
vSQL = vSQL & "    , sum(f.�o�א���) AS �o�א��� "
vSQL = vSQL & "    , c.���[�J�[�� "
vSQL = vSQL & "    , d.���i�� "
vSQL = vSQL & "    , d.���i�T��Web "
vSQL = vSQL & "    , d.���i�摜�t�@�C����_�� "
vSQL = vSQL & "    , d.Web���i�t���O "
vSQL = vSQL & "    , d.�Z�b�g���i�t���O " '2016.05.09 GV add
'vSQL = vSQL & "    , e.�����ԍ� "	'2015.03.09 GV mod
vSQL = vSQL & "    , NULL AS �o�ח\��� "
vSQL = vSQL & "    , NULL AS �o�ח\��e�L�X�g "
vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "�󒍖���     b WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "���[�J�[     c WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "���i         d WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "�󒍑����   e WITH (NOLOCK) "	'2015.03.09 GV mod
vSQL = vSQL & "    , " & gLinkServer & "�o�ז���View f WITH (NOLOCK) "
vSQL = vSQL & "    , " & gLinkServer & "�F�K�i�ʍ݌� z WITH (NOLOCK) "	'2015.03.09 GV add
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        c.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
vSQL = vSQL & "    AND d.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
vSQL = vSQL & "    AND d.���i�R�[�h = b.���i�R�[�h "
vSQL = vSQL & "    AND e.�󒍔ԍ� = b.�󒍔ԍ� "	'2015.03.09 GV mod
vSQL = vSQL & "    AND f.�o�הԍ� = e.�o�הԍ� "	'2015.03.09 GV mod
vSQL = vSQL & "    AND f.�󒍔ԍ� = b.�󒍔ԍ� "
vSQL = vSQL & "    AND f.�󒍖��הԍ� = b.�󒍖��הԍ� "
vSQL = vSQL & "    AND f.�Z�b�g�i�e���הԍ� = 0 "
'2015.03.09 GV add start
vSQL = vSQL & "    AND z.���[�J�[�R�[�h = b.���[�J�[�R�[�h "
vSQL = vSQL & "    AND z.���i�R�[�h = b.���i�R�[�h "
vSQL = vSQL & "    AND z.�F = b.�F "
vSQL = vSQL & "    AND z.�K�i = b.�K�i "
'2015.03.09 GV add end
vSQL = vSQL & "    AND b.�󒍔ԍ� = " & wOrderNo & " "
vSQL = vSQL & "GROUP BY  "
vSQL = vSQL & "      b.�󒍖��הԍ� "
vSQL = vSQL & "    , b.���[�J�[�R�[�h "
vSQL = vSQL & "    , b.���i�R�[�h "
vSQL = vSQL & "    , b.�F "
vSQL = vSQL & "    , b.�K�i "
vSQL = vSQL & "    , z.���iID "		'2015.03.09 GV add
vSQL = vSQL & "    , b.�󒍒P�� "
vSQL = vSQL & "    , b.�󒍐��� "
vSQL = vSQL & "    , b.�󒍈������v���� "
vSQL = vSQL & "    , b.�o�׍��v���� "
vSQL = vSQL & "    , c.���[�J�[�� "
vSQL = vSQL & "    , d.���i�� "
vSQL = vSQL & "    , d.���i�T��Web "
vSQL = vSQL & "    , d.���i�摜�t�@�C����_�� "
vSQL = vSQL & "    , d.Web���i�t���O "
vSQL = vSQL & "    , d.�Z�b�g���i�t���O " '2016.05.09 GV add
vSQL = vSQL & "ORDER BY  "
'vSQL = vSQL & "      e.�����ԍ� "	'2015.03.09 GV mod
vSQL = vSQL & "      b.�󒍖��הԍ� "	'2015.03.09 GV add
vSQL = vSQL & "    , c.���[�J�[�� "
vSQL = vSQL & "    , d.���i�� "

'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' ���X�g�ǉ�
	oJSON.data.Add "list" ,oJSON.Collection()

	Do Until vRS.EOF = True
		' �o�ח\���
		If (IsNull(vRS("�o�ח\���"))) Then
			shippingDate = ""
		Else
			shippingDate = CStr(Trim(vRS("�o�ח\���")))
		End If


		If (IsNull(vRS("���i�T��Web"))) Then
			webOutline = ""
		Else
			webOutline = CStr(vRS("���i�T��Web"))
			webOutline = Replace(Trim(webOutline), """", "�h")
			webOutline = Replace(Trim(webOutline), vbCrLf, "")
		End If

		If (IsNull(vRS("�o�ח\��e�L�X�g"))) Then
			shippingText = ""
		Else
			shippingText = CStr(vRS("�o�ח\��e�L�X�g"))
			shippingText = Replace(Trim(shippingText), """", "�h")
		End If

		If (IsNull(vRS("���i�摜�t�@�C����_��"))) Then
			itemPicSmall = ""
		Else
			itemPicSmall = CStr(vRS("���i�摜�t�@�C����_��"))
		End If

		'2015.03.09 GV mod start
		'If (IsNull(vRS("�����ԍ�"))) Then
		'	shippingNo = ""
		'Else
		'	shippingNo = CStr(vRS("�����ԍ�"))
		'End If
		'2015.03.09 GV mod start

		makerName = Replace(Trim(vRS("���[�J�[��")), """", "�h")
		makerName = CStr(makerName)

		itemName = Replace(Trim(vRS("���i��")), """", "�h")
		itemName = CStr(itemName)

		'2016.05.09 GV add start
		If (IsNull(vRS("�Z�b�g���i�t���O"))) Then
			setItem = ""
		Else
			setItem = CStr(vRS("�Z�b�g���i�t���O"))
		End If

		' �Z�b�g�i�̏ꍇ
		If setItem = "Y" Then
			shipSuu = CDbl(vRS("�o�׍��v����"))
		Else
			shipSuu = CDbl(vRS("�o�א���"))
		End If
		'2016.05.09 GV add end

		With oJSON.data("list")
			'.Add i ,oJSON.Collection()
			'With .item(i)
			.Add "d"&CStr(vRS("�󒍖��הԍ�")) ,oJSON.Collection()
			With .item("d"&CStr(vRS("�󒍖��הԍ�")))
				.Add "order_detail_no", CStr(Trim(vRS("�󒍖��הԍ�")))
				.Add "maker_cd", CStr(vRS("���[�J�[�R�[�h"))
				.Add "item_cd", CStr(vRS("���i�R�[�h"))
				.Add "iro", CStr(Trim(vRS("�F")))
				.Add "kikaku",  CStr(Trim(vRS("�K�i")))
				.Add "item_id",  CStr(Trim(vRS("���iID")))	'2015.03.09 GV add
				.Add "order_tanka", CDbl(Trim(vRS("�󒍒P��")))
				.Add "order_suu", CDbl(vRS("�󒍐���")) 
				.Add "total_order_hikiate_suu", CDbl(vRS("�󒍈������v����"))
				.Add "total_shipping_suu", CDbl(vRS("�o�׍��v����"))
				'.Add "shipping_suu", CDbl(vRS("�o�א���"))		'2015.03.09 GV add
				.Add "shipping_suu", shipSuu  '2016.05.09 GV add
				.Add "maker_name", makerName
				.Add "item_name", itemName
				.Add "web_outline", webOutline
				.Add "item_pic_small", itemPicSmall
				.Add "web_flag", CStr(vRS("Web���i�t���O"))
				.Add "set_item", setItem '2016.05.09 GV add
				.Add "shipping_yotei_text", shippingText
				.Add "invoice_no", invoiceNo
				.Add "shipping_no", shippingNo
			End With
		End With

		i = i + 1

		vRS.MoveNext
	Loop





End If

'���R�[�h�Z�b�g�����
vRS.Close

'���R�[�h�Z�b�g�̃N���A
Set vRS = Nothing

' -------------------------------------------------
' JSON�f�[�^�̕ԋp
' -------------------------------------------------
' �w�b�_�o��
Response.AddHeader "Content-Type", "application/json"
' JSON�f�[�^�̏o��
Response.Write oJSON.JSONoutput()

End Function
%>
