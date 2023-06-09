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
'	�I�[�_�[���o�^
'
'------------------------------------------------------------------------
'	
'		���󒍏��̊�{���ڂ̓o�^
'		���󒍖��׏��̓o�^
'
'		�J�[�g�֓����{�^���ŌĂяo�����C�J�[�g�փf�[�^�[�Z�b�g��order.asp��
'
'------------------------------------------------------------------------

'	�X�V����
'2005/01/13 ���i���L�����ǂ����̃`�F�b�N�i�L���b�V����ʂ���̓o�^�Ή�)
'2005/02/16 �����萔�ʒP�����o�����̏��������@�����萔�ʁ�0��ǉ�
'2006/06/26 �p�Տ��i�̏ꍇ�A�����\���ȏ�Ɏ󒍂��Ȃ��悤�ɕύX
'2007/03/15 �p�����[�^�ɑ΂���ReplaceInput��ǉ�
'2007/04/18 B�i�ǉ�
'2007/07/05 Item�p�����[�^�i���[�J�[�R�[�h^���i�R�[�h^�F^�K�i)�擾�Ώ�
'2007/07/05 ���i�o�^��1�݂̂ɕύX
'2007/07/13 �F�K�i���菤�i�͐F�K�i���I�����ꂽ���ă`�F�b�N
'2008/05/23 ���̓f�[�^�`�F�b�N�����iLEFT, Numeric, EOF��)
'2008/12/16 On Error Resume Next �ǉ�
'2008/12/24 AdditionalProd��ǉ��i�������i�����o�^���\�F�ꏏ�ɍw���@�\)
'           B�i�t���O=Y�̏��i�̂Ƃ��̂�B�i�P���g�p�ɕύX
'2011/04/14 hn SessionID�֘A�ύX
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/09/16 an #1112 �؂蔄�菤�i�̏ꍇ�͍��Z�����ɕʖ��ׂƂ��ēo�^
'========================================================================

On Error Resume Next		'2008/12/16

Dim userID

Dim qt
Dim maker_cd
Dim product_cd
Dim iro
Dim kikaku

Dim item
Dim item_list()
Dim item_cnt

Dim AdditionalItem()
Dim AdditionalItemCnt

Dim Connection
Dim RS_order_header
Dim RS_order_detail
Dim RS_product
Dim RS

Dim w_sql
Dim w_msg
Dim w_html

Dim w_detail_cnt
Dim wErrDesc   '2011/08/01 an add

'=======================================================================

'---- execute main process

Session.Timeout = 20


Session("msg") = ""
w_msg = ""

call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderPreInsert.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

call close_db()

if Err.Description <> "" then				'2008/12/16
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

if w_msg <> "" then
	w_msg = "<font color='#ff0000'>" & w_msg & "</font>"
	Session("msg") = w_msg
end if

Response.Redirect "Order.asp"

'=======================================================================

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()

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
Dim i
Dim v_item

'---- ���M�f�[�^�[�̎��o��
qt = ReplaceInput(Trim(Request("qt")))
maker_cd = Left(ReplaceInput(Trim(Request("maker_cd"))), 8)
product_cd = Left(ReplaceInput(Trim(Request("product_cd"))), 20)
iro = Left(ReplaceInput(Trim(Request("iro"))), 20)
kikaku = Left(ReplaceInput(Trim(Request("kikaku"))), 20)

if isNumeric(qt) = false Or qt = "" then
	qt = 0
end if

'if qt > 10000 then
'	qt = 10000
'end if

item = ReplaceInput(Trim(Request("Item")))

if item <> "" then
	item_cnt = cf_unstring(item, item_list, "^")
	maker_cd = Left(ReplaceInput(Trim(item_list(0))), 8)
	product_cd = Left(ReplaceInput(Trim(item_list(1))), 20)
	if item_cnt > 2 then
		iro = Left(ReplaceInput(Trim(item_list(2))), 20)
		if item_cnt > 3 then
			kikaku = Left(ReplaceInput(Trim(item_list(3))), 20)
		end if
	end if
end if

if ReplaceInput(Trim(Request("AdditionalItem"))) <> "" then
	AdditionalItemCnt = cf_unstring(ReplaceInput(Trim(Request("AdditionalItem"))), AdditionalItem, ",")
end if

'---- ���󒍏��o�^
call insert_order_header()

'---- ��{���i�o�^
call insert_order_detail(maker_cd, product_cd, iro, kikaku, qt)

'---- �ǉ����i�o�^
for i=1 to AdditionalItemCnt-1
	if AdditionalItem(i) <> "" then
		item_cnt = cf_unstring(AdditionalItem(i), item_list, "^")
		maker_cd = Left(ReplaceInput(Trim(item_list(0))), 8)
		product_cd = Left(ReplaceInput(Trim(item_list(1))), 20)
		if item_cnt > 2 then
			iro = Left(ReplaceInput(Trim(item_list(2))), 20)
			if item_cnt > 3 then
				kikaku = Left(ReplaceInput(Trim(item_list(3))), 20)
			end if
		end if
	end if

	call insert_order_detail(maker_cd, product_cd, iro, kikaku, qt)
Next

End Function

'========================================================================
'
'	Function	insert order_header
'
'========================================================================
'
Function insert_order_header()
Dim i

'---- �Y��SessionID�ŉ��󒍂��o�^����Ă邩�`�F�b�N
w_sql = "SELECT *" _
	  & "  FROM ����" _
		& " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RS_order_header = Server.CreateObject("ADODB.Recordset")
RS_order_header.CursorType = adOpenDynamic
RS_order_header.LockType = adLockOptimistic
RS_order_header.Open w_sql, Connection

if RS_order_header.EOF = true then

	'---- insert ����
	RS_order_header.AddNew

	RS_order_header("SessionID") = gSessionID		'2011/04/14 hn mod
	RS_order_header("���͓�") = now()
	RS_order_header("�L���R�[�h") = Session("AdID")
	RS_order_header("�ŏI�X�V��") = now()

	RS_order_header.update
end if

RS_order_header.close

End function

'========================================================================
'
'	Function	insert ���󒍖���
'
'========================================================================
'
Function insert_order_detail(pMakerCd, pProductCd, pIro, pKikaku, pQt)
Dim i
Dim w_max_detail_no
Dim w_update_cnt
Dim wPrice

'---- MAX�󒍖��הԍ���o��
w_sql = "SELECT MAX(�󒍖��הԍ�) AS MAX�󒍖��הԍ�" _
	    & "  FROM ���󒍖���" _
	    & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	  
Set RS = Server.CreateObject("ADODB.Recordset")
RS.CursorType = adOpenDynamic
RS.LockType = adLockOptimistic
RS.Open w_sql, Connection

if RS.EOF = false then
	if isNULL(RS("MAX�󒍖��הԍ�")) = false then
		w_max_detail_no = RS("MAX�󒍖��הԍ�")
	else
		w_max_detail_no = 0
	end if
else
	w_max_detail_no = 0
end if

RS.close

w_update_cnt = 0

if pQt > 0 then
	w_update_cnt = w_update_cnt + 1

	'---- ���󒍖���Recordset�쐬
	w_sql = ""
	w_sql = w_sql & "SELECT *"
	w_sql = w_sql & "  FROM ���󒍖���"
	w_sql = w_sql & " WHERE SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod
	w_sql = w_sql & "   AND ���[�J�[�R�[�h = '" & pMakerCd & "'"
	w_sql = w_sql & "   AND ���i�R�[�h = '" & pProductCd & "'"
	w_sql = w_sql & "   AND �F = '" & pIro & "'"
	w_sql = w_sql & "   AND �K�i = '" & pKikaku & "'"
		  
	Set RS_order_detail = Server.CreateObject("ADODB.Recordset")
	RS_order_detail.CursorType = adOpenDynamic
	RS_order_detail.LockType = adLockOptimistic
	RS_order_detail.Open w_sql, Connection

'@@@@@		response.write(w_sql)

	'---- ���i����o��
	w_sql = ""
	w_sql = w_sql & "SELECT a.���[�J�[�R�[�h"
	w_sql = w_sql & "     , a.���i�R�[�h"
	w_sql = w_sql & "     , a.���i��"
	w_sql = w_sql & "     , CASE"
	w_sql = w_sql & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
	w_sql = w_sql & "         ELSE a.�̔��P��"
	w_sql = w_sql & "       END AS �̔��P��"
	w_sql = w_sql & "     , a.B�i�P��"	
	w_sql = w_sql & "     , a.�����萔��"	
	w_sql = w_sql & "     , a.������󒍍ϐ���"	
	w_sql = w_sql & "     , a.ASK���i�t���O"
	w_sql = w_sql & "     , a.B�i�t���O"
	w_sql = w_sql & "     , a.�p�ԓ�"
	w_sql = w_sql & "     , a.�؂蔄��t���O"    '2011/09/16 an add
	w_sql = w_sql & "     , b.���[�J�[��"
	w_sql = w_sql & "     , c.�����\����"
	w_sql = w_sql & "     , c.B�i�����\����"

	'�F�K�i�����邩�ǂ��� 2007/07/13
	w_sql = w_sql & "     , (SELECT COUNT(*)"
	w_sql = w_sql & "          FROM Web�F�K�i�ʍ݌� t"
	w_sql = w_sql & "         WHERE t.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
	w_sql = w_sql & "           AND t.���i�R�[�h = a.���i�R�[�h"
	w_sql = w_sql & "           AND (t.�F != '' OR t.�K�i != '')"
	w_sql = w_sql & "           AND t.�I���� IS NULL"
	w_sql = w_sql & "       ) AS �F�K�iCNT"

	w_sql = w_sql & "  FROM Web���i a"
	w_sql = w_sql & "     , ���[�J�[ b"
	w_sql = w_sql & "     , Web�F�K�i�ʍ݌� c"
	w_sql = w_sql & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
	w_sql = w_sql & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
	w_sql = w_sql & "   AND c.���i�R�[�h = a.���i�R�[�h"
	w_sql = w_sql & "   AND c.�F = '" & pIro & "'"
	w_sql = w_sql & "   AND c.�K�i = '" & pKikaku & "'"
	w_sql = w_sql & "   AND a.���[�J�[�R�[�h = '" & pMakerCd & "'"
	w_sql = w_sql & "   AND a.���i�R�[�h = '" & pProductCd & "'"
	w_sql = w_sql & "   AND a.Web���i�t���O = 'Y'"
		  
	Set RS_product = Server.CreateObject("ADODB.Recordset")
	RS_product.CursorType = adOpenDynamic
	RS_product.LockType = adLockOptimistic
	RS_product.Open w_sql, Connection

	if RS_product.EOF = true then
		w_msg = w_msg & pProductCd & "�́A�戵�����Ă���܂���B<br>"
	else
		if RS_product("�F�K�iCNT") > 0 AND pIro ="" AND pKikaku = "" then
			w_msg = w_msg & "�F�K�i��I�����Ă��������B<br>"
		else
			'---- ���o�^���i�A�؂蔄��t���OY�̏ꍇ�͉��󒍖��ׂ�ǉ�
			if RS_order_detail.EOF = true OR RS_product("�؂蔄��t���O") = "Y" then    '2011/09/16 an mod
				if isNull(RS_product("�p�ԓ�")) = false AND RS_product("�����\����") < CLng(pQt) then
					w_msg = w_msg & pProductCd & "�́A�݌ɂ�" & RS_product("�����\����") & "��������܂���B�@���ʂ�ύX���Ă��������������B<br>"
				else
					if RS_product("B�i�t���O") = "Y" AND RS_product("B�i�����\����") > 0 AND RS_product("B�i�����\����") < CLng(pQt) then
						w_msg = w_msg & pProductCd & "�́A�݌ɂ�" & RS_product("B�i�����\����") & "��������܂���B�@���ʂ�ύX���Ă��������������B<br>"
					else
						'---- insert ���󒍖���
						w_max_detail_no = w_max_detail_no + 1

						if RS_product("B�i�t���O") = "Y" then
							wPrice = RS_product("B�i�P��")
						else
							wPrice = RS_product("�̔��P��")
						end if

						RS_order_detail.AddNew
						RS_order_detail("SessionID") = gSessionID		'2011/04/14 hn mod
						RS_order_detail("�󒍖��הԍ�") = w_max_detail_no
						RS_order_detail("���[�J�[�R�[�h") = pMakerCd
						RS_order_detail("���i�R�[�h") = pProductCd
						RS_order_detail("�F") = pIro
						RS_order_detail("�K�i") = pKikaku
						RS_order_detail("���[�J�[��") = RS_product("���[�J�[��")
						RS_order_detail("���i��") = RS_product("���i��")
						RS_order_detail("�󒍒P��") = wPrice
						RS_order_detail("�󒍐���") = Clng(pQt)
						RS_order_detail("�󒍋��z") = Fix(RS_order_detail("�󒍒P��")) * Clng(pQt)

						if RS_product("�����萔��") > RS_product("������󒍍ϐ���") then
							RS_order_detail("������P���t���O") = "Y"
						else
							RS_order_detail("������P���t���O") = ""
						end if

						RS_order_detail("B�i�t���O") = RS_product("B�i�t���O")

						RS_order_detail.Update
					end if
				end if
			'---- �o�^�ςݏ��i�͎󒍐��ʂ�ǉ�����i�؂蔄�菤�i�ȊO�j
			else
				if isNull(RS_product("�p�ԓ�")) = false AND RS_product("�����\����") < RS_order_detail("�󒍐���") + Clng(pQt) then
					w_msg = w_msg & pProductCd & "�́A�݌ɂ�" & RS_product("�����\����") & "��������܂���B�@���ʂ�ύX���Ă��������������B<br>"
				else
					if RS_product("B�i�����\����") > 0 AND RS_product("B�i�����\����") < CLng(pQt) then
						w_msg = w_msg & pProductCd & "�́A�݌ɂ�" & RS_product("B�i�����\����") & "��������܂���B�@���ʂ�ύX���Ă��������������B<br>"
					else
						'---- update ���󒍖���
						RS_order_detail("�󒍐���") = RS_order_detail("�󒍐���") + Clng(pQt)
						RS_order_detail("�󒍋��z") = Fix(RS_order_detail("�󒍒P��")) * RS_order_detail("�󒍐���")
						RS_order_detail.Update
					end if
				end if
			end if
		end if
	end if
end if

if w_update_cnt > 0 then
	RS_order_detail.close
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

%>
