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
'	Emax�󒍖��ׁ@�擾API
'
'
'�ύX����
'2016/03/29 GV �V�K�쐬
'2016.09.06 GV �L�����Z�����̈������߂������̉��C�Ή��B
'2020.02.28 GV �N�[�|���ƃ|�C���g�K�p����(�����\���`�F�b�N)�Ή��B
'2020.11.20 GV �����ύX���̔z�����w��`�F�b�N���C�B(#2602)
'
'========================================================================
'On Error Resume Next

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
wUserID = ReplaceInput(Trim(Request("cno")))
wOrderNo = ReplaceInput(Trim(Request("ono")))

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

Set ConnectionEmax = Server.CreateObject("ADODB.Connection")
ConnectionEmax.Open g_connectionEmax

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

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
Dim i
Dim j

Dim iro
Dim kikaku
Dim makerName
Dim itemName
Dim kosuuGenteiTankaFlg
Dim bItemFlg
Dim estimateHikiateSuu ' 2016.09.06 GV add
Dim hikiateSuuAtOrder  ' 2020.02.28 GV add

Set oJSON = New aspJSON
i = 0
j = 0

'--- ���ו����̏���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "  od.�󒍔ԍ� "
vSQL = vSQL & " ,od.�󒍖��הԍ� "
vSQL = vSQL & " ,od.���[�J�[�R�[�h "
vSQL = vSQL & " ,od.���i�R�[�h "
vSQL = vSQL & " ,od.�F "
vSQL = vSQL & " ,od.�K�i "
vSQL = vSQL & " ,mk.���[�J�[�� "
vSQL = vSQL & " ,od.���i�� "
vSQL = vSQL & " ,od.�󒍐��� "
vSQL = vSQL & " ,od.�󒍒P�� "
vSQL = vSQL & " ,od.�󒍋��z "
vSQL = vSQL & " ,od.������P���t���O "
vSQL = vSQL & " ,od.B�i�t���O "
vSQL = vSQL & " ,od.���ψ������v���� " ' 2016.09.06 GV add
vSQL = vSQL & " ,od.�󒍎������\�݌ɐ��� " ' 2016.09.06 GV add

vSQL = vSQL & "FROM "
vSQL = vSQL & "      " & gLinkServer & "�� o WITH (NOLOCK) "
vSQL = vSQL & "INNER JOIN " & gLinkServer & "�󒍖��� od WITH (NOLOCK) "
vSQL = vSQL & "  ON od.�󒍔ԍ� =o.�󒍔ԍ� "
vSQL = vSQL & "  AND od.�Z�b�g�i�e���הԍ� = 0 "
vSQL = vSQL & "INNER JOIN " & gLinkServer & "���[�J�[ mk WITH (NOLOCK) "
vSQL = vSQL & "  ON mk.���[�J�[�R�[�h = od.���[�J�[�R�[�h "

vSQL = vSQL & "WHERE "
vSQL = vSQL & "      o.�󒍔ԍ� = " & wOrderNo & " "
vSQL = vSQL & "  AND o.�ڋq�ԍ� = " & wUserID & " "
vSQL = vSQL & "  AND od.�󒍐��� > 0 "
vSQL = vSQL & "  AND od.�󒍋��z > 0 "

vSQL = vSQL & " ORDER BY "
vSQL = vSQL & "        od.�󒍖��הԍ� ASC "


'@@@@Response.Write(vSQL)

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

If vRS.EOF = False Then

	' ���X�g�ǉ�
	oJSON.data.Add "list" ,oJSON.Collection()

	' --------------------
	For i = 0 To (vRS.RecordCount - 1)
		'�F
		iro = Replace(Trim(vRS("�F")), """", "�h")
		iro = CStr(iro)

		'�K�i
		kikaku = Replace(Trim(vRS("�K�i")), """", "�h")
		kikaku = CStr(kikaku)

		'���[�J�[��
		makerName = Replace(Trim(vRS("���[�J�[��")), """", "�h")
		makerName = CStr(makerName)

		'���i��
		'itemName = Replace(Trim(vRS("���i��")), """", "�h")
		'itemName = CStr(itemName)

		itemName = Replace(Trim(vRS("���i��")), """", "�h")
		itemName = CStr(itemName)

		If (IsNull(vRS("������P���t���O"))) Then
			kosuuGenteiTankaFlg = ""
		Else
			kosuuGenteiTankaFlg = CStr(Trim(vRS("������P���t���O")))
		End If

		If (IsNull(vRS("B�i�t���O"))) Then
			bItemFlg = ""
		Else
			bItemFlg = CStr(Trim(vRS("B�i�t���O")))
		End If

		If (IsNull(vRS("�󒍎������\�݌ɐ���"))) Then
			hikiateSuuAtOrder = 0
		Else
			hikiateSuuAtOrder = CDbl(vRS("�󒍎������\�݌ɐ���"))
		End If

		'2020.11.20 GV add
		If (IsNull(vRS("���ψ������v����"))) Then
			estimateHikiateSuu = 0
		Else
			estimateHikiateSuu = CDbl(vRS("���ψ������v����"))
		End If


		'--- ���׍s����
		With oJSON.data("list")
			.Add j ,oJSON.Collection()
			With .item(j)
				.Add "o_no" ,CStr(Trim(vRS("�󒍔ԍ�")))
				.Add "od_no" ,CStr(Trim(vRS("�󒍖��הԍ�")))
				.Add "m_cd" ,CStr(Trim(vRS("���[�J�[�R�[�h")))
				.Add "i_cd" ,CStr(Trim(vRS("���i�R�[�h")))
				.Add "iro" ,iro
				.Add "kikaku" ,kikaku
				.Add "m_nm" ,makerName
				.Add "i_nm" ,itemName
				.Add "i_suu", CDbl(vRS("�󒍐���")) 
				.Add "i_tanka", CDbl(Trim(vRS("�󒍒P��")))
				.Add "i_am", CDbl(Trim(vRS("�󒍋��z")))
				.Add "kosuu_lmt", kosuuGenteiTankaFlg '������P���t���O
				.Add "b_item", bItemFlg 'B�i�t���O
				.Add "est_hikiate_suu", estimateHikiateSuu  ' 2020.11.20 GV add
				.Add "hikiate_suu_at_order", hikiateSuuAtOrder ' 2020.04.15 GV add
			End With
		End With

		' ���̃��R�[�h�s�ֈړ�
		vRS.MoveNext

		If vRS.EOF Then
			Exit For
		End If

		j = j + 1
	Next


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
Response.AddHeader "X-Content-Type-Options", "nosniff"

' JSON�f�[�^�̏o��
Response.Write oJSON.JSONoutput()

End Function
%>
