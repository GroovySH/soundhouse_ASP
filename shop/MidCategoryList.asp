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
<!--#include file="../common/SearchListCommon.inc"-->
<!--#include file="./MidCategoryList/MidCategoryList.inc"-->
<%
'========================================================================
'
'	���J�e�S���[�ꗗ�y�[�W
'
'�X�V����
'2012/08/21 ok ���j���[�A���ɔ����V�f�U�C���ɐV�K�쐬
'2012/09/03 GV #1426 ��E���J�e�S����ʂŕ\�������SALES&OUTLET���̕\���f�[�^����ӂɎ擾�E�\������
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================
On Error Resume Next

Response.Buffer = True

Dim MidCategoryCd
Dim wLargeCategoryCd
Dim wLargeCategoryName
Dim wMidCategoryName
Dim wMidCategoryOverview
Dim wMetaTag
Dim wNoData
Dim wSalesTaxRate
Dim wErrDesc

Dim wMidCategoryComment

Dim wNaviMakerHTML				' (��)NAVI�p
Dim wNaviCategoryHTML			' (��)NAVI�p
Dim wNaviPricerangeHTML			' (��)NAVI�p
Dim s_category_cd				' (��)NAVI�p
Dim s_mid_category_cd			' (��)NAVI�p
Dim s_large_category_cd			' (��)NAVI�p
Dim s_maker_cd					' (��)NAVI�p
Dim sPriceFrom					' (��)NAVI�p
Dim sPriceTo					' (��)NAVI�p

Dim wInsertHTMLPath1
Dim wInsertHTMLPath2
Dim wStaticHTML(2)
Dim wSaleAndOutletHTML
Dim wChumokuHTML

Dim Connection
Dim RS
Dim wHTML

'=======================================================================
'	�󂯓n�������o��
'=======================================================================
MidCategoryCd = ReplaceInput(Trim(Request("MidCategoryCd")))

'=======================================================================
'	Execute main
'=======================================================================
Call connect_db()
Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^
If Err.Description <> "" Then
	wErrDesc = "MidCategoryList.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If

Call close_db()

If wNoData = "Y" Or Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End Function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection = Nothing

End Function

'========================================================================
'
'	Function	Main
'
'========================================================================
Function main()

Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2
Dim vSQL
Dim vSQLMaker					' (��)NAVI�p
Dim vSQLCategory				' (��)NAVI�p
Dim vSQLMiddleCategory			' (��)NAVI�p
Dim vSQLLargeCategory			' (��)NAVI�p
Dim vSQLPricerange				' (��)NAVI�p
Dim vFilePath
Dim vMsg

' ����ŗ�
Call getCntlMst("����", "����ŗ�", "1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)
wSalesTaxRate = CLng(vItemNum1)

' ���J�e�S�����擾
Call MidCategoryInfo()

If wNoData <> "Y" Then
	' �ς񂭂��E���J�e�S���[�ɂ��āE�J�e�S���[���X�g�E�������ߏ��i�E�j���[�X �ÓIHTML�t�@�C���̑��݃`�F�b�N (�L�������؂�`�F�b�N)
	If fExistMidCategoryStaticMainHTMLFile(MidCategoryCd) = False Then

		' �ς񂭂��E���J�e�S���[�ɂ��āE�J�e�S���[���X�g�E�������ߏ��i�E�j���[�X �ÓIHTML�e�L�X�g�t�@�C���쐬
		If fMakeMidCategoryStaticMainHTMLFile(MidCategoryCd, vFilePath, vMsg) = False Then
			Exit Function
		End If

	End If
End If

Call fCreateGetProductsSQL(  "cg" _
                           , "" _
                           , "" _
                           , "" _
                           , "" _
                           , "" _
                           , "" _
                           , MidCategoryCd _
                           , "" _
                           , "" _
                           , "" _
                           , "" _
                           , vSQL _
                           , vSQLMaker _
                           , vSQLCategory _
                           , vSQLMiddleCategory _
                           , vSQLLargeCategory _
                           , vSQLPricerange)
'--- NAVI�p�p�����[�^�Z�b�g
s_large_category_cd = ""
s_mid_category_cd = MidCategoryCd


' ��NAVI�p ���[�J�[�[�ꗗ�쐬
Call fCreateNAVIMaker2HTML(vSQLMaker, wNaviMakerHTML)

' ��NAVI�p �J�e�S���[�ꗗ�쐬
Call fCreateNAVICategory2HTML(vSQLCategory, wNaviCategoryHTML)

' ��NAVI�p ���i�ёI���쐬
Call fCreateNAVIPriceRange2HTML(vSQLPriceRange, wSalesTaxRate, wNaviPriceRangeHTML)

End Function

'========================================================================
'
'	Function	���J�e�S�����擾
'
'========================================================================
Function MidCategoryInfo()

Dim vFilePath
Dim vMsg
Dim vSQL

'---- �J�e�S���[ ���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.���J�e�S���[�����{�� "
vSQL = vSQL & "    , a.���J�e�S���[���� "
vSQL = vSQL & "    , a.���^�^�O "
vSQL = vSQL & "    , a.��J�e�S���[�R�[�h "
vSQL = vSQL & "    , b.��J�e�S���[�� "
vSQL = vSQL & "    , a.���ڏ��i���[�J�[�R�[�h AS ���ڏ��i���[�J�[�R�[�h1 "
vSQL = vSQL & "    , a.���ڏ��i���i�R�[�h AS ���ڏ��i���i�R�[�h1 "
vSQL = vSQL & "    , a.���ڏ��i�R�����g AS ���ڏ��i�R�����g1 "
vSQL = vSQL & "    , a.���ڏ��i���[�J�[�R�[�h2 "
vSQL = vSQL & "    , a.���ڏ��i���i�R�[�h2 "
vSQL = vSQL & "    , a.���ڏ��i�R�����g2 "
vSQL = vSQL & "    , a.���ڏ��i���[�J�[�R�[�h3 "
vSQL = vSQL & "    , a.���ڏ��i���i�R�[�h3 "
vSQL = vSQL & "    , a.���ڏ��i�R�����g3 "
vSQL = vSQL & "    , a.���ڏ��i���[�J�[�R�[�h4 "
vSQL = vSQL & "    , a.���ڏ��i���i�R�[�h4 "
vSQL = vSQL & "    , a.���ڏ��i�R�����g4 "
vSQL = vSQL & "    , (SELECT ���[�J�[�� FROM ���[�J�[ WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h) AS ���ڏ��i���[�J�[��1 "
vSQL = vSQL & "    , (SELECT ���i�� FROM Web���i WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h "
vSQL = vSQL & "                          AND ���i�R�[�h     = a.���ڏ��i���i�R�[�h) AS ���ڏ��i���i��1 "
vSQL = vSQL & "    , (SELECT ���[�J�[�� FROM ���[�J�[ WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h2) AS ���ڏ��i���[�J�[��2 "
vSQL = vSQL & "    , (SELECT ���i�� FROM Web���i WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h2 "
vSQL = vSQL & "                          AND ���i�R�[�h     = a.���ڏ��i���i�R�[�h2) AS ���ڏ��i���i��2 "
vSQL = vSQL & "    , (SELECT ���[�J�[�� FROM ���[�J�[ WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h3) AS ���ڏ��i���[�J�[��3 "
vSQL = vSQL & "    , (SELECT ���i�� FROM Web���i WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h3 "
vSQL = vSQL & "                          AND ���i�R�[�h     = a.���ڏ��i���i�R�[�h3) AS ���ڏ��i���i��3 "
vSQL = vSQL & "    , (SELECT ���[�J�[�� FROM ���[�J�[ WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h4) AS ���ڏ��i���[�J�[��4 "
vSQL = vSQL & "    , (SELECT ���i�� FROM Web���i WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h4 "
vSQL = vSQL & "                          AND ���i�R�[�h     = a.���ڏ��i���i�R�[�h4) AS ���ڏ��i���i��4 "
vSQL = vSQL & "    , (SELECT ���i�摜�t�@�C����_�� FROM Web���i WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h "
vSQL = vSQL & "                          AND ���i�R�[�h     = a.���ڏ��i���i�R�[�h) AS ���i�摜�t�@�C����_��1 "
vSQL = vSQL & "    , (SELECT ���i�摜�t�@�C����_�� FROM Web���i WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h2 "
vSQL = vSQL & "                          AND ���i�R�[�h     = a.���ڏ��i���i�R�[�h2) AS ���i�摜�t�@�C����_��2 "
vSQL = vSQL & "    , (SELECT ���i�摜�t�@�C����_�� FROM Web���i WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h3 "
vSQL = vSQL & "                          AND ���i�R�[�h     = a.���ڏ��i���i�R�[�h3) AS ���i�摜�t�@�C����_��3 "
vSQL = vSQL & "    , (SELECT ���i�摜�t�@�C����_�� FROM Web���i WITH (NOLOCK) "
vSQL = vSQL & "                        WHERE ���[�J�[�R�[�h = a.���ڏ��i���[�J�[�R�[�h4 "
vSQL = vSQL & "                          AND ���i�R�[�h     = a.���ڏ��i���i�R�[�h4) AS ���i�摜�t�@�C����_��4 "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    ���J�e�S���[ AS a WITH (NOLOCK) "
vSQL = vSQL & "  , ��J�e�S���[ AS b WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.���J�e�S���[�R�[�h = '" & MidCategoryCd & "'"
vSQL = vSQL & "    AND a.��J�e�S���[�R�[�h = b.��J�e�S���[�R�[�h"

'@@@@response.write(vSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open vSQL, Connection, adOpenStatic

If RS.EOF = True Then
	wNoData = "Y"
Else
	' ��J�e�S���R�[�h
	wLargeCategoryCd = RS("��J�e�S���[�R�[�h")
	' ��J�e�S����
	wLargeCategoryName = RS("��J�e�S���[��")
	' ���J�e�S����
	wMidCategoryName = RS("���J�e�S���[�����{��")
	' ���J�e�S���T�v
	wMidCategoryOverview = RS("���J�e�S���[����")

	wInsertHTMLPath1 = fGetInsertHTMLPath(MidCategoryCd,"1")
	wInsertHTMLPath2 = fGetInsertHTMLPath(MidCategoryCd,"2")

	' ���^�^�O <����n�܂��Ă��Ȃ��ꍇ�͖���
	If Left(RS("���^�^�O"), 1) = "<" Then
		wMetaTag = RS("���^�^�O")
	End If

	'----- HTML�쐬
	Call CreateChumokuHTML()				' ���ڏ��i
	
	If fExistMidCategoryStaticMainHTMLFile(MidCategoryCd) = False Then
		' �ꉟ�����i�p �ÓIHTML�e�L�X�g�t�@�C���쐬
		If fMakeMidCategoryStaticMainHTMLFile(MidCategoryCd, vFilePath, vMsg) = False Then
			Exit Function
		End If
	End If

	Call fIncludeMidCategoryStaticTextMain(MidCategoryCd)
	Call CreateSaleAndOutletHTML()

End If

RS.Close
Set RS = Nothing

End Function

'========================================================================
'
'	Function	���ڏ��i
'
'========================================================================
Function CreateChumokuHTML()

Dim vItem
Dim i
Dim vCnt

'----- ���ڏ��iHTML�ҏW
wHTML = ""
wHTML = wHTML & "        <h2 class='subtitle04' id='pickup'>" & wMidCategoryName & "�̃s�b�N�A�b�v�A�C�e��" & "</h2>" & vbNewLine

wHTML = wHTML & "        <div id='pickup_box'>" & vbNewLine

vCnt=1
For i = 1 To 4 Step 1
	If GetProductFlag(RS("���ڏ��i���[�J�[�R�[�h" & i),RS("���ڏ��i���i�R�[�h" & i)) = "Y" Then
		vItem = Server.URLEncode(RS("���ڏ��i���[�J�[�R�[�h" & i) & "^" & RS("���ڏ��i���i�R�[�h" & i))
		
		If (vCnt Mod 2) = 1 Then
			If vCnt <> 1 Then
				wHTML = wHTML & "            </ul>" & vbNewLine
			End If
			wHTML = wHTML & "            <ul class='pickup col" & int(vCnt/2)+(vCnt Mod 2)  & "' >" & vbNewLine
		End If
		wHTML = wHTML & "                <li>" & vbNewLine
		wHTML = wHTML & "                    <div class='pickup_inner'><div class='pickup_inner02'>" & vbNewLine
		wHTML = wHTML & "                        <div class='item_name_box'>" & vbNewLine
		wHTML = wHTML & "                            <p class='left'><a href='ProductDetail.asp?Item=" & vItem & "'>"
		If RS("���i�摜�t�@�C����_��" & i) <> "" Then
			wHTML = wHTML & "<img src='prod_img/" & RS("���i�摜�t�@�C����_��" & i) & "' alt='" & RS("���ڏ��i���[�J�[��" & i) & " / " & RS("���ڏ��i���i��" & i) & "' class='opover'>"
		End If
		wHTML = wHTML & "</a></p>" & vbNewLine
		wHTML = wHTML & "                            <p class='item_name'><a href='ProductDetail.asp?Item=" & vItem & "'>" & RS("���ڏ��i���[�J�[��" & i) & " / " & RS("���ڏ��i���i��" & i) & "</a><p>" & vbNewLine
		wHTML = wHTML & "                            <p>" & GetPrice(RS("���ڏ��i���[�J�[�R�[�h" & i), RS("���ڏ��i���i�R�[�h" & i)) & "</p>" & vbNewLine
		wHTML = wHTML & "                        </div>" & vbNewLine
		wHTML = wHTML & "                        <p class='desc'>" & RS("���ڏ��i�R�����g" & i) & "</p>" & vbNewLine
		wHTML = wHTML & "                    </div></div>" & vbNewLine
		wHTML = wHTML & "                </li>" & vbNewLine
		
		vCnt = vCnt+1
	End If
Next

If vCnt = 1 Then
	Exit Function
End If

wHTML = wHTML & "            </ul>" & vbNewLine
wHTML = wHTML & "        </div>" & vbNewLine

wChumokuHTML = wHTML

End Function

'========================================================================
'
'	Function	SALE&OUTLET���i
'	2012/08/21 ok Add
'========================================================================
Function CreateSaleAndOutletHTML()

Dim RSv
Dim vSQL
Dim v_price
Dim v_exprice
' 2012/09/03 GV #1426 Add Start
Dim wHTML1
Dim cnt
Dim ctr
Dim dcnt
Dim flg
Dim w_MakerCd()
Dim w_ItemCd()
Dim w_price1()
Dim w_price2()
cnt = 0
dcnt = 0
' 2012/09/03 GV #1426 Add End

'---- �Z�[�����i���o��
vSQL = ""
vSQL = vSQL & "SELECT "
' 2012/09/03 GV #1426 Mod Start
'vSQL = vSQL & "    TOP 5 "
vSQL = vSQL & "    TOP 20 "
' 2012/09/03 GV #1426 Mod End
vSQL = vSQL & "      a.���i�R�[�h "
vSQL = vSQL & "    , a.���i�� "
vSQL = vSQL & "    , a.���[�J�[�R�[�h "
vSQL = vSQL & "    , a.���[�J�[�� "
vSQL = vSQL & "    , a.���i�摜�t�@�C����_�� "
vSQL = vSQL & "    , a.�̔��P�� "
vSQL = vSQL & "    , a.�O��̔��P�� "
vSQL = vSQL & "    , a.ASK���i�t���O "
vSQL = vSQL & "    , a.B�i�t���O "
vSQL = vSQL & "    , a.�����萔�� "
vSQL = vSQL & "    , a.������P�� "
vSQL = vSQL & "    , a.������󒍍ϐ��� "
vSQL = vSQL & "    , a.�O��P���ύX�� "
vSQL = vSQL & "    , a.B�i�t���O "
vSQL = vSQL & "    , a.B�i�P�� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web�Z�[�����i a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "    a.�Z�[���敪�ԍ� BETWEEN 1 AND 4"
vSQL = vSQL & " AND a.���J�e�S���[�R�[�h = '" & MidCategoryCd & "' "
vSQL = vSQL & "ORDER BY NEWID() "

'@@@@@@@@@@Response.Write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic
wHTML = ""

If RSv.EOF = false Then
	'----- �Z�[�����iHTML�ҏW
	wHTML = wHTML & "        <h2 class='subtitle_red'>" & wMidCategoryName & "��SALE &amp; OUTLET</h2>" & vbNewLine
	wHTML = wHTML & "        <div class='box'><div class='box_inner01'>" & vbNewLine
	wHTML = wHTML & "            <ul class='list'>" & vbNewLine

	Do Until RSv.EOF = True OR dcnt > 4
' 2012/09/03 GV #1426 Add Start
		ReDim Preserve w_MakerCd(cnt)
		w_MakerCd(cnt) = RSv("���[�J�[�R�[�h")
		ReDim Preserve w_ItemCd(cnt)
		w_ItemCd(cnt) = RSv("���i�R�[�h")
		wHTML1 = ""
' 2012/09/03 GV #1426 Add End
		wHTML1 = wHTML1 & "                <li><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h")) & "'>"
		If RSv("���i�摜�t�@�C����_��") <> "" Then
			wHTML1 = wHTML1 & "<img src='prod_img/" & RSv("���i�摜�t�@�C����_��") & "' alt='" & RSv("���[�J�[��") & " / " & RSv("���i��") & "' class='opover'>"
		End If
		wHTML1 = wHTML1 & RSv("���[�J�[��") & " / " & RSv("���i��") & "</a><span>"
		
		'---- �̔��P��
		v_price = calcPrice(RSv("�̔��P��"), wSalesTaxRate)
		v_exprice = calcPrice(RSv("�O��̔��P��"), wSalesTaxRate)
		'1�s�ڂ̕\���iASK���i�ł͂Ȃ��l�����i�̋����i�j
		If RSv("ASK���i�t���O") <> "Y" Then
			If RSv("B�i�t���O") = "Y" OR (RSv("�����萔��") > RSv("������󒍍ϐ���") AND RSv("�����萔��") > 0) OR ( isNULL(RSv("�O��P���ύX��")) = False AND DateAdd("d", 60, RSv("�O��P���ύX��")) >= Date() AND RSv("�O��̔��P��") > RSv("�̔��P��") AND RSv("�O��̔��P��") <> 0) Then

				'�l�����i�̋����i��\��
				If isNULL(RSv("�O��P���ύX��")) = False AND DateAdd("d", 60, RSv("�O��P���ύX��")) >= Date() AND RSv("�O��̔��P��") > RSv("�̔��P��") Then
'2013/03/19 GV mod start ---->
'�O��P���͂��΂炭�\�������Ȃ�
'					wHTML1 = wHTML1 & FormatNumber(v_exprice,0) & "�~�i�ō��j��<br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
					ReDim Preserve w_price1(cnt)
					w_price1(cnt) = FormatNumber(v_exprice,0)
' 2012/09/03 GV #1426 Add End
				'B�i�A����i�͔̔����i�������i�Ƃ��ĕ\��
				Else
'2013/03/19 GV mod start ---->
'�O��P���͂��΂炭�\�������Ȃ�
'					wHTML1 = wHTML1 & FormatNumber(v_price,0) & "�~�i�ō��j��<br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
					ReDim Preserve w_price1(cnt)
					w_price1(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
				End If
' 2012/09/03 GV #1426 Add Start
			Else
				ReDim Preserve w_price1(cnt)
				w_price1(cnt) = 0
' 2012/09/03 GV #1426 Add End
			End If
' 2012/09/03 GV #1426 Add Start
		Else
			ReDim Preserve w_price1(cnt)
			w_price1(cnt) = 0
' 2012/09/03 GV #1426 Add End
		End If

		'2�s�ڂ̕\���i�ʏ퉿�i or ASK or �l�����㉿�i�j
		If RSv("ASK���i�t���O") <> "Y" Then
			'---- B�i�P��
			If RSv("B�i�t���O") = "Y" Then
				v_price = calcPrice(RSv("B�i�P��"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "<strong>�y�킯����i�����z" & FormatNumber(v_price,0) & "�~(�ō�)</strong>"
				wHTML1 = wHTML1 & "<strong>�y�킯����i�����z" & FormatNumber(RSv("B�i�P��"),0) & "�~(�Ŕ�)</strong><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			'---- ������P��
			ElseIf RSv("�����萔��") > RSv("������󒍍ϐ���") AND RSv("�����萔��") > 0 Then
				v_price = calcPrice(RSv("������P��"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "<strong>�y��������z" & FormatNumber(v_price,0) & "�~(�ō�)</strong>"
				wHTML1 = wHTML1 & "<strong>�y��������z" & FormatNumber(RSv("������P��"),0) & "�~(�Ŕ�)</strong><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			'---- �ʏ폤�i
			Else
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "<strong>�y�Ռ������z" & FormatNumber(v_price,0) & "�~(�ō�)</strong>"
				wHTML1 = wHTML1 & "<strong>�y�Ռ������z" & FormatNumber(RSv("�̔��P��"),0) & "�~(�Ŕ�)</strong><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			End If

			wHTML1 = wHTML1 & "(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)"	'2013/03/19 GV ad

			wHTML1 = wHTML1 & "</span></li>" & vbNewLine

		Else
			'---- B�i�P��
			If RSv("B�i�t���O") = "Y" Then
				v_price = calcPrice(RSv("B�i�P��"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "�y�킯����i�����z</span><a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "�~(�ō�)</span>"
				wHTML1 = wHTML1 & "�y�킯����i�����z</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("B�i�P��"),0) & "�~(�Ŕ�)</span><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			'---- ������P��
			ElseIf RSv("�����萔��") > RSv("������󒍍ϐ���") AND RSv("�����萔��") > 0 Then
				v_price = calcPrice(RSv("������P��"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "�y��������z</span><a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "�~(�ō�)</span>"
				wHTML1 = wHTML1 & "�y��������z</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("������P��"),0) & "�~(�Ŕ�)</span><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			'---- �ʏ폤�i
			Else
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "�y�Ռ������z</span><a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "�~(�ō�)</span>"
				wHTML1 = wHTML1 & "�y�Ռ������z</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("�̔��P��"),0) & "�~(�Ŕ�)</span><br>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			End If

			wHTML1 = wHTML1 & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)</span>"	'2013/03/19 GV ad

			wHTML1 = wHTML1 & "</a></li>" & vbNewLine

		End If

' 2012/09/03 GV #1426 Add Start
		flg = True
		For ctr = 0 to Ubound(w_ItemCd)
			If ctr < cnt Then
				If w_MakerCd(ctr) = w_MakerCd(cnt) AND w_ItemCd(ctr) = w_ItemCd(cnt) Then
					if w_price1(ctr) = w_price1(cnt) AND w_price2(ctr) = w_price2(cnt) Then
						flg = False
						Exit For
					End If
				End If
			End If
		Next
		if flg Then
			dcnt = dcnt + 1
			wHTML = wHTML & wHTML1
		End If
		cnt = cnt + 1
' 2012/09/03 GV #1426 Add End

		RSv.MoveNext
	Loop

	wHTML = wHTML & "            </ul>" & vbNewLine
	wHTML = wHTML & "        </div></div>" & vbNewLine
End If
wSaleAndOutletHTML = wHTML

RSv.Close

End Function

'========================================================================
'
'	Function	���i���i�擾
'
'========================================================================
Function GetPrice(pMakerCd, pProductCd)

Dim RSv
Dim vSQL
Dim v_price
GetPrice = ""

'---- Web���i�t���O���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.�̔��P�� "
vSQL = vSQL & "    , a.�O��̔��P�� "
vSQL = vSQL & "    , a.ASK���i�t���O "
vSQL = vSQL & "    , a.B�i�t���O "
vSQL = vSQL & "    , a.�����萔�� "
vSQL = vSQL & "    , a.������P�� "
vSQL = vSQL & "    , a.������󒍍ϐ��� "
vSQL = vSQL & "    , a.�O��P���ύX�� "
vSQL = vSQL & "    , a.B�i�t���O "
vSQL = vSQL & "    , a.B�i�P�� "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web���i a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.���[�J�[�R�[�h = '" & pMakerCd & "' "
vSQL = vSQL & "    AND a.���i�R�[�h     = '" & pProductCd & "'"
vSQL = vSQL & "    AND a.Web���i�t���O  = 'Y'"

'@@@@@@@@@@Response.Write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = false Then
	'---- �̔��P��
	v_price = calcPrice(RSv("�̔��P��"), wSalesTaxRate)

	If RSv("ASK���i�t���O") <> "Y" Then
		'---- B�i�P��
		If RSv("B�i�t���O") = "Y" Then
			v_price = calcPrice(RSv("B�i�P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice & "�y�킯����i�����z" & FormatNumber(v_price,0) & "�~(�ō�)"
			GetPrice = GetPrice & "�y�킯����i�����z" & FormatNumber(RSv("B�i�P��"),0) & "�~(�Ŕ�)<br>"
			GetPrice = GetPrice & "(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)"
'2014/03/19 GV mod end   <----
		'---- ������P��
		ElseIf RSv("�����萔��") > RSv("������󒍍ϐ���") AND RSv("�����萔��") > 0 Then
			v_price = calcPrice(RSv("������P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice& "�y��������z" & FormatNumber(v_price,0) & "�~(�ō�)"
			GetPrice = GetPrice& "�y��������z" & FormatNumber(RSv("������P��"),0) & "�~(�Ŕ�)<br>"
			GetPrice = GetPrice & "(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)"
'2014/03/19 GV mod end   <----
		'---- �ʏ폤�i
		Else
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice &  FormatNumber(v_price,0) & "�~(�ō�)"
			GetPrice = GetPrice &  FormatNumber(RSv("�̔��P��"),0) & "�~(�Ŕ�)<br>"
			GetPrice = GetPrice & "(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)"
'2014/03/19 GV mod end   <----
		End If
	Else
		'---- B�i�P��
		If RSv("B�i�t���O") = "Y" Then
			v_price = calcPrice(RSv("B�i�P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice & "�y�킯����i�����z<a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "�~(�ō�)</span></a>"
			GetPrice = GetPrice & "�y�킯����i�����z<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("B�i�P��"),0) & "�~(�Ŕ�)</span><br>"
			GetPrice = GetPrice & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)</span></a>"
'2014/03/19 GV mod end   <----
		'---- ������P��
		ElseIf RSv("�����萔��") > RSv("������󒍍ϐ���") AND RSv("�����萔��") > 0 Then
			v_price = calcPrice(RSv("������P��"), wSalesTaxRate)
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice & "�y��������z<a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "�~(�ō�)</span></a>"
			GetPrice = GetPrice & "�y��������z<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("������P��"),0) & "�~(�Ŕ�)</span><br>"
			GetPrice = GetPrice & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)</span></a>"
'2014/03/19 GV mod end   <----
		'---- �ʏ폤�i
		Else
'2014/03/19 GV mod start ---->
'			GetPrice = GetPrice & "<a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "�~(�ō�)</span></a>"
			GetPrice = GetPrice & "<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("�̔��P��"),0) & "�~(�Ŕ�)</span><br>"
			GetPrice = GetPrice & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)</span></a>"
'2014/03/19 GV mod end   <----
		End If
	End If
End If

RSv.Close

End Function

'========================================================================
'
'	Function	Web���i�t���O�`�F�b�N
'
'========================================================================
Function GetProductFlag(pMakerCd, pProductCd)

Dim RSv
Dim vSQL
GetProductFlag = ""

'---- Web���i�t���O���o��
vSQL = ""
vSQL = vSQL & "SELECT "
vSQL = vSQL & "      a.Web���i�t���O "
vSQL = vSQL & "FROM "
vSQL = vSQL & "    Web���i a WITH (NOLOCK) "
vSQL = vSQL & "WHERE "
vSQL = vSQL & "        a.���[�J�[�R�[�h = '" & pMakerCd & "' "
vSQL = vSQL & "    AND a.���i�R�[�h     = '" & pProductCd & "'"

'@@@@@@@@@@Response.Write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = false Then
	GetProductFlag = RSv("Web���i�t���O")
End If

RSv.Close

End Function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<meta name="robots" content="noindex,nofollow">
<title><% = wMidCategoryName %> �ꗗ�b�T�E���h�n�E�X</title>
<% = wMetaTag %>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css?20121116" type="text/css">
<link rel="stylesheet" href="style/categorylist.css?20140812" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
	<span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
<!-- �R���e���cstart -->
	<div id="globalContents">
		<div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
			<p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
			<ul id="path">
				<li><span itemscope itemtype="http://data-vocabulary.org/Breadcrumb"><a href="LargeCategoryList.asp?LargeCategoryCd=<%=wLargeCategoryCd%>" itemprop="url"><span itemprop="title"><%=wLargeCategoryName%></span></a></span></li>
				<li class="now"><span itemscope itemtype="http://data-vocabulary.org/Breadcrumb"><span itemprop='title'><%=wMidCategoryName%></span></span></li>
			</ul>
		</div></div></div>
<!-- �y�[�W���C�������̋L�q START -->

<%=fIncludeInsertHTML(wInsertHTMLPath1)%>

<!-- ���J�e�S���[�ɂ��āE�J�e�S���[����I�ԁE�ŐV�j���[�X�E�V���i -->

<%=wStaticHTML(0)%>

<%=wChumokuHTML%>

<%=wStaticHTML(1)%>

<%=fIncludeInsertHTML(wInsertHTMLPath2)%>

<%=wSaleAndOutletHTML%>

<%=wStaticHTML(2)%>

    <!--/#contents --></div>
  
<!-- �i�������pForm -->
    <form name='f_search' method='get' action='SearchList.asp'>
      <input type='hidden' name='s_maker_cd' value=''>
      <input type='hidden' name='s_category_cd' value=''>
      <input type='hidden' name='s_mid_category_cd' value='<% = MidCategoryCd %>'>
      <input type='hidden' name='s_large_category_cd' value=''>
      <input type='hidden' name='s_product_cd' value=''>
      <input type='hidden' name='search_all' value=''>
      <input type='hidden' name='sSeriesCd' value=''>
      <input type='hidden' name='sPriceFrom' value=''>
      <input type='hidden' name='sPriceTo' value=''>
      <input type='hidden' name='i_type' value=''>
      <input type='hidden' name='i_sub_type' value=''>
      <input type='hidden' name='i_page' value='1'>
      <input type='hidden' name='i_sort' value=''>
      <input type='hidden' name='i_page_size' value=''>
      <input type='hidden' name='i_ListType' value=''>
    </form>

	<div id="globalSide">
<%
	' ��NAVI�p�p�����[�^�Z�b�g
	NAVIMidCategoryCd = MidCategoryCd
	NAVISearchListMakerListHTML = wNaviMakerHTML
	NAVISearchListCategoryListHTML = wNaviCategoryHTML
	NAVISearchListPriceRangeListHTML = wNaviPriceRangeHTML
%>
<!--#include file="../Navi/NaviSideShop.inc"-->
<!--#include file="../Navi/NaviSide.inc"-->
    <!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/MidCategoryList.js"></script>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript" src="jslib/SearchList.js?20121108" charset="Shift_JIS"></script>
<script type="text/javascript" src="../jslib/jquery.tinyscrollbar.min.js"></script>
<script type="text/javascript">
$(function(){
    $('#scrollbar1').tinyscrollbar();
});
</script>
</body>
</html>
