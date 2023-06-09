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
<!--#include file="./LargeCategoryList/LargeCategoryList.inc"-->
<%
'========================================================================
'
'	��J�e�S���[�ꗗ�y�[�W
'
'�X�V����
'2008/12/19 ���j���[�A���@�V�K
'2009/04/13 ���ALL�\���ɕύX
'2009/05/20 MidCategoryList�ւ̃����N�ǉ�
'2009/08/18 �g�s�b�N�X(News)�̕\�������ɏ��i�L��.��J�e�S���[�R�[�h=�Y����J�e�S���[�R�[�h��ǉ�
'2009/11/05 an META�^�O�ǉ��@�\��ǉ�
'2010/01/26 an ���݂��Ȃ��J�e�S���[���w�肵���ꍇ�́AError.asp��\���iError.asp����TOP�Ƀ��_�C���N�g�j
'2010/05/17 ko-web �����΍�̂���HTML�^�O�ihX,p,strong�j�ǉ�
'2010/12/07 an ��ʋL����\������悤�ɏC��
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2012/01/20 GV WITH (NOLOCK) �R�� �ǉ�
'2012/01/20 GV �ŐV�j���[�X ����� �V���i �̏����uWeb���i�L��TOP10�v�e�[�u������сuWeb�V���iTOP10�v�e�[�u�����擾����悤�ύX
'2012/01/23 GV ���i�L���̃\�[�g�� �L�����t DESC �� �L���ԍ� DESC �֕ύX
'2012/03/13 GV #1224 �u�ꉟ�����i�v�̕\�����ȊO���O���̐ÓI�e�L�X�g�t�@�C�������荞�ނ悤�ɕύX
'2012/03/13 GV #1224 �u�ꉟ�����i�v�̕\�����ȊO���O���̐ÓI�e�L�X�g�t�@�C�������݂������͗L�������؂�̏ꍇ�A�������鏈���ǉ�
'2012/07/23 ok ���j���[�A���ɔ����V�f�U�C���ɕύX
'2012/09/03 GV #1426 ��E���J�e�S����ʂŕ\�������SALES&OUTLET���̕\���f�[�^����ӂɎ擾�E�\������
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================
On Error Resume Next

Dim LargeCategoryCd
'Dim ALLFl										' 2012/03/13 GV Del

Dim wSalesTaxRate

Dim wLargeCategoryName
Dim wLargeCategoryComment
Dim wMetaTag
Dim wNoData '2010/01/26 an �ǉ�

Dim wIchioshiHTML
'Dim wMidCategoryListHTML						' 2012/03/13 GV Del
'Dim wNewsHTML									' 2012/03/13 GV Del
'Dim wNewItemHTML								' 2012/03/13 GV Del

Dim Connection
Dim RS

Dim wSQL

Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add
Dim wInsertHTMLPath1		'2012/07/24 ok Add
Dim wInsertHTMLPath2		'2012/07/24 ok Add
Dim wStaticHTML(2)			'2012/07/24 ok Add
Dim wSaleAndOutletHTML		'2012/07/24 ok Add

'========================================================================

Response.Buffer = True

'---- Get input data
LargeCategoryCd = ReplaceInput(Trim(Request("LargeCategoryCd")))
'AllFl = ReplaceInput(Trim(Request("AllFl")))							' 2012/03/13 GV Del (���g�p�̈�)

'AllFl = "Y"															' 2012/03/13 GV Del (���g�p�̈�)

'---- Execute main
Call connect_db()
Call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
If Err.Description <> "" Then
	wErrDesc = "LargeCategoryList.asp" & " " & Replace(Replace(Err.Description, vbCr, " "), vbLf, " ")
	Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
End If                                           '2011/08/01 an add e

Call close_db()

If wNoData = "Y" Or Err.Description <> "" Then '2010/01/26 an �C��
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
'	Function	Main
'
'========================================================================
Function main()

Dim vFilePath			' 2012/03/13 GV Add
Dim vMsg				' 2012/03/13 GV Add
Dim vItemChar1
Dim vItemChar2
Dim vItemNum1
Dim vItemNum2
Dim vItemDate1
Dim vItemDate2


'--- ����ŗ���o��
Call getCntlMst("����","����ŗ�","1", vItemChar1, vItemChar2, vItemNum1, vItemNum2, vItemDate1, vItemDate2)			'����ŗ�
wSalesTaxRate = Clng(vItemNum1)

'---- ��J�e�S���[ ���o��
wSQL = ""
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.��J�e�S���[�� "
wSQL = wSQL & "    , a.��J�e�S���[���� "
wSQL = wSQL & "    , a.�ꉟ�����[�J�[�R�[�h1 "
wSQL = wSQL & "    , a.�ꉟ�����i�R�[�h1 "
'wSQL = wSQL & "    , a.�ꉟ���摜�t�@�C����1 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.�ꉟ�����[�J�[�R�[�h2 "
wSQL = wSQL & "    , a.�ꉟ�����i�R�[�h2 "
'wSQL = wSQL & "    , a.�ꉟ���摜�t�@�C����2 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.�ꉟ�����[�J�[�R�[�h3 "
wSQL = wSQL & "    , a.�ꉟ�����i�R�[�h3 "
'wSQL = wSQL & "    , a.�ꉟ���摜�t�@�C����3 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.�ꉟ�����[�J�[�R�[�h4 "
wSQL = wSQL & "    , a.�ꉟ�����i�R�[�h4 "
'wSQL = wSQL & "    , a.�ꉟ���摜�t�@�C����4 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.�ꉟ�����[�J�[�R�[�h5 "
wSQL = wSQL & "    , a.�ꉟ�����i�R�[�h5 "
'wSQL = wSQL & "    , a.�ꉟ���摜�t�@�C����5 "		'2012/07/24 ok Del
wSQL = wSQL & "    , a.���^�^�O "
wSQL = wSQL & "    , (SELECT ���[�J�[�� FROM ���[�J�[ WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) �t��
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h1) AS �ꉟ�����[�J�[��1 "
wSQL = wSQL & "    , (SELECT ���i�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h1 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h1) AS �ꉟ�����i��1 "
wSQL = wSQL & "    , (SELECT ���[�J�[�� FROM ���[�J�[ WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) �t��
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h2) AS �ꉟ�����[�J�[��2 "
wSQL = wSQL & "    , (SELECT ���i�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h2 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h2) AS �ꉟ�����i��2 "
wSQL = wSQL & "    , (SELECT ���[�J�[�� FROM ���[�J�[ WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) �t��
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h3) AS �ꉟ�����[�J�[��3 "
wSQL = wSQL & "    , (SELECT ���i�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h3 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h3) AS �ꉟ�����i��3 "
wSQL = wSQL & "    , (SELECT ���[�J�[�� FROM ���[�J�[ WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) �t��
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h4) AS �ꉟ�����[�J�[��4 "
wSQL = wSQL & "    , (SELECT ���i�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h4 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h4) AS �ꉟ�����i��4 "
wSQL = wSQL & "    , (SELECT ���[�J�[�� FROM ���[�J�[ WITH (NOLOCK) "												' 2012/01/20 GV Mod  WITH (NOLOCK) �t��
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h5) AS �ꉟ�����[�J�[��5 "
wSQL = wSQL & "    , (SELECT ���i�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h5 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h5) AS �ꉟ�����i��5 "
'2012/07/25 ok Add Start
wSQL = wSQL & "    , (SELECT ���i�摜�t�@�C����_�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h1 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h1) AS ���i�摜�t�@�C����_��1 "
wSQL = wSQL & "    , (SELECT ���i�摜�t�@�C����_�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h2 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h2) AS ���i�摜�t�@�C����_��2 "
wSQL = wSQL & "    , (SELECT ���i�摜�t�@�C����_�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h3 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h3) AS ���i�摜�t�@�C����_��3 "
wSQL = wSQL & "    , (SELECT ���i�摜�t�@�C����_�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h4 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h4) AS ���i�摜�t�@�C����_��4 "
wSQL = wSQL & "    , (SELECT ���i�摜�t�@�C����_�� FROM Web���i WITH (NOLOCK) "
wSQL = wSQL & "                        WHERE ���[�J�[�R�[�h = a.�ꉟ�����[�J�[�R�[�h5 "
wSQL = wSQL & "                          AND ���i�R�[�h     = a.�ꉟ�����i�R�[�h5) AS ���i�摜�t�@�C����_��5 "
wSQL = wSQL & "    , �ꉟ���R�����g "
'2012/07/25 ok Add End
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ��J�e�S���[ a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "    a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "

'@@@@@@@@@@Response.Write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

If RS.EOF = True Then
	wNoData = "Y" '2010/01/26 an �C��
Else
	'----- ��J�e�S���[��
	wLargeCategoryName = RS("��J�e�S���[��")
	wLargeCategoryComment = RS("��J�e�S���[����")

'2012/07/24 ok Add Start
	wInsertHTMLPath1 = fGetInsertHTMLPath(LargeCategoryCd,"1")
	wInsertHTMLPath2 = fGetInsertHTMLPath(LargeCategoryCd,"2")
'2012/07/24 ok Add End

	'----- ���^�^�O <����n�܂��Ă��Ȃ��ꍇ�͖���
	If Left(RS("���^�^�O"),1) = "<" Then
		wMetaTag = RS("���^�^�O")
	End If

	'----- HTML�쐬
	Call CreateIchioshiHTML()				' �ꉟ�����i
'	Call CreateMidCategoryListHTML()		' ���J�e�S���[�ꗗ
'	Call CreatewNewsHTML()					' �g�s�b�N�X News
'	Call CreateNewItemHTML()				' �g�s�b�N�X �V���i

' 2012/03/13 GV Add Start
	' �ꉟ�����i�p �ÓIHTML�t�@�C���̑��݃`�F�b�N (�L�������؂�`�F�b�N)
	If fExistLargeCategoryStaticHTMLFile(LargeCategoryCd) = False Then

		' �ꉟ�����i�p �ÓIHTML�e�L�X�g�t�@�C���쐬
		If fMakeLargeCategoryStaticHTMLFile(LargeCategoryCd, vFilePath, vMsg) = False Then
			Exit Function
		End If

	End If
' 2012/03/13 GV Add End

'2012/07/24 ok Add Start
	Call fIncludeLargeCategoryStaticText(LargeCategoryCd)
	Call CreateSaleAndOutletHTML()
'2012/07/24 ok Add End
End If

RS.Close

End Function

'========================================================================
'
'	Function	�ꉟ�����i
'
'========================================================================
Function CreateIchioshiHTML()

'Dim vPrice		'2012/07/24 ok Del
Dim vItem
Dim i
Dim vCnt

'----- �ꉟ�����iHTML�ҏW
wHTML = ""
'2012/07/24 ok Del Start
'wHTML = wHTML & "<table width='794' border='0' cellpadding='0' cellspacing='0' id='Shop_LargeCategory_HotItem'>" & vbNewLine
'
'vPrice = getPrice(RS("�ꉟ�����[�J�[�R�[�h1"), RS("�ꉟ�����i�R�[�h1"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("�ꉟ�����[�J�[�R�[�h1") & "^" & RS("�ꉟ�����i�R�[�h1"))
'
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td rowspan='2' class='1oshi'>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("�ꉟ���摜�t�@�C����1") & "' alt='" & RS("�ꉟ�����[�J�[��1") & " " & RS("�ꉟ�����i��1") & "' width='406' height='320' border='0'></a><br>�Ռ�����&nbsp;" & FormatNumber(vPrice,0) & "�~�i�ō��j&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'
'vPrice = getPrice(RS("�ꉟ�����[�J�[�R�[�h2"), RS("�ꉟ�����i�R�[�h2"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("�ꉟ�����[�J�[�R�[�h2") & "^" & RS("�ꉟ�����i�R�[�h2"))
'
'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("�ꉟ���摜�t�@�C����2") & "' alt='" & RS("�ꉟ�����[�J�[��2") & " " & RS("�ꉟ�����i��2") & "' width='190' height='150' border='0'></a><br>" & FormatNumber(vPrice,0) & "�~�i�ō��j&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'
'vPrice = getPrice(RS("�ꉟ�����[�J�[�R�[�h3"), RS("�ꉟ�����i�R�[�h3"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("�ꉟ�����[�J�[�R�[�h3") & "^" & RS("�ꉟ�����i�R�[�h3"))
'
'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("�ꉟ���摜�t�@�C����3") & "' alt='" & RS("�ꉟ�����[�J�[��3") & " " & RS("�ꉟ�����i��3") & "' width='190' height='150' border='0'></a><br>" & FormatNumber(vPrice,0) & "�~�i�ō��j&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'
'vPrice = getPrice(RS("�ꉟ�����[�J�[�R�[�h4"), RS("�ꉟ�����i�R�[�h4"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("�ꉟ�����[�J�[�R�[�h4") & "^" & RS("�ꉟ�����i�R�[�h4"))
'
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("�ꉟ���摜�t�@�C����4") & "' alt='" & RS("�ꉟ�����[�J�[��4") & " " & RS("�ꉟ�����i��4") & "' width='190' height='150' border='0'></a><br>" & FormatNumber(vPrice,0) & "�~�i�ō��j&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'
'vPrice = getPrice(RS("�ꉟ�����[�J�[�R�[�h5"), RS("�ꉟ�����i�R�[�h5"))
'vPrice = calcPrice(vPrice, wSalesTaxRate)
'vItem = Server.URLEncode(RS("�ꉟ�����[�J�[�R�[�h5") & "^" & RS("�ꉟ�����i�R�[�h5"))
'
'wHTML = wHTML & "    <td>" & vbNewLine
'wHTML = wHTML & "      <a href='ProductDetail.asp?Item=" & vItem & "'><img src='cat_hotitem/" & RS("�ꉟ���摜�t�@�C����5") & "' alt='" & RS("�ꉟ�����[�J�[��5") & " " & RS("�ꉟ�����i��5") & "' width='190' height='150' border='0'></a><br>" & FormatNumber(vPrice,0) & "�~�i�ō��j&nbsp;" & vbNewLine
'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'
'wHTML = wHTML & "</table>"
'2012/07/24 ok Del End

'2012/07/24 ok Add Start
wHTML = wHTML & "  <h2 class='subtitle pickup'>" & wLargeCategoryName & "�̃C�`�I�V���i"
If RS("�ꉟ���R�����g") <> "" Then
	wHTML = wHTML & "<span>�m" & RS("�ꉟ���R�����g") & "�n</span>"
End If
wHTML = wHTML & "</h2>" & vbNewLine
wHTML = wHTML & "  <ul class='rank'>" & vbNewLine

vCnt = 0
For i = 1 To 5 Step 1
	If GetProductFlag(RS("�ꉟ�����[�J�[�R�[�h" & i),RS("�ꉟ�����i�R�[�h" & i)) = "Y" Then
		vItem = Server.URLEncode(RS("�ꉟ�����[�J�[�R�[�h" & i) & "^" & RS("�ꉟ�����i�R�[�h" & i))
		wHTML = wHTML & "    <li class='rank0" & i-vCnt & "' ><a href='ProductDetail.asp?Item=" & vItem & "'>"
		If RS("���i�摜�t�@�C����_��" & i) <> "" Then
			wHTML = wHTML & "<img src='prod_img/" & RS("���i�摜�t�@�C����_��" & i) & "' alt='" & RS("�ꉟ�����[�J�[��" & i) & " / " & RS("�ꉟ�����i��" & i) & "' class='opover'>"
		End If
		wHTML = wHTML & RS("�ꉟ�����[�J�[��" & i) & " / " & RS("�ꉟ�����i��" & i) & "</a></li>" & vbNewLine
	Else
		vCnt = vCnt + 1
	End If
Next

wHTML = wHTML & "  </ul>" & vbNewLine
'2012/07/24 ok Add End

wIchioshiHTML = wHTML

End Function

'========================================================================
'
'	Function	�P�����o��
'
'========================================================================
Function GetPrice(pMakerCd, pProductCd)

Dim RSv

GetPrice = 0

'---- �P�����o��
wSQL = ""
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.�̔��P�� "
wSQL = wSQL & "    , a.B�i�P�� "
wSQL = wSQL & "    , a.������P�� "
wSQL = wSQL & "    , a.B�i�t���O "
wSQL = wSQL & "    , a.�����萔�� "
wSQL = wSQL & "    , a.������󒍍ϐ��� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web���i a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.���[�J�[�R�[�h = '" & pMakerCd & "' "
wSQL = wSQL & "    AND a.���i�R�[�h     = '" & pProductCd & "'"

'@@@@@@@@@@Response.Write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	Exit Function
End If

If RSv("B�i�t���O") = "Y" Then

	'---- B�i����
	GetPrice = RSv("B�i�P��")

Else

	If RSv("�����萔��") > RSv("������󒍍ϐ���") And RSv("�����萔��") > 0 Then
		'---- ������P��
		GetPrice = RSv("������P��")
	Else
		'---- �̔��P��
		GetPrice = RSv("�̔��P��")
	End If

End If

End Function

' 2012/03/13 GV Del Start
''========================================================================
''
''	Function	���J�e�S���[�ꗗ
''
''========================================================================
'Function CreateMidCategoryListHTML()
'
'Dim RSv
'
''---- ���J�e�S���[�A�J�e�S���[ ���o��
'wSQL = ""
'wSQL = wSQL & "SELECT a.���J�e�S���[�R�[�h"
'wSQL = wSQL & "     , a.���J�e�S���[�����{��"
'wSQL = wSQL & "     , ISNULL(a.���J�e�S���[�摜�t�@�C����,'') AS ���J�e�S���[�摜�t�@�C����"
'wSQL = wSQL & "  FROM ���J�e�S���[ a WITH (NOLOCK)"
'wSQL = wSQL & " WHERE a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "'"
'wSQL = wSQL & " ORDER BY a.�\����"
'
''@@@@@@@@@@Response.Write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'If RSv.EOF = True Then
'	Exit Function
'End If
'
'wHTML = ""
'wHTML = wHTML & "<table width='794' border='0' cellspacing='4' cellpadding='0' id='Shop_LargeCategory_MidCat'>" & vbNewLine
'wHTML = wHTML & "<tr>" & vbNewLine
'
'Do Until RSv.EOF = True
'
'	'---- ���@�ҏW
'	wHTML = wHTML & "    <td>" & vbNewLine
'	wHTML = wHTML & "      <table width='253' height:'100%' border='0' cellspacing='0' cellpadding='0' id='Shop_LargeCategory_SmallCat'>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td align='center' class='cat_left'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("���J�e�S���[�R�[�h") & "'><img src='cat_img/" & RSv("���J�e�S���[�摜�t�@�C����") & "' width='50' height='50' border='0' alt='" & RSv("���J�e�S���[�����{��") & "'></a></td>" & vbNewLine
'	wHTML = wHTML & "          <td class='cat_right'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("���J�e�S���[�R�[�h") & "'><h3>" & RSv("���J�e�S���[�����{��") & "</h3></a><br>" & vbNewLine
'
'	wHTML = wHTML & SetCategory(RSv("���J�e�S���[�R�[�h"))
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'
'	RSv.MoveNext
'	If RSv.EOF = True Then
' 		wHTML = wHTML & " </tr>" & vbNewLine
'		Exit Do
'	End If
'
'	'---- ���@�ҏW
'	wHTML = wHTML & "    <td>" & vbNewLine
'	wHTML = wHTML & "      <table width='253' height:'100%' border='0' cellspacing='0' cellpadding='0' id='Shop_LargeCategory_SmallCat'>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td align='center' class='cat_left'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("���J�e�S���[�R�[�h") & "'><img src='cat_img/" & RSv("���J�e�S���[�摜�t�@�C����") & "' width='50' height='50' border='0' alt='" & RSv("���J�e�S���[�����{��") & "'></a></td>" & vbNewLine
'	wHTML = wHTML & "          <td class='cat_right'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("���J�e�S���[�R�[�h") & "' class='link'><h3>" & RSv("���J�e�S���[�����{��") & "</h3></a><br>" & vbNewLine
'
'	wHTML = wHTML & SetCategory(RSv("���J�e�S���[�R�[�h"))
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'
'	RSv.MoveNext
'	If RSv.EOF = True Then
' 		wHTML = wHTML & " </tr>" & vbNewLine
'		Exit Do
'	End If
'
'	'---- �E�@�ҏW
'	wHTML = wHTML & "    <td>" & vbNewLine
'	wHTML = wHTML & "      <table width='253' height:'100%' border='0' cellspacing='0' cellpadding='0' id='Shop_LargeCategory_SmallCat'>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td align='center' class='cat_left'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("���J�e�S���[�R�[�h") & "'><img src='cat_img/" & RSv("���J�e�S���[�摜�t�@�C����") & "' width='50' height='50' border='0' alt='" & RSv("���J�e�S���[�����{��") & "'></a></td>" & vbNewLine
'	wHTML = wHTML & "          <td class='cat_right'><a href='MidCategoryList.asp?MidCategoryCd=" & RSv("���J�e�S���[�R�[�h") & "' class='link'><h3>" & RSv("���J�e�S���[�����{��") & "</h3></a><br>" & vbNewLine
'
'	wHTML = wHTML & SetCategory(RSv("���J�e�S���[�R�[�h"))
'
'	wHTML = wHTML & "          </td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'
'	RSv.MoveNext
'	If RSv.EOF = True Then
' 		wHTML = wHTML & " </tr>" & vbNewLine
'		Exit Do
'	End If
'
' 	wHTML = wHTML & " </tr>" & vbNewLine
'Loop
'
'wHTML = wHTML & "</table>" & vbNewLine
'wMidCategoryListHTML = wHTML
'
'RSv.Close
'
'End Function
'
''========================================================================
''
''	Function	�J�e�S���[�ꗗ
''
''========================================================================
'Function SetCategory(pMidCategoryCd)
'
'Dim RSv
'Dim vHTML
'Dim i
'
''---- ���J�e�S���[�A�J�e�S���[ ���o��
'wSQL = ""
'wSQL = wSQL & "SELECT a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , a.�J�e�S���[��"
'wSQL = wSQL & "  FROM �J�e�S���[ a WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[���J�e�S���[ b WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND b.���J�e�S���[�R�[�h = '" & pMidCategoryCd & "'"
'wSQL = wSQL & "   AND A.Web�J�e�S���[�t���O = 'Y'"
'wSQL = wSQL & " ORDER BY a.�\����"
'
''@@@@@@@@@@Response.Write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'i = 0
'vHTML = ""
'
'Do Until RSv.EOF = True
'	vHTML = vHTML & "            <a href='SearchList.asp?i_type=c&s_category_cd=" & RSv("�J�e�S���[�R�[�h") & "'>- " & RSv("�J�e�S���[��") & "</a><br>" & vbNewLine
'
'	i = i + 1
'	If AllFl <> "Y" And i >= 4 Then
'		Exit Do
'	End If
'
'	RSv.MoveNext
'Loop
'
'If AllFl <> "Y" Then
'	vHTML = vHTML & "            <a href='LargeCategoryList.asp?AllFl=Y&LargeCategoryCd=" & LargeCategoryCd & "'><strong>+ �S�Ă�����</strong></a>"
'End If
'
'RSv.Close
'
'SetCategory = vHTML
'
'End Function
'
''========================================================================
''
''	Function	�g�s�b�N�X News
''
''========================================================================
'Function CreatewNewsHTML()
'
'Dim RSv
'
''---- ���i�L�� ���o��
'wSQL = ""
'' 2012/01/20 GV Mod Start
''wSQL = wSQL & "SELECT TOP 5 * "
''wSQL = wSQL & "FROM "
''wSQL = wSQL & "(SELECT "
''wSQL = wSQL & "      a.�L���ԍ� "
''wSQL = wSQL & "    , a.�L�����t "
''wSQL = wSQL & "    , a.�L���^�C�g�� "
''wSQL = wSQL & " FROM "
''wSQL = wSQL & "      ���i�L�� a WITH (NOLOCK) "
''wSQL = wSQL & "    , ���i�L�����J�e�S���[ b WITH (NOLOCK) "
''wSQL = wSQL & "    , ���J�e�S���[ c WITH (NOLOCK) "
''wSQL = wSQL & "WHERE b.�L���ԍ� = a.�L���ԍ�"
''wSQL = wSQL & "  AND c.���J�e�S���[�R�[�h = b.���J�e�S���[�R�[�h"
''wSQL = wSQL & "  AND c.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "'"
''wSQL = wSQL & "  AND ((getDate() BETWEEN a.�\������From AND a.�\������To)"
''wSQL = wSQL & "    OR (a.�\������From IS NULL AND a.�\������To IS NULL))"
''wSQL = wSQL & "UNION "
''wSQL = wSQL & "SELECT "
''wSQL = wSQL & "      a.�L���ԍ� "
''wSQL = wSQL & "    , a.�L�����t "
''wSQL = wSQL & "    , a.�L���^�C�g�� "
''wSQL = wSQL & " FROM "
''wSQL = wSQL & "      ���i�L�� a WITH (NOLOCK)  "
''wSQL = wSQL & "WHERE (a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "'"    '2010/12/07 an mod
''wSQL = wSQL & "   OR a.�L���敪 = '��ʋL��')"                             '2010/12/07 an add
''wSQL = wSQL & "  AND ((getDate() BETWEEN a.�\������From AND a.�\������To)"
''wSQL = wSQL & "    OR (a.�\������From IS NULL AND a.�\������To IS NULL)) "
''wSQL = wSQL & ")AS inLineView "
''wSQL = wSQL & "ORDER BY �L�����t DESC"
'wSQL = wSQL & "SELECT DISTINCT TOP 5 "
'wSQL = wSQL & "      a.�L���ԍ� "
'wSQL = wSQL & "    , a.�L�����t "
'wSQL = wSQL & "    , a.�L���^�C�g�� "
'wSQL = wSQL & "FROM "
'wSQL = wSQL & "    Web���i�L��TOP10 a WITH (NOLOCK) "
'wSQL = wSQL & "WHERE "														' 2012/01/20 GV Mod Mail�ɂĒ����˗��ׁ̈A�����ύX
'wSQL = wSQL & "        (    (    a.�L���敪 = '��ʋL��' "
'wSQL = wSQL & "              OR  a.�L���敪 = '�ʋL��') "
'wSQL = wSQL & "         AND a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "') "
'wSQL = wSQL & "    OR  (    a.�L���敪 = '��ʋL��' "
'wSQL = wSQL & "         AND a.��J�e�S���[�R�[�h = '' "
'wSQL = wSQL & "         AND a.���J�e�S���[�R�[�h = '') "
'wSQL = wSQL & "ORDER BY "
'' 2012/01/23 GV Mod Start
''wSQL = wSQL & "      a.�L�����t DESC "
'wSQL = wSQL & "      a.�L���ԍ� DESC "
'' 2012/01/23 GV Mod End
'' 2012/01/20 GV Mod End
'
''@@@@Response.Write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
''----- NewsHTML�ҏW
'wNewsHTML = ""
'
'If RSv.EOF = False Then
'	wNewsHTML = wNewsHTML & "<table width='794' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine
'
'	Do until RSv.EOF = True
'		wNewsHTML = wNewsHTML & "  <tr>" & vbNewLine
'		wNewsHTML = wNewsHTML & "    <td class='honbun'>" & fFormatDate(RSv("�L�����t")) & " <a href='News.asp?NewsNo=" & RSv("�L���ԍ�") & "' class='link'>" & RSv("�L���^�C�g��") & "</a></td>" & vbNewLine
'		wNewsHTML = wNewsHTML & "  </tr>" & vbNewLine
'		RSv.MoveNext
'	Loop
'
'	wNewsHTML = wNewsHTML & "</table>" & vbNewLine
'End If
'
'RSv.Close
'
'End Function
'
''========================================================================
''
''	Function	�g�s�b�N�X �V���i
''
''========================================================================
'Function CreateNewItemHTML()
'
'Dim RSv
'
''---- �V���i ���o��
'wSQL = ""
'' 2012/01/20 GV Mod Start
''wSQL = wSQL & "SELECT DISTINCT TOP 10"
''wSQL = wSQL & "       a.������"
''wSQL = wSQL & "     , a.���[�J�[�R�[�h"
''wSQL = wSQL & "     , a.���i�R�[�h"
''wSQL = wSQL & "     , a.���i��"
''wSQL = wSQL & "     , b.���[�J�[��"
''wSQL = wSQL & "     , c.�J�e�S���[��"
''wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
''wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
''wSQL = wSQL & "     , �J�e�S���[ c WITH (NOLOCK)"
''wSQL = wSQL & "     , �J�e�S���[���J�e�S���[ d WITH (NOLOCK)"
''wSQL = wSQL & "     , ���J�e�S���[ e WITH (NOLOCK)"
''wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
''wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
''wSQL = wSQL & "   AND d.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
''wSQL = wSQL & "   AND e.���J�e�S���[�R�[�h = d.���J�e�S���[�R�[�h"
''wSQL = wSQL & "   AND e.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "'"
''wSQL = wSQL & "   AND a.�I���� IS NULL"
''wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
''wSQL = wSQL & " ORDER BY a.������ DESC"
'wSQL = wSQL & "SELECT DISTINCT TOP 10 "
'wSQL = wSQL & "      a.������ "
'wSQL = wSQL & "    , a.���[�J�[�R�[�h "
'wSQL = wSQL & "    , a.���i�R�[�h "
'wSQL = wSQL & "    , a.���i�� "
'wSQL = wSQL & "    , a.���[�J�[�� "
'wSQL = wSQL & "    , a.�J�e�S���[�� "
'wSQL = wSQL & "FROM "
'wSQL = wSQL & "    Web�V���iTOP10 a WITH (NOLOCK) "
'wSQL = wSQL & "WHERE "
'wSQL = wSQL & "        a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
'wSQL = wSQL & "ORDER BY "
'wSQL = wSQL & "      a.������ DESC "
'' 2012/01/20 GV Mod End
'
''@@@@@@@@@@Response.Write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
''----- �V���iHTML�ҏW
'wNewItemHTML = ""
'
'If RSv.EOF = False Then
'	wNewItemHTML = wNewItemHTML & "<table width='794' border='0' cellspacing='0' cellpadding='0'>" & vbNewLine
'
'	Do until RSv.EOF = True
'		wNewItemHTML = wNewItemHTML & "  <tr>" & vbNewLine
'		wNewItemHTML = wNewItemHTML & "     <td class='honbun'>" & fFormatDate(RSv("������")) & " <a href='ProductDetail.asp?Item=" & RSv("���[�J�[�R�[�h") & "^" & Server.URLEncode(RSv("���i�R�[�h")) & "' class='link'>" & RSv("���i��") & " " & RSv("�J�e�S���[��") & " (" & RSv("���[�J�[��") & ")</a></td>" & vbNewLine
'		wNewItemHTML = wNewItemHTML & "  </tr>" & vbNewLine
'		RSv.MoveNext
'	Loop
'
'	wNewItemHTML = wNewItemHTML & "</table>" & vbNewLine
'
'End If
'
'RSv.Close
'
'End Function
' 2012/03/13 GV Del End

'========================================================================
'
'	Function	SALE&OUTLET���i
'	2012/07/24 ok Add
'========================================================================
Function CreateSaleAndOutletHTML()

Dim RSv
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
wSQL = ""
wSQL = wSQL & "SELECT "
' 2012/09/03 GV #1426 Mod Start
'wSQL = wSQL & "    TOP 5 "
wSQL = wSQL & "    TOP 20 "
' 2012/09/03 GV #1426 Mod End
wSQL = wSQL & "      a.���i�R�[�h "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���[�J�[�� "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.�̔��P�� "
wSQL = wSQL & "    , a.�O��̔��P�� "
wSQL = wSQL & "    , a.ASK���i�t���O "
wSQL = wSQL & "    , a.B�i�t���O "
wSQL = wSQL & "    , a.�����萔�� "
wSQL = wSQL & "    , a.������P�� "
wSQL = wSQL & "    , a.������󒍍ϐ��� "
wSQL = wSQL & "    , a.�O��P���ύX�� "
wSQL = wSQL & "    , a.B�i�t���O "
wSQL = wSQL & "    , a.B�i�P�� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web�Z�[�����i a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "    a.�Z�[���敪�ԍ� BETWEEN 1 AND 4"
wSQL = wSQL & " AND a.��J�e�S���[�R�[�h = '" & LargeCategoryCd & "' "
wSQL = wSQL & "ORDER BY NEWID() "

'@@@@@@@@@@Response.Write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic
wHTML = ""

If RSv.EOF = false Then
	'----- �Z�[�����iHTML�ҏW
	wHTML = wHTML & "<h2 class='subtitle_red'>" & wLargeCategoryName & "��SALE &amp; OUTLET</h2>" & vbNewLine
	wHTML = wHTML & "<div class='box'><div class='box_inner01'>" & vbNewLine
	wHTML = wHTML & "  <ul class='list'>" & vbNewLine

	Do Until RSv.EOF = True OR dcnt > 4
' 2012/09/03 GV #1426 Add Start
		ReDim Preserve w_MakerCd(cnt)
		w_MakerCd(cnt) = RSv("���[�J�[�R�[�h")
		ReDim Preserve w_ItemCd(cnt)
		w_ItemCd(cnt) = RSv("���i�R�[�h")
		wHTML1 = ""
' 2012/09/03 GV #1426 Add End
		wHTML1 = wHTML1 & "    <li><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h")) & "'>"
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
'					wHTML1 = wHTML1 & FormatNumber(v_exprice,0) & "�~�i�ō��j��<br>" & vbNewLine
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
					ReDim Preserve w_price1(cnt)
					w_price1(cnt) = FormatNumber(v_exprice,0)
' 2012/09/03 GV #1426 Add End
				'B�i�A����i�͔̔����i�������i�Ƃ��ĕ\��
				Else
'2013/03/19 GV mod start ---->
'�O��P���͂��΂炭�\�������Ȃ�
'					wHTML1 = wHTML1 & FormatNumber(v_price,0) & "�~�i�ō��j��<br>" & vbNewLine
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
				wHTML1 = wHTML1 & "(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)"
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
				wHTML1 = wHTML1 & "(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)"
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
				wHTML1 = wHTML1 & "(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			End If

			wHTML1 = wHTML1 & "</span></li>" & vbNewLine

		Else
			'---- B�i�P��
			If RSv("B�i�t���O") = "Y" Then
				v_price = calcPrice(RSv("B�i�P��"), wSalesTaxRate)
'2013/03/19 GV mod start ---->
'				wHTML1 = wHTML1 & "�y�킯����i�����z</span><a class='tip'>ASK<span>" & FormatNumber(v_price,0) & "�~(�ō�)</span>"
				wHTML1 = wHTML1 & "�y�킯����i�����z</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("B�i�P��"),0) & "�~(�Ŕ�)</span><br>"
				wHTML1 = wHTML1 & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)</span>"
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
				wHTML1 = wHTML1 & "�y��������z</span><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(v_price,0) & "�~(�Ŕ�)</span><br>"
				wHTML1 = wHTML1 & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)</span>"
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
				wHTML1 = wHTML1 & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(v_price,0) & "�~)</span>"
'2013/03/19 GV mod end   <----
' 2012/09/03 GV #1426 Add Start
				ReDim Preserve w_price2(cnt)
				w_price2(cnt) = FormatNumber(v_price,0)
' 2012/09/03 GV #1426 Add End
			End If

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

	wHTML = wHTML & "  </ul>" & vbNewLine
	wHTML = wHTML & "</div></div>" & vbNewLine
End If
wSaleAndOutletHTML = wHTML

RSv.Close

End Function

'========================================================================
'
'	Function	Web���i�t���O�`�F�b�N
'
'========================================================================
Function GetProductFlag(pMakerCd, pProductCd)

Dim RSv
GetProductFlag = ""

'---- Web���i�t���O���o��
wSQL = ""
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.Web���i�t���O "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web���i a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.���[�J�[�R�[�h = '" & pMakerCd & "' "
wSQL = wSQL & "    AND a.���i�R�[�h     = '" & pProductCd & "'"

'@@@@@@@@@@Response.Write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

If RSv.EOF = false Then
	GetProductFlag = RSv("Web���i�t���O")
End If

RSv.Close

End Function

'========================================================================
'
'	Function	Close database
'
'========================================================================
'
Function close_db()

Connection.Close
Set Connection = Nothing    '2011/08/01 an add

End Function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<meta name="robots" content="noindex,nofollow">
<title><%=wLargeCategoryName%> �ꗗ�b�T�E���h�n�E�X</title>
<%=wMetaTag%>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css" type="text/css">
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
        <li class="now"><%=wLargeCategoryName%></li>
      </ul>
    </div></div></div>

<%=fIncludeInsertHTML(wInsertHTMLPath1)%>

<!-- ��J�e�S���[�ɂ��āE�J�e�S���[����I�ԁE�ŐV�j���[�X�E�V���i -->

<%=wStaticHTML(0)%>

<%=wIchioshiHTML%>

<%=wStaticHTML(1)%>

<%=fIncludeInsertHTML(wInsertHTMLPath2)%>

<%=wSaleAndOutletHTML%>

<%=wStaticHTML(2)%>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/LargeCategoryList.js?20130805"></script>
<script type="text/javascript" src="../jslib/jquery.carouFredSel-5.5.0-packed.js"></script>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript">
var userAgent = window.navigator.userAgent.toLowerCase();
var appVersion = window.navigator.appVersion.toLowerCase();
if(userAgent.indexOf("msie")!=-1){
	if(appVersion.indexOf("msie 7.")!=-1){
		$("ul.cate_tab li a span").each(function(){
			if($(this).height()>20){
				$(this).css("top","4px");
			}
		});
	}
}
</script>
</body>
</html>