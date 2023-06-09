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
'2016/02/04 GV �V�K�쐬�B(�����ύX�L�����Z���@�\)
'2016.04.05 GV �o�׊������Ƃ͕ʂɁA�o�ד���ǉ��B(�����ύX�L�����Z���@�\)
'2016.06.03 GV ���������t���O��ǉ��B(�����X�e�[�^�X���C�Ή�)
'2018.01.12 GV �����m�F�����؂ꌩ�ς��蒍���͕ύX�L�����Z���s�B
'2018.12.21 GV Pa��Pal�Ή��B
'2020.02.05 GV ������DL�Ή��B
'2020.06.31 GV �~���������X�g�Ή��B(#2841)
'2022.03.23 GV �ƎҌ����T�C�g�B(#3110)
'
'========================================================================
'On Error Resume Next

Const PAGE_SIZE = 10			' �w����������1�y�[�W������̕\���s��

Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim wFlg						' ���s�t���O
Dim wIPage						' �\������y�[�W�ʒu (�p�����[�^)
Dim wYear						' �����@��
Dim oJSON						' JSON�I�u�W�F�N�g
Dim wOrderHidden				' ��\���t���O
Dim wOrderCancelled				' �L�����Z�������t���O
Dim wOrderShipping				' �����������t���O
Dim wSlipNo						' �����ԍ�
Dim wReceipt					' �̎���
Dim wDepositTerm				' �����m�F�����i���j
Dim wOrderGift					'�M�t�g�����t���O
Dim wTantouName					' �S���Ҏ��� 2022.03.23 GV add
Dim wTantouEmail				' �S����e_mail 2022.03.23 GV add

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
wFlg = True

' Get�p�����[�^
' �ڋq�ԍ�
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
	wFlg = False
End If


'�y�[�W�ԍ�
wIPage = ReplaceInput_NoCRLF(Trim(Request("page")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wIPage) = False) Or (cf_checkNumeric(wIPage) = False) Then
	wIPage = 1
Else
	wIPage = CLng(wIPage)
End If

'��������
wYear = ReplaceInput_NoCRLF(Trim(Request("year")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wYear) = False) Or (cf_checkNumeric(wYear) = False) Then
	wYear = null
Else
	wYear = CLng(wYear)
End If

'�󒍔ԍ�
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wOrderNo = null
Else
	wOrderNo = CLng(wOrderNo)
End If

'��\���t���O
wOrderHidden = ReplaceInput_NoCRLF(Trim(Request("hide")))
If ((IsNull(wOrderHidden) = True) Or (UCase(wOrderHidden) <> "Y")) Then
	wOrderHidden = "N"
Else
	wOrderHidden = "Y"
End If

'�L�����Z�������t���O
wOrderCancelled = ReplaceInput_NoCRLF(Trim(Request("cancelled")))
If ((IsNull(wOrderCancelled) = True) Or (UCase(wOrderCancelled) <> "Y")) Then
	wOrderCancelled = "N"
Else
	wOrderCancelled = "Y"
End If

'�����������t���O
wOrderShipping = ReplaceInput_NoCRLF(Trim(Request("shipping")))
If ((IsNull(wOrderShipping) = True) Or (UCase(wOrderShipping) <> "Y")) Then
	wOrderShipping = "N"
Else
	wOrderShipping = "Y"
End If

'�����ԍ�
wSlipNo = ReplaceInput_NoCRLF(Trim(Request("slip")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wSlipNo) = False) Or (cf_checkNumeric(wSlipNo) = False) Then
	wSlipNo = null
End If


'�̎����p
wReceipt = ReplaceInput_NoCRLF(Trim(Request("receipt")))
If ((IsNull(wReceipt) = True) Or (UCase(wReceipt) <> "Y")) Then
	wReceipt = "N"
Else
	wReceipt = "Y"
End If

If (wReceipt = "Y") Then
	'��\���t���O�𖳌�
	wOrderHidden = "N"
End If

'�M�t�g�����t���O
wOrderGift = ReplaceInput_NoCRLF(Trim(Request("gift")))
If ((IsNull(wOrderGift) = True) Or (UCase(wOrderGift) <> "Y")) Then
	wOrderGift = "N"
Else
	wOrderGift = "Y"
End If

'2022.03.23 GV add start
'�S���Ҏ���
wTantouName = ReplaceInput_NoCRLF(Trim(Request("tantou_name")))
wTantouName = CStr(wTantouName)

'�S����e_mail
wTantouEmail = ReplaceInput_NoCRLF(Trim(Request("tantou_email")))
wTantouEmail = CStr(wTantouEmail)
'2022.03.23 GV add end


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
Dim vWHERE
Dim i
Dim j
Dim vRS
Dim orderDate
Dim deleteDate
Dim shippingCompDate
Dim shippingDate
Dim shippingSuu
Dim itemPicSmall
Dim makerName
Dim makerChokusou
Dim itemName
Dim iro
Dim kikaku
Dim allCount
Dim orderTotalAm2
Dim usedPoint
Dim shipNo
Dim slipNo
Dim modifyFlag  '�ύX�\�t���O
Dim cancelFlag  '�L�����Z���\�t���O
Dim modifyNg    '�ύXNG���R
Dim cancelNg    '�L�����Z��NG���R
Dim dateTerm
Dim maxDate
Dim ngReason
Dim ffCd
Dim orderType
Dim modifiable
Dim setItemFlag
Dim promote
Dim estMemo
Dim buy
Dim webModCancelFlg
Dim webOutline
Dim btnOn '�{�^���\���t���O
Dim depositFlag ' ���������t���O 2016.06.03 GV add
Dim receiptFlag '�̎������s�t���O 2020.02.05 GV add
Dim receiptDate '�̎������s�� 2020.02.05 GV add
' 2018.01.12 GV add start
Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2
' 2018.01.12 GV end

Dim isOtherAmountOk '2018.12.03 GV add
Dim wPaymentMethodDetail '2018.12.21 GV add
Dim giftCustomerNo '2021.06.30 GV add
Dim giftNo '2021.06.30 GV add

'2022.03.23 GV add start
Dim tantouName
Dim tantouEmail

Dim searchTantouName  '�����S���Ҏ���
Dim searchTantouEmail '�����S����e_mail
'2022.03.23 GV add end

Set oJSON = New aspJSON

' ������
i = 0
j = 0
allCount = 0
modifyFlag = "Y"
cancelFlag = "Y"
modifiable = "Y"
maxDate = ""
ngReason = ""
ffCd = ""
promote = ""
isOtherAmountOk = True '2018.12.03 GV add
searchTantouName = ""  '2022.03.23 GV add
searchTantouEmail = "" '2022.03.23 GV add

' �󒍌`��(�J���}��؂�Ŏw��)
orderType = ""
orderType = orderType & "  'E-mail'"
orderType = orderType & " ,'FAX'"
orderType = orderType & " ,'�C���^�[�l�b�g'"
orderType = orderType & " ,'�g��'"
orderType = orderType & " ,'�d�b'"
orderType = orderType & " ,'�X��'"
orderType = orderType & " ,'���X'"
orderType = orderType & " ,'�X�}�[�g�t�H��'"
orderType = orderType & " ,'�M�t�g'" '2021.06.30 GV add

'�R���g���[���}�X�^���猩�ς���L���������擾 2018.01.12 GV add
call getEmaxCntlMst("��","�����m�F�҂�����","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
If (IsNull(wItemNum1)) Then
	wDepositTerm = 10
Else
	wDepositTerm = wItemNum1
End If


' ���͒l������̏ꍇ
If (wFlg = True) Then
	' ---------------------------
	'�������ԂɎw�肪�Ȃ��ꍇ
	If ((IsNull(wYear) = True) Or (wYear = "")) Then
	ElseIf (wYear = 6) Then
			dateTerm = " AND o1.���ϓ� >= DATEADD(mm, -6, '" & Date() & " 23:59:59') "
	Else
		' �N�̏ꍇ
		dateTerm = " AND o1.���ϓ� "
		dateTerm = dateTerm & "BETWEEN '" & wYear & "-01-01 00:00:00' "
		dateTerm = dateTerm & " AND '" & wYear & "-12-31 23:59:59' "
	End If

	'2022.03.23 GV add start
	'�����S���Ҏ����Ɏw�肪�Ȃ��ꍇ
	If ((IsNull(wTantouName) = True) Or (wTantouName = "")) Then
	Else
		searchTantouName = " AND o1.�����S���� = '" & wTantouName & "' "
	End If

	'�����S����e_mail�Ɏw�肪�Ȃ��ꍇ
	If ((IsNull(wTantouEmail) = True) Or (wTantouEmail = "")) Then
	Else
		searchTantouEmail = " AND o1.�ڋqE_mail = '" & wTantouEmail & "' "
	End If
	'2022.03.23 GV add end

	' ---------------------------
	' �������擾
	vSQL = ""
	vSQL = vSQL & "SELECT count(o.�󒍔ԍ�) AS cnt "
	vSQL = vSQL & "FROM "
	vSQL = vSQL & " (SELECT DISTINCT "
	vSQL = vSQL & "   o1.�ڋq�ԍ� "
	vSQL = vSQL & "  ,o1.�󒍔ԍ� "
	vSQL = vSQL & "  ,o1.���ϓ� "
	vSQL = vSQL & "  ,o1.�폜�� "

	If wReceipt <> "Y" Then
		vSQL = vSQL & "  ,ov.��\���t���O "
	End If

	vSQL = vSQL & "  FROM �� AS o1 "
	vSQL = vSQL & "    INNER JOIN �󒍖��� od1 WITH (NOLOCK) "
	vSQL = vSQL & "      ON od1.�󒍔ԍ� = o1.�󒍔ԍ� "
	vSQL = vSQL & "     AND od1.�Z�b�g�i�e���הԍ� = 0 "
	vSQL = vSQL & "    LEFT JOIN �󒍔�\�����X�g ov WITH (NOLOCK) "
	vSQL = vSQL & "      ON ov.�󒍔ԍ� = od1.�󒍔ԍ� "
	vSQL = vSQL & "     AND ov.�󒍖��הԍ� = od1.�󒍖��הԍ� "
	vSQL = vSQL & "    LEFT JOIN �����ԍ�View slip "
	vSQL = vSQL & "      ON slip.�󒍔ԍ� = od1.�󒍔ԍ� "
	vSQL = vSQL & "     AND slip.�󒍖��הԍ� = od1.�󒍖��הԍ� "
	vSQL = vSQL & "  WHERE o1.�ڋq�ԍ� = " & wCustomerNo & " "
	vSQL = vSQL & "    AND o1.�󒍌`�� IN (" & orderType & ") "


	' �����������t���O
	If wOrderShipping = "Y" Then
		vSQL = vSQL & "  AND od1.�󒍐��� > od1.�o�׍��v���� "
	End If

	'�̎���
	If wReceipt = "Y" Then
		'vSQL = vSQL & "  AND o1.�폜�� IS NULL "
		vSQL = vSQL & "  AND ISNULL(od1.Web�L�����Z���t���O, 'N') <> 'Y' "
	Else
		' ��\���t���O
		If wOrderHidden = "Y" Then
			vSQL = vSQL & "    AND ov.��\���t���O = 'Y' "
		Else
			'�M�t�g���[�h�ł͂Ȃ�
			If (wOrderGift = "N") Then
				vSQL = vSQL & "    AND ov.��\���t���O IS NULL "
			End If
		End If

		' �L�����Z�������t���O
		If wOrderCancelled = "Y" Then
			'vSQL = vSQL & "  AND o1.�폜�� IS NOT NULL "
			vSQL = vSQL & "  AND od1.Web�L�����Z���t���O = 'Y' "
		Else
			If wOrderHidden = "Y" Then
			'��\���t���O��Y�̏ꍇ�A���w��
			Else
				vSQL = vSQL & "  AND o1.�폜�� IS NULL "
				vSQL = vSQL & "  AND ISNULL(od1.Web�L�����Z���t���O, 'N') <> 'Y' "
			End If
		End If
	End If


	' �󒍔ԍ�
	If (IsNull(wOrderNo) = False) Or (wOrderNo <> "") Then
		vSQL = vSQL & " AND o1.�󒍔ԍ� = " & wOrderNo
	End If

	' �����ԍ����w�肳�ꂢ�Ă��� = �ǐՃy�[�W�ł̂ݎg�p
	If (IsNull(wSlipNo) = False) Or (wSlipNo <> "") Then
		vSQL = vSQL & " AND slip.�����ԍ� = '" & wSlipNo & "' "
	End If

	vSQL = vSQL & dateTerm

	'2022.03.23 GV add start
	vSQL = vSQL & searchTantouName  ' �����S���Ҏ���������
	vSQL = vSQL & searchTantouEmail ' �����S����e_mail������
	'2022.03.23 GV add end

	vSQL = vSQL & " ) AS o "

	'�������Ԃ�����
	'@@@@Response.Write vSQL & "<br>"

	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂��Ă���ꍇ
	If vRS.EOF = False Then
		allCount = vRS("cnt")
	End If

	'--- �Y���ڋq�̎󒍈ꗗ���o��
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & " o.* "
	vSQL = vSQL & " , od.�������͐於�O "
	vSQL = vSQL & " , od.ORG_�������͐於�O "
	vSQL = vSQL & " , od.�󒍖��הԍ� "
	vSQL = vSQL & " , od.���[�J�[�R�[�h "
	vSQL = vSQL & " , m.���[�J�[�� "
	vSQL = vSQL & " , od.���i�R�[�h "
	vSQL = vSQL & " , i.���i�� "
	vSQL = vSQL & " , od.�F "
	vSQL = vSQL & " , od.�K�i "
	vSQL = vSQL & " , iz.���iID "
	vSQL = vSQL & " , i.���i�摜�t�@�C����_�� "
	vSQL = vSQL & " , i.Web���i�t���O "
	vSQL = vSQL & " , i.�Z�b�g���i�t���O "
	vSQL = vSQL & " ,(CASE "
	vSQL = vSQL & "    WHEN i.B�i�t���O = 'Y' THEN iz.B�i�����\���� "
	vSQL = vSQL & "    ELSE iz.�����\���� "
	vSQL = vSQL & "  END) AS �݌ɐ� "
	vSQL = vSQL & " ,i.�戵���~�� "
	vSQL = vSQL & " ,i.�I���� "

	If wReceipt = "Y" Then
		vSQL = vSQL & " ,i.���i�T��web "
	End If

	vSQL = vSQL & " , od.���[�J�[�����t���O "
	vSQL = vSQL & " , od.�󒍒P�� "
	vSQL = vSQL & " , od.�󒍋��z "
	vSQL = vSQL & " , od.�󒍐��� "
	vSQL = vSQL & " , od.�o�׎w�����v���� "
	vSQL = vSQL & " , (CASE "
	vSQL = vSQL & "     WHEN o.�󒍓� IS NOT NULL AND o.�폜�� IS NULL THEN 'Y' "
	vSQL = vSQL & "     ELSE 'N' END "
	vSQL = vSQL & "   ) AS ���c "
	'vSQL = vSQL & " , iz.�K���݌ɐ��� "
	vSQL = vSQL & " , ISNULL(od.�K���݌ɐ���, 0) AS �K���݌ɐ��� "   '�������̓K���݌ɐ���
	vSQL = vSQL & " ,od.�����ԍ� "
	vSQL = vSQL & " ,od.�o�הԍ� "
	vSQL = vSQL & " ,od.�o�א��� "
	vSQL = vSQL & " ,od.�o�ד� "
	vSQL = vSQL & " ,od.�^����ЃR�[�h "
	vSQL = vSQL & ", od.�󒍖��ה��l "

	vSQL = vSQL & "FROM "

' ROW_NUMBER�t�̎󒍏��
	vSQL = vSQL & "(SELECT * FROM "
	vSQL = vSQL & "  (SELECT "
	vSQL = vSQL & "     ROW_NUMBER() OVER(ORDER BY o2.���ϓ� DESC) AS RN "
	vSQL = vSQL & "     ,o2.* "
	vSQL = vSQL & "   FROM "
	vSQL = vSQL & "     (SELECT DISTINCT "
	vSQL = vSQL & "         o1.�ڋq�ԍ� "
	vSQL = vSQL & "       , o1.�󒍔ԍ� "
	vSQL = vSQL & "       , o1.�󒍓� "
	vSQL = vSQL & "       , o1.���ϓ� "
	vSQL = vSQL & "       , o1.�폜�� "
	vSQL = vSQL & "       , o1.�o�׊����� "
	vSQL = vSQL & "       , o1.�󒍌`�� "
	vSQL = vSQL & "       , o1.�x�����@ "
	vSQL = vSQL & "       , o1.����ŗ� "
	vSQL = vSQL & "       , o1.�󒍍��v���z "
	vSQL = vSQL & "       , o1.���v���z "
	vSQL = vSQL & "       , o1.���p�|�C���g "
	vSQL = vSQL & "       , o1.Web�����ύX�L�����Z�����t���O "
	vSQL = vSQL & "       , o1.���̑����v���z "
	' 2018.12.03 GV add start
	vSQL = vSQL & ",(SELECT "
	vSQL = vSQL & " count(*) FROM �󒍂��̑����� other1 WITH(NOLOCK) "
	vSQL = vSQL & " WHERE "
	vSQL = vSQL & " other1.�󒍔ԍ� = o1.�󒍔ԍ� " 
	vSQL = vSQL & " AND �󒍂��̑��R�[�h <> 'COUPON' "
	vSQL = vSQL & ") as ���̑����׌��� "
	' 2018.12.03 GV add start

	vSQL = vSQL & "       , ISNULL(o1.�z����񖾍׎w��t���O, 'N') AS  �z����񖾍׎w��t���O "
	vSQL = vSQL & "       , o1.���������t���O " '2016.06.03 GV add
	vSQL = vSQL & "       , o1.�̎����ԍ� " '2020.02.05 GV add
	vSQL = vSQL & "       , o1.�̎������s�� " '2020.02.05 GV add
	vSQL = vSQL & "       , o1.�x�����@�ڍ� " '2018.12.21 GV add

	If (wReceipt <> "Y") And (wOrderGift = "N") Then
		vSQL = vSQL & "       ,ov.��\���t���O "
	End If

	vSQL = vSQL & "       , o1.�M�t�g�ڋq�ԍ� "
	vSQL = vSQL & "       , o1.�M�t�g�ԍ� "

	vSQL = vSQL & "       , o1.�����S���� " '2022.03.23 GV add
	vSQL = vSQL & "       , o1.�ڋqE_mail "   '2022.03.23 GV add

	vSQL = vSQL & "      FROM �� AS o1 "

	vSQL = vSQL & "      INNER JOIN �󒍖��� od1 WITH (NOLOCK) "
	vSQL = vSQL & "        ON od1.�󒍔ԍ� = o1.�󒍔ԍ� "
	vSQL = vSQL & "       AND od1.�Z�b�g�i�e���הԍ� = 0 "

	vSQL = vSQL & "      LEFT JOIN �󒍔�\�����X�g ov WITH (NOLOCK) "
	vSQL = vSQL & "        ON ov.�󒍔ԍ� = od1.�󒍔ԍ� "
	vSQL = vSQL & "       AND ov.�󒍖��הԍ� = od1.�󒍖��הԍ� "

	vSQL = vSQL & "       WHERE "
'	vSQL = vSQL & "             o1.�ڋq�ԍ� =  " & wCustomerNo '2021.06.30 GV mod

	'2021.06.30 GV add start
	'�M�t�g�������[�h
	If (wOrderGift = "Y") Then
		vSQL = vSQL & "             o1.�M�t�g�ڋq�ԍ� =  " & wCustomerNo
	'�ʏ�
	Else
		vSQL = vSQL & "             o1.�ڋq�ԍ� =  " & wCustomerNo
	End If
	'2021.06.30 GV add start

	vSQL = vSQL & "         AND o1.�󒍌`�� IN (" & orderType & " ) "

	'�̎���
	If wReceipt = "Y" Then
		'vSQL = vSQL & "  AND o1.�폜�� IS NULL "
		vSQL = vSQL & "  AND ISNULL(od1.Web�L�����Z���t���O, 'N') <> 'Y' "
	Else
		'��\���t���O (�L�����Z�����i���\��������j
		If wOrderHidden = "Y" Then
			vSQL = vSQL & "  AND ov.��\���t���O = 'Y' "
		Else
			'vSQL = vSQL & "  AND ov.��\���t���O IS NULL " '2021.06.30 GV mod

			' �M�t�g�����łȂ� 2021.06.30 GV add
			If (wOrderGift = "N") Then
				vSQL = vSQL & "  AND ov.��\���t���O IS NULL "
			End If

			' �L�����Z�������t���O
			If wOrderCancelled = "Y" Then
				vSQL = vSQL & "  AND od1.Web�L�����Z���t���O = 'Y' "
			Else
				vSQL = vSQL & "  AND o1.�폜�� IS NULL "
				vSQL = vSQL & "  AND ISNULL(od1.Web�L�����Z���t���O, 'N') <> 'Y' "
			End If
		End If
	End If



	' �����������t���O
	If wOrderShipping = "Y" Then
		vSQL = vSQL & "  AND od1.�󒍐��� > od1.�o�׍��v���� "
	End If

	' �����ԍ����w�肳�ꂢ�Ă��� = �ǐՃy�[�W�ł̂ݎg�p
'	If (IsNull(wSlipNo) = False) Or (wSlipNo <> "") Then
'		vSQL = vSQL & " AND slip.�����ԍ� = '" & wSlipNo & "' "
'	End If

	' �󒍔ԍ����w�肳��Ă��� = �ڍ׃y�[�W�ł̂ݎg�p
	If (IsNull(wOrderNo) = False) Or (wOrderNo <> "") Then
		If (wOrderGift = "N") Then
			vSQL = vSQL & " AND o1.�󒍔ԍ� = " & wOrderNo
		Else
			vSQL = vSQL & " AND o1.�M�t�g�ԍ� = " & wOrderNo
		End If
	End If

	vSQL = vSQL & dateTerm ' �����@�ւ�����

	'2022.03.23 GV add start
	vSQL = vSQL & searchTantouName  ' �����S���Ҏ���������
	vSQL = vSQL & searchTantouEmail ' �����S����e_mail������
	'2022.03.23 GV add end

	vSQL = vSQL & "     ) AS o2 " 'ROW_NUMBER �ƈꏏ�ɏo�͂��鍀��
	vSQL = vSQL & "  ) as o3 " 'ROW_NUMBER ���܂ގ󒍏��

	'����2
	vSQL = vSQL & "  WHERE "
	vSQL = vSQL & "    RN BETWEEN " & ((PAGE_SIZE * (wIPage - 1)) + 1) & " AND " & (PAGE_SIZE * wIPage)
	vSQL = vSQL & ") AS o "

	' �󒍖���
	vSQL = vSQL & "INNER JOIN ("
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "  od2.�󒍔ԍ� "
	vSQL = vSQL & ", (CASE "
	vSQL = vSQL & "     WHEN o4.�󒍌`�� = '�M�t�g' THEN gift_c.�n���h���l�[�� "
	vSQL = vSQL & "     ELSE od2.�������͐於�O END "
	vSQL = vSQL & "   ) AS �������͐於�O "
	vSQL = vSQL & ", od2.�������͐於�O AS ORG_�������͐於�O "

	vSQL = vSQL & ", od2.�󒍖��הԍ� "
	vSQL = vSQL & ", od2.���[�J�[�R�[�h "
	vSQL = vSQL & ", od2.���i�R�[�h "
	vSQL = vSQL & ", od2.�F "
	vSQL = vSQL & ", od2.�K�i "
	vSQL = vSQL & ", od2.���[�J�[�����t���O "
	vSQL = vSQL & ", od2.�󒍒P�� "
	vSQL = vSQL & ", od2.�󒍋��z "
	vSQL = vSQL & ", od2.�󒍐��� "
	vSQL = vSQL & ", od2.�o�׎w�����v���� "
	vSQL = vSQL & ", od2.�󒍖��ה��l "
	vSQL = vSQL & ", od2.�K���݌ɐ��� "

' �����������̏ꍇ
If (wOrderShipping = "Y") Then
	vSQL = vSQL & ", '' AS �����ԍ� "
	vSQL = vSQL & ", '' AS �o�הԍ� "
	vSQL = vSQL & ", '' AS �o�א��� "
	vSQL = vSQL & ", '' AS �o�ד� "
	vSQL = vSQL & ", '' AS �^����ЃR�[�h "
Else
	vSQL = vSQL & ", slip.�����ԍ� "
	vSQL = vSQL & ", slip.�o�הԍ� "
	vSQL = vSQL & ", slip.�o�א��� "
	vSQL = vSQL & ", slip.�o�ד� "
	vSQL = vSQL & ", slip.�^����ЃR�[�h "
End If

	vSQL = vSQL & "FROM �󒍖��� od2 WITH (NOLOCK) "

	vSQL = vSQL & "INNER JOIN �� o4 WITH (NOLOCK) "
	vSQL = vSQL & "  ON o4.�󒍔ԍ� = od2.�󒍔ԍ� "

If (wOrderGift = "N") Then
	vSQL = vSQL & "LEFT JOIN �󒍔�\�����X�g ov2 WITH (NOLOCK) "
	vSQL = vSQL & "  ON ov2.�󒍔ԍ� = od2.�󒍔ԍ� "
	vSQL = vSQL & " AND ov2.�󒍖��הԍ� = od2.�󒍖��הԍ� "
End If

' �����������̏ꍇ
If (wOrderShipping = "Y") Then
Else
	vSQL = vSQL & "LEFT JOIN �����ԍ�View slip "
	vSQL = vSQL & "  ON slip.�󒍔ԍ� = od2.�󒍔ԍ� "
	vSQL = vSQL & " AND slip.�󒍖��הԍ� = od2.�󒍖��הԍ� "
End If

	vSQL = vSQL & "LEFT JOIN �ڋq gift_c WITH (NOLOCK) "
	vSQL = vSQL & "  ON gift_c.�ڋq�ԍ� = o4.�M�t�g�ڋq�ԍ� "

	vSQL = vSQL & "WHERE "
	vSQL = vSQL & "     od2.�Z�b�g�i�e���הԍ� = 0 "

	'�̎���
	If wReceipt = "Y" Then
		vSQL = vSQL & "  AND ISNULL(od2.Web�L�����Z���t���O, 'N') <> 'Y' "
	Else
		'��\���t���O
		If wOrderHidden = "Y" Then
			vSQL = vSQL & "  AND ov2.��\���t���O = 'Y' "
		Else
			' �M�t�g���[�h�łȂ�
			If (wOrderGift = "N") Then
				vSQL = vSQL & "  AND ov2.��\���t���O IS NULL "
			End If

			' �L�����Z�������t���O
			If wOrderCancelled = "Y" Then
				'vSQL = vSQL & "  AND o4.�폜�� IS NOT NULL "
				vSQL = vSQL & "  AND od2.Web�L�����Z���t���O = 'Y' "
			Else
				vSQL = vSQL & "  AND ISNULL(od2.Web�L�����Z���t���O, 'N') <> 'Y' "
			End If
		End If
	End If

	If (IsNull(wSlipNo) = False) Or (wSlipNo <> "") Then
		vSQL = vSQL & " AND slip.�����ԍ� = '" & wSlipNo & "' "
	End If

	' �����������t���O
	If wOrderShipping = "Y" Then
		'vSQL = vSQL & "  AND slip.�����ԍ� IS NULL "
		'vSQL = vSQL & "  AND slip.�o�הԍ� IS NULL "
		vSQL = vSQL & "  AND od2.�󒍐��� > od2.�o�׍��v���� "
	End If


vSQL = vSQL & ") AS od "
vSQL = vSQL & "ON od.�󒍔ԍ� = o.�󒍔ԍ� "


	vSQL = vSQL & "INNER JOIN �F�K�i�ʍ݌� iz WITH (NOLOCK) "
	vSQL = vSQL & "   ON iz.���[�J�[�R�[�h = od.���[�J�[�R�[�h "
	vSQL = vSQL & "  AND iz.���i�R�[�h = od.���i�R�[�h "
	vSQL = vSQL & "  AND iz.�F = od.�F "
	vSQL = vSQL & "  AND iz.�K�i = od.�K�i "

	vSQL = vSQL & "INNER JOIN ���i i WITH (NOLOCK) "
	vSQL = vSQL & "   ON i.���[�J�[�R�[�h = iz.���[�J�[�R�[�h "
	vSQL = vSQL & "  AND i.���i�R�[�h = iz.���i�R�[�h "

	vSQL = vSQL & "INNER JOIN ���[�J�[ m WITH (NOLOCK) "
	vSQL = vSQL & "   ON m.���[�J�[�R�[�h = i.���[�J�[�R�[�h "

	vSQL = vSQL & " ORDER BY "
	vSQL = vSQL & "   ���ϓ� DESC"
	vSQL = vSQL & "   ,�󒍔ԍ� desc"
	vSQL = vSQL & "   ,�󒍖��הԍ� asc"
	'@@@@Response.Write(vSQL) & "<br>"


	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂��Ă���ꍇ
	If vRS.EOF = False Then

		' �S������JSON�f�[�^�ɃZ�b�g
		oJSON.data.Add "count" ,allCount

		' �y�[�W�ԍ�
		oJSON.data.Add "page" ,wIPage

		' �y�[�W������̍s��
		oJSON.data.Add "page_size" ,PAGE_SIZE

		' ��\��
		oJSON.data.Add "hidden" ,wOrderHidden

		' �L�����Z��
		oJSON.data.Add "cancelled" ,wOrderCancelled

		' ������
		oJSON.data.Add "shipping" ,wOrderShipping

		' �̎���
		oJSON.data.Add "receipt" ,wReceipt

		' ���X�g�ǉ�
		oJSON.data.Add "list" ,oJSON.Collection()

		For i = 0 To (vRS.RecordCount - 1)
			' �󒍓�
			If (IsNull(vRS("�󒍓�"))) Then
				orderDate = ""
			Else
				orderDate = CStr(Trim(vRS("�󒍓�")))
			End If

			' �o�׊�����
			If (IsNull(vRS("�o�׊�����"))) Then
				shippingCompDate = ""
			Else
				shippingCompDate = CStr(Trim(vRS("�o�׊�����")))
			End If

			' �o�ד�
			If (IsNull(vRS("�o�ד�"))) Then
				shippingDate = ""
			Else
				shippingDate = CStr(Trim(vRS("�o�ד�")))
			End If

			' �폜��
			If (IsNull(vRS("�폜��"))) Then
				deleteDate = ""
			Else
				deleteDate = CStr(Trim(vRS("�폜��")))
			End If

			If (IsNull(vRS("���i�摜�t�@�C����_��"))) Then
				itemPicSmall = ""
			Else
				itemPicSmall = CStr(vRS("���i�摜�t�@�C����_��"))
			End If


			If (IsNull(vRS("�Z�b�g���i�t���O"))) Then
				setItemFlag = ""
			Else
				setItemFlag = CStr(vRS("�Z�b�g���i�t���O"))
			End If

			makerName = Replace(Trim(vRS("���[�J�[��")), """", "�h")
			makerName = CStr(makerName)

			itemName = Replace(Trim(vRS("���i��")), """", "�h")
			itemName = CStr(itemName)

			iro = Replace(Trim(vRS("�F")), """", "�h")
			iro = CStr(iro)

			kikaku = Replace(Trim(vRS("�K�i")), """", "�h")
			kikaku = CStr(kikaku)

			If (IsNull(vRS("���[�J�[�����t���O"))) Then
				makerChokusou = ""
			Else
				makerChokusou = CStr(vRS("���[�J�[�����t���O"))
			End If

			If (IsNull(vRS("���v���z"))) Then
				orderTotalAm2 = 0
			Else
				orderTotalAm2 = CDbl(vRS("���v���z"))
			End If

			'�����ԍ�
			If (IsNull(vRS("�����ԍ�"))) Then
				slipNo = ""
			Else
				slipNo = CStr(vRS("�����ԍ�"))
			End If

			' �o�א���
			If (IsNull(vRS("�o�א���")) Or (vRS("�o�א���") = "")) Then
				shippingSuu = 0
			Else
				shippingSuu = CDbl(vRS("�o�א���"))
			End If

			' ���p�|�C���g
			If (IsNull(vRS("���p�|�C���g"))) Then
				usedPoint = 0
			Else
				usedPoint = CDbl(vRS("���p�|�C���g"))
			End If

			' �^����ЃR�[�h
			If (IsNull(vRS("�^����ЃR�[�h"))) Then
				ffCd = ""
			Else
				ffCd = CStr(vRS("�^����ЃR�[�h"))
			End If

			'2016.06.03 GV add start
			'���������t���O
			If (IsNull(vRS("���������t���O"))) Then
				depositFlag = ""
			Else
				depositFlag = CStr(Trim(vRS("���������t���O")))
			End If
			'2016.06.03 GV add start

			'2020.02.05 GV add start
			'�̎������s�t���O
			receiptFlag = getReceiptFlag(vRS("�x�����@"), CStr(Trim(vRS("�󒍔ԍ�"))))

			'�̎������s��
			If (IsNull(vRS("�̎������s��"))) Then
				receiptDate = ""
			Else
				receiptDate = CStr(Trim(vRS("�̎������s��")))
			End If
			'2020.02.05 GV add end

			'�̑��i����
			promote = "N"
			If (CDbl(Trim(vRS("�󒍒P��"))) = 0) Then
				'�󒍖��ה��l�Ɂu�̑��i�v�Ɗ܂܂��ꍇ�A
				'estMemo = InStr(Trim(vRS("�󒍖��ה��l")), "�̑��i")
				'If (IsNull(estMemo) = False) And (IsNumeric(estMemo)) And (estMemo > 0) Then
				If (InStr(Trim(vRS("�󒍖��ה��l")), "�̑��i") > 0) Then
					promote = "Y"
				ElseIf (InStr(Trim(vRS("���i�R�[�h")), "HOTMENU") > 0) Then
					promote = "Y"
				End If
			End If

			'Web�����ύX�L�����Z�����t���O
			If (IsNull(vRS("Web�����ύX�L�����Z�����t���O"))) Then
				webModCancelFlg = "N"
			Else
				If (Trim(vRS("Web�����ύX�L�����Z�����t���O")) <> "Y") Then
					webModCancelFlg = "N"
				Else
					webModCancelFlg = "Y"
				End If
			End If

			' 2018.12.21 GV add start
			'�x�������@�ڍ�
			If (IsNull(vRS("�x�����@�ڍ�"))) Then
				wPaymentMethodDetail = ""
			Else
				wPaymentMethodDetail = CStr(vRS("�x�����@�ڍ�"))
			End If
			' 2018.12.21 GV add end

			' �M�t�g�ڋq�ԍ� 2021.06.30 GV add
			If (IsNull(vRS("�M�t�g�ڋq�ԍ�"))) Then
				giftCustomerNo = 0
			Else
				giftCustomerNo =CStr(vRS("�M�t�g�ڋq�ԍ�"))
			End If

			' �M�t�g�ԍ� 2021.06.30 GV add
			If (IsNull(vRS("�M�t�g�ԍ�"))) Then
				giftNo = 0
			Else
				giftNo = CStr(vRS("�M�t�g�ԍ�"))
			End If

			'2022.03.23 GV add start
			' �����S����
			If (IsNull(vRS("�����S����"))) Then
				tantouName = ""
			Else
				tantouName = CStr(vRS("�����S����"))
			End If

			' �ڋqE_mail
			If (IsNull(vRS("�ڋqE_mail"))) Then
				tantouEmail = ""
			Else
				tantouEmail = CStr(vRS("�ڋqE_mail"))
			End If
			'2022.03.23 GV add end

			' ---------------------------------------------------
			'�ύX�\����
			' ---------------------------------------------------
			modifyFlag = "Y"  '�ύX�\�t���O
			modifyNg = ""     '�ύXNG���R
			cancelFlag = "Y"  '�L�����Z���\�t���O
			cancelNg = ""     '�L�����Z��NG���R
			ngReason = ""     '
			btnOn = "Y"       '�{�^���\������

			'�o�׊������Ă���ꍇ�A�{�^����\��
			If (shippingCompDate <> "") Then
				btnOn = "N"
			Else
				'�폜����Ă��Ȃ��@���A�o�׎w�����������Ă��Ȃ�
				If (deleteDate = "") And (vRS("�o�׎w�����v����") = 0) Then
					btnOn = "Y"
				End If
			End If

			' �̑��i�ł͂Ȃ��AWeb�Ɍf�ڂ��Ă��鏤�i�ł���
			If (promote <> "Y" And Trim(vRS("Web���i�t���O")) = "Y") Then
				'Web�ύX�L�����Z���t���O��N
				If (webModCancelFlg = "N") Then
					If (vRS("�z����񖾍׎w��t���O") <> "Y") Then
						' 2018.12.03 GV add start
						If (vRS("���̑����v���z") <> 0) Then
							' �N�[�|���ȊO�̎󒍂��̑����ׂ����݂���ꍇ�ANG
							If (vRS("���̑����׌���") > 0) Then
								isOtherAmountOk = False
							End If
						End If
						' 2018.12.03 GV add end

						'If (vRS("���̑����v���z") = 0) Then ' 2018.12.03 GV mod
						If (isOtherAmountOk) Then
							If ((vRS("�󒍌`��") = "�C���^�[�l�b�g") Or (vRS("�󒍌`��") = "�X�}�[�g�t�H��")) Then
								If (Mid(vRS("�x�����@"), 1, 3) <> "���[��") Then
										If (vRS("�o�׎w�����v����") = 0) Then
											If (vRS("���[�J�[�����t���O") <> "Y") Then
												'If (((orderDate = "") And (deleteDate = "")) And (vRS("�K���݌ɐ���") > 0)) Then
												'��荞�܂ꂽ�����̏��
												If (orderDate = "") And (deleteDate = "") Then
													'�ύX�L�����Z���\
													'modifyFlag = "Y"
												Else
													'�Z�b�g�i�̏ꍇ�͓K���݌ɐ��ʂ��݂Ȃ�
													If (vRS("�Z�b�g���i�t���O") = "Y") Then
														If (((orderDate <> "") And (deleteDate = ""))) Then
															'�ύX�L�����Z���\
															'modifyFlag = "Y"
														Else
															ngReason = "5"
															btnOn = "N" '2018.01.12 GV add
															modifyFlag = "N"
															cancelFlag = "N"
														End If
													Else
														'�̑��i
														If (promote = "Y") Then
															If (((orderDate <> "") And (deleteDate = ""))) Then
																'�ύX�L�����Z���\
																'modifyFlag = "Y"
															Else
																ngReason = "5"
																btnOn = "N" '2018.01.12 GV add
																modifyFlag = "N"
																cancelFlag = "N"
															End If
														Else
															If (((orderDate <> "") And (deleteDate = "")) And (vRS("�K���݌ɐ���") > 0)) Then
																'�ύX�L�����Z���\
																'modifyFlag = "Y"
															Else
																If (((orderDate <> "") And (deleteDate = "")) And (vRS("�K���݌ɐ���") < 1)) Then
																	ngReason = "5"
																	cancelFlag = "N" '�L�����Z���͕s�����A�ύX�͎󂯕t����
																End If '�K���݌�
															End If '�K���݌�
														End If '�̑��i
													End If '�Z�b�g�i
												End If
											Else
												ngReason = "4" '���[�J�[�����t���O
												btnOn = "N" '2018.01.12 GV add
												modifyFlag = "N"
												cancelFlag = "N"
											End If
										Else
											ngReason = "3" '�o�׎w��
											btnOn = "N"
											modifyFlag = "N"
											cancelFlag = "N"
										End If
								Else
									ngReason = "2" '�x�����@
									btnOn = "N" '2018.01.12 GV add
									modifyFlag = "N"
									cancelFlag = "N"
								End If
							Else
								ngReason = "1" '�󒍌`��
								btnOn = "N" '2018.01.12 GV add
								modifyFlag = "N"
								cancelFlag = "N"
							End If
						Else
							ngReason = "10" '���̑����v���z
							btnOn = "N" '2018.01.12 GV add
							modifyFlag = "N"
							cancelFlag = "N"
						End If
					Else
						ngReason = "11" '�z����񖾍׎w��t���O
						btnOn = "N" '2018.01.12 GV add
						modifyFlag = "N"
						cancelFlag = "N"
					End If
				Else
					ngReason = "8" 'Web�ύX�L�����Z����
					btnOn = "N"
					modifyFlag = "N"
					cancelFlag = "N"
				End If
			Else
				' �̑��i�ł͂Ȃ��AWeb�Ɍf�ڂ��Ă��鏤�i�łȂ�
				If (promote <> "Y") And (Trim(vRS("Web���i�t���O")) <> "Y") Then
					ngReason = "9" 'Web���i�t���O
					btnOn = "N" '2018.01.12 GV add
					modifyFlag = "N"
					cancelFlag = "N"
				Else
					'�̑��i�̏ꍇ�A�L�����Z���ۂ͔��肵�Ȃ�(�ύX�s�Ƃ͂��Ȃ��j
					'PHP���ŃL�����Z���I����s�Ƃ���
					'modifyFlag = "Y"
				End If
			End If

			' ---------------------------------------------------

			'�ύX�s���R���������ꍇ
			'If (ngReason <> "") Then
			'	'�ύX�s��
			'	modifiable = "N"
			'End If


			' �^����ЃR�[�h
			If (IsNull(vRS("�o�הԍ�"))) Then
				shipNo = ""
			Else
				shipNo = CStr(vRS("�o�הԍ�"))
			End If

			'�ēx�w���\���t���O
			buy = "N"

			If (Trim(vRS("Web���i�t���O")) = "Y") Then
				If (IsNull(vRS("�݌ɐ�")) = false) Then
					If vRS("�݌ɐ�") > 0 Then
						buy = "Y" '�ēx�w���\��
					Else
						If (IsNull(vRS("�戵���~��")) = true) And (IsNull(vRS("�I����")) = true) Then
							buy = "Y" '�ēx�w���\��
						End If
					End If
				End If
			End If 


			If wReceipt = "Y" Then
				If (IsNull(vRS("���i�T��web"))) Then
					webOutline = ""
				Else
					webOutline = Replace(Trim(vRS("���i�T��web")), """", "�h")
					webOutline = Replace(webOutline, vbCrLf, "")
				End If
			Else
				webOutline = ""
			End If

			'2018.01.12 GV add start
			'�폜����Ă��Ȃ��A���ς����ԁA�����������Ă��Ȃ�
			If ((deleteDate = "") And (orderDate = "") And (depositFlag <> "Y")) Then
				'���ϓ���Null�łȂ��A�{���Ƃ̍���������m�F�����ȏ�
				If (IsNull(vRS("���ϓ�")) = False) And (DateDiff("d", vRS("���ϓ�"), Now()) >= CInt(wDepositTerm)) Then
					ngReason = "12" '�����m�F�����؂�
					modifyFlag = "N"
					cancelFlag = "N"
					btnOn = "N"
				End If
			End If
			'2018.01.12 GV add end



			'--- ���׍s����
			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
					.Add "o_no" ,CStr(Trim(vRS("�󒍔ԍ�")))
					.Add "o_dt" ,orderDate '�󒍓�
					.Add "est_dt" ,CStr(Trim(vRS("���ϓ�")))
					.Add "ship_comp_dt" , shippingCompDate  '�o�׊�����
					.Add "del_dt" ,deleteDate '�폜��
					.Add "o_type" ,CStr(Trim(vRS("�󒍌`��")))
					.Add "pay_method" ,get_paymetMethodWord(vRS("�x�����@"))
					.Add "pay_method_detail" ,wPaymentMethodDetail '2018.12.21 GV add
					.Add "tax_rate", CDbl(vRS("����ŗ�")) 
					.Add "total_order_am", CDbl(vRS("�󒍍��v���z")) 
					.Add "total_order_am2",  orderTotalAm2  ' ���v���z
					.Add "used_pt", usedPoint  ' ���p�|�C���g
					.Add "ff_cd" ,ffCd ' �^����ЃR�[�h
					.Add "ship_name" ,CStr(Trim(vRS("�������͐於�O")))
					.Add "org_ship_name" ,CStr(Trim(vRS("ORG_�������͐於�O")))
					.Add "web_flg", CStr(vRS("Web���i�t���O"))
					.Add "buy", buy '�ēx�w���\���\
					.Add "set_flg", setItemFlag
					.Add "od_no" ,CStr(Trim(vRS("�󒍖��הԍ�")))
					.Add "m_cd" ,CStr(Trim(vRS("���[�J�[�R�[�h")))
					.Add "m_name" ,makerName
					.Add "i_cd" ,CStr(Trim(vRS("���i�R�[�h")))
					.Add "i_name" ,itemName
					.Add "iro" ,iro
					.Add "kikaku" ,kikaku
					.Add "i_id" ,CStr(Trim(vRS("���iID")))
					.Add "i_pic", itemPicSmall
					.Add "outline", webOutline
					.Add "m_chokusou", makerChokusou
					.Add "i_tanka", CDbl(Trim(vRS("�󒍒P��")))
					.Add "i_am", CDbl(Trim(vRS("�󒍋��z")))
					.Add "i_suu", CDbl(vRS("�󒍐���")) 
					.Add "ship_inst_suu", CDbl(vRS("�o�׎w�����v����"))
					.Add "o_zan", CStr(Trim(vRS("���c"))) 
					.Add "t_zaiko_suu", CDbl(vRS("�K���݌ɐ���"))
					.Add "slip_no", slipNo
					.Add "ship_no", shipNo '�o�הԍ�
					.Add "ship_suu", shippingSuu
					.Add "ship_dt" , shippingDate  '�o�ד�
					.Add "promote" , promote '�̑��i����
					.Add "modify_flg", modifyFlag '�ύX�\�t���O
					.Add "cancel_flg", cancelFlag '�L�����Z���\�t���O
					.Add "ng_rsn", ngReason
					.Add "btn_on", btnOn '�{�^���\������
					.Add "modifying", webModCancelFlg
					.Add "deposit", depositFlag '���������t���O 2016.06.03 GV add
					.Add "receipt_flg", receiptFlag '�̎������s�t���O 2020.02.05 GV add
					.Add "receipt_no" ,CStr(Trim(vRS("�̎����ԍ�"))) '�̎����ԍ� 2020.02.05 GV add
					.Add "receipt_dt" , receiptDate '�̎������s�� 2020.02.05 GV add
					.Add "gift_cst_no" , giftCustomerNo '�M�t�g�ڋq�ԍ� 2021.06.30 GV add
					.Add "gift_no" , giftNo '�M�t�g�ԍ� 2021.06.30 GV add
					.Add "tantou_name", tantouName '�����S���� 2022.03.23 GV add
					.Add "tantou_email", tantouEmail '�ڋqE_mail 2022.03.23 GV add
				End With
			End With

			' ���̃��R�[�h�s�ֈړ�
			vRS.MoveNext

			If vRS.EOF Then
				Exit For
			End If

			j = j + 1
		Next

		'�󒍔ԍ��w��̏ꍇ
		'If (wOrderNo <> "") Then
		'	' �ύX�\���Z�b�g
		'	oJSON.data.Add "modifiable" ,modifiable
		'End If
	End If

	'���R�[�h�Z�b�g�����
	vRS.Close

	'���R�[�h�Z�b�g�̃N���A
	Set vRS = Nothing
End If

' -------------------------------------------------
' JSON�f�[�^�̕ԋp
' -------------------------------------------------
' �w�b�_�o��
Response.AddHeader "Content-Type", "application/json; charset=shift_jis"
Response.AddHeader "Cache-Control", "no-cache,must-revalidate"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "X-Content-Type-Options", "nosniff"

' JSON�f�[�^�̏o��
Response.Write oJSON.JSONoutput()

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
'	Function	Emax�̃R���g���[���}�X�^����f�[�^�擾
'
'========================================================================

Function getEmaxCntlMst(pSubSystemCd, pItemCd, pItemSubCd, pItemChar1, pItemChar2, pItemNum1, pItemNum2, pItemDate1, pItemDate2)

Dim RS_cntl
Dim v_sql

'---- �R���g���[���}�X�^���o��

v_sql = ""
v_sql = v_sql & "SELECT a.*"
v_sql = v_sql & "  FROM �R���g���[���}�X�^ a WITH (NOLOCK)"
v_sql = v_sql & " WHERE a.sub_system_cd = '" & pSubSystemCd & "'"
v_sql = v_sql & "   AND a.item_cd = '" & pItemCd & "'"
v_sql = v_sql & "   AND a.item_sub_cd = '" & pItemSubCd & "'"

'@@@@@@response.write(v_sql)

Set RS_cntl = Server.CreateObject("ADODB.Recordset")
RS_cntl.Open v_sql, ConnectionEmax, adOpenStatic

If RS_cntl.EOF <> True Then
	pItemChar1 = RS_cntl("item_char1")
	pItemChar2 = RS_cntl("item_char2")
	pItemNum1 = RS_cntl("item_num1")
	pItemNum2 = RS_cntl("item_num2")
	pItemDate1 = RS_cntl("item_date1")
	pItemDate2 = RS_cntl("item_date2")
End If

RS_cntl.Close

End Function
'========================================================================
%>
