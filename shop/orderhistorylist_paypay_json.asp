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
'2020.02.27 GV orderhistorylist_v2_json.asp���x�[�X�ɐV�K�쐬�B(PayPay�Ή�)(#2405)
'2020.06.01 GV �C���B
'2020.06.03 GV PayPay�ԋ����C�B(#2440)
'2021.01.04 GV �����e�i���X�|�[�^��PayPay�ԋ��@�\���C�B(#2647)
'2021.01.21 GV �����e�i���X�|�[�^��PayPay�ԋ��@�\���C�B(#2662)
'
'========================================================================
'On Error Resume Next

Const PAGE_SIZE = 20			' 1�y�[�W������̕\���s��

Dim ConnectionEmax

Dim wErrMsg						' �G���[���b�Z�[�W (���̃y�[�W����n����郁�b�Z�[�W)
Dim wDispMsg					' �ʏ탁�b�Z�[�W(�G���[�ȊO) (���̃y�[�W����n����郁�b�Z�[�W)
Dim wErrDesc
Dim wMsg						' �G���[���b�Z�[�W (�{�y�[�W�ō쐬���郁�b�Z�[�W)
Dim wCustomerNo					' �ڋq�ԍ�
Dim wOrderNo					' �󒍔ԍ�
Dim wFlg						' ���s�t���O
Dim wIPage						' �\������y�[�W�ʒu (�p�����[�^)
Dim estimateStartDate			' �������Ԏ�
Dim estimateEndDate				' �������Ԏ�
Dim oJSON						' JSON�I�u�W�F�N�g
Dim wOrderHidden				' ��\���t���O
Dim wOrderCancelled				' �L�����Z�������t���O
Dim wOrderShipping				' �����������t���O
Dim wSlipNo						' �����ԍ�
Dim wReceipt					' �̎���
Dim wDepositTerm				' �����m�F�����i���j
Dim wPaypayPaymentId			' PayPay���ϔԍ�(�J�[�h�^�M�m�F�ԍ�)

'=======================================================================
'	�󂯓n�������o�� & �����ݒ�
'=======================================================================
wFlg = True

' Get�p�����[�^
' �ڋq�ԍ�
wCustomerNo = ReplaceInput_NoCRLF(Trim(Request("cno")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
'If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
'	wFlg = False
'End If
If (IsNull(wCustomerNo) = False) And (wCustomerNo <> "") Then
	If (IsNumeric(wCustomerNo) = False) Or (cf_checkNumeric(wCustomerNo) = False) Then
		wFlg = False
	End If
End If



'�y�[�W�ԍ�
wIPage = ReplaceInput_NoCRLF(Trim(Request("page")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wIPage) = False) Or (cf_checkNumeric(wIPage) = False) Then
	wIPage = 1
Else
	wIPage = CLng(wIPage)
End If

'�������Ԏ�
estimateStartDate = ReplaceInput_NoCRLF(Trim(Request("est_from")))
estimateStartDate = CStr(estimateStartDate)

'�������Ԏ�
estimateEndDate = ReplaceInput_NoCRLF(Trim(Request("est_to")))
estimateEndDate = CStr(estimateEndDate)


'�󒍔ԍ�
wOrderNo = ReplaceInput_NoCRLF(Trim(Request("ono")))
' ���l�̂݃`�F�b�N (ASP�͑S�p�ł������Ȃ�True��Ԃ�)
If (IsNumeric(wOrderNo) = False) Or (cf_checkNumeric(wOrderNo) = False) Then
	wOrderNo = null
Else
	wOrderNo = CLng(wOrderNo)
End If

'PayPay���ϔԍ�
wPaypayPaymentId = ReplaceInput_NoCRLF(Trim(Request("pay_id")))
wPaypayPaymentId = CStr(wPaypayPaymentId)

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
Dim orderName
Dim customerName
Dim shippingCompDate
Dim allCount
Dim orderTotalAm2
Dim usedPoint
Dim dateTerm
Dim orderType
Dim webModCancelFlg
Dim depositFlag ' ���������t���O
Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2
Dim wPaymentMethodDetail
Dim ccTotalAm
Dim ccCreditNo
Dim ccSlipNo
Dim isEstimateDateExist
Dim totalAmAtAuth '�I�[�\�����󒍍��v���z 2021.01.04 GV add

'2020.06.01 GV add
Dim totalAmAtOrder
Dim usedPointAtOrder
Dim kabusokuAmAtOrder
'2020.06.01 GV add

Dim shipSuuSum '2020.06.03 GV add

Set oJSON = New aspJSON

' ������
i = 0
j = 0
allCount = 0
dateTerm = ""

' �󒍌`��(�J���}��؂�Ŏw��)
orderType = ""
orderType = orderType & "  '�C���^�[�l�b�g'"
'orderType = orderType & " ,'E-mail'"
'orderType = orderType & " ,'FAX'"
'orderType = orderType & " ,'�g��'"
'orderType = orderType & " ,'�d�b'"
'orderType = orderType & " ,'�X��'"
'orderType = orderType & " ,'���X'"
orderType = orderType & " ,'�X�}�[�g�t�H��'"

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
	'�������Ԏ�
	If ((IsNull(estimateStartDate) = True) Or (estimateStartDate = "")) Then
	ElseIf ((IsNull(estimateStartDate) = False) Or (estimateStartDate <> "")) Then
		dateTerm = dateTerm & " AND o1.���ϓ� >= '" & estimateStartDate & " 00:00:00' "
	End If

	'�������Ԏ�
	If ((IsNull(estimateEndDate) = True) Or (estimateEndDate = "")) Then
	ElseIf ((IsNull(estimateEndDate) = False) Or (estimateEndDate <> "")) Then
		dateTerm = dateTerm & " AND o1.���ϓ� <= '" & estimateEndDate & " 23:59:59' "
	End If

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

	vSQL = vSQL & "  FROM �� AS o1 WITH (NOLOCK) "
	vSQL = vSQL & "      LEFT JOIN �󒍃J�[�h��� AS cc WITH (NOLOCK) "
	vSQL = vSQL & "        ON cc.�󒍔ԍ� = o1.�󒍔ԍ� "

'	vSQL = vSQL & "  WHERE o1.�ڋq�ԍ� = " & wCustomerNo & " "
'	vSQL = vSQL & "    AND o1.�󒍌`�� IN (" & orderType & ") "
	vSQL = vSQL & "  WHERE 1=1 "

	' �ڋq�ԍ�
	If (IsNull(wCustomerNo) = False) And (wCustomerNo <> "") Then
		vSQL = vSQL & " AND o1.�ڋq�ԍ� = " & wCustomerNo
	End If

	' �󒍔ԍ�
	If (IsNull(wOrderNo) = False) And (wOrderNo <> "") Then
		vSQL = vSQL & " AND o1.�󒍔ԍ� = " & wOrderNo
	End If

	vSQL = vSQL & "    AND o1.�󒍌`�� IN (" & orderType & ") "
	vSQL = vSQL & "    AND o1.�x�����@ = '�N���W�b�g�J�[�h' "
'	vSQL = vSQL & "    AND o1.�x�����@�ڍ� = '03' " ' 03�͓X�ܗp
	vSQL = vSQL & "    AND o1.�x�����@�ڍ� = '05' "


	'PayPay���ϔԍ�
	'If (IsNull(wPaypayPaymentId) = False) Or (wPaypayPaymentId <> "") Then
	If (wPaypayPaymentId <> "") Then
		vSQL = vSQL & " AND cc.�J�[�h�^�M�m�F�ԍ� = '" & wPaypayPaymentId & "' "  ' card_credit_no
	End If

	vSQL = vSQL & dateTerm

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
	vSQL = vSQL & "  o.* "
	vSQL = vSQL & "  , (CASE WHEN o.�󒍓� IS NOT NULL AND o.�폜�� IS NULL THEN 'Y' ELSE 'N' END ) AS ���c "
	vSQL = vSQL & "FROM ("
	vSQL = vSQL & "  SELECT "
	vSQL = vSQL & "    * "
	vSQL = vSQL & "  FROM ("
	vSQL = vSQL & "    SELECT "
	vSQL = vSQL & "      ROW_NUMBER() OVER(ORDER BY o2.���ϓ� DESC) AS RN "
	vSQL = vSQL & "      ,o2.* "
	vSQL = vSQL & "    FROM ("
	vSQL = vSQL & "      SELECT DISTINCT "
	vSQL = vSQL & "         o1.�󒍔ԍ� "
	vSQL = vSQL & "        , o1.�ڋq�ԍ� "
	vSQL = vSQL & "        , o1.�����Җ��O "
	vSQL = vSQL & "        , c.�ڋq�� "
	vSQL = vSQL & "        , o1.�󒍓� "
	vSQL = vSQL & "        , o1.���ϓ� "
	vSQL = vSQL & "        , o1.�폜�� "
	vSQL = vSQL & "        , o1.�o�׊����� "
	vSQL = vSQL & "        , o1.����ŗ� "
	vSQL = vSQL & "        , o1.�󒍌`�� "
	vSQL = vSQL & "        , o1.�x�����@ "
	vSQL = vSQL & "        , o1.�x�����@�ڍ� "
	vSQL = vSQL & "        , o1.�󒍍��v���z "
	vSQL = vSQL & "        , o1.���v���z "
'2020.06.01 GV add start
	vSQL = vSQL & "        , o1.�󒍎����� "
	vSQL = vSQL & "        , o1.�󒍎�����萔�� "
	vSQL = vSQL & "        , o1.�󒍎��ߕs�����E���z "
	vSQL = vSQL & "        , o1.�󒍎����p�|�C���g "
	vSQL = vSQL & "        , o1.�󒍎����v���z "
'2020.06.01 GV add end
	vSQL = vSQL & "        , o1.���p�|�C���g "
	vSQL = vSQL & "        , o1.Web�����ύX�L�����Z�����t���O "
	vSQL = vSQL & "        , o1.���̑����v���z "
	vSQL = vSQL & "        , o1.���������t���O "
'	vSQL = vSQL & "        ,cc.�J�[�h�x�����z "			' card_total_amount
	vSQL = vSQL & "        ,cc.�J�[�h�^�M�m�F�ԍ� "		' card_credit_no
	vSQL = vSQL & "        ,cc.�J�[�h�l�b�g�`�[�ԍ� "	' card_net_slip_no
'	vSQL = vSQL & "        ,(SELECT TOP 1 "
'	vSQL = vSQL & "            cc1.�J�[�h�^�M�m�F�ԍ� "
'	vSQL = vSQL & "          FROM �󒍃J�[�h��� AS cc1 WITH (NOLOCK) "
'	vSQL = vSQL & "          WHERE "
'	vSQL = vSQL & "            cc1.�󒍔ԍ� = o1.�󒍔ԍ� "
'	vSQL = vSQL & "         ) as �J�[�h�^�M�m�F�ԍ� "
'	vSQL = vSQL & "        ,(SELECT TOP 1 "
'	vSQL = vSQL & "            cc2.�J�[�h�l�b�g�`�[�ԍ� "
'	vSQL = vSQL & "          FROM �󒍃J�[�h��� AS cc2 WITH (NOLOCK) "
'	vSQL = vSQL & "          WHERE "
'	vSQL = vSQL & "            cc2.�󒍔ԍ� = o1.�󒍔ԍ� "
'	vSQL = vSQL & "         ) as �J�[�h�l�b�g�`�[�ԍ� "
' 2020.06.03 GV add start
	vSQL = vSQL & "        ,(SELECT "
'	vSQL = vSQL & "            sum(od1.�o�׎w�����v����)  "
	vSQL = vSQL & "            sum(od1.�o�׍��v����)  "
	vSQL = vSQL & "          FROM "
	vSQL = vSQL & "            �󒍖��� AS od1 WITH (NOLOCK) "
	vSQL = vSQL & "          WHERE "
	vSQL = vSQL & "            od1.�󒍔ԍ� = o1.�󒍔ԍ� "
'	vSQL = vSQL & "         ) as �o�׎w�����v���� "
	vSQL = vSQL & "         ) as �o�׍��v���� "
' 2020.06.03 GV add end
' 2021.01.04 GV add start
	vSQL = vSQL & "        ,(SELECT "
	vSQL = vSQL & "            �I�[�\�����󒍍��v���z  "
	vSQL = vSQL & "          FROM "
	vSQL = vSQL & "            �󒍃J�[�h��� AS oc1 WITH (NOLOCK) "
	vSQL = vSQL & "          WHERE "
	vSQL = vSQL & "            oc1.�󒍔ԍ� = o1.�󒍔ԍ� "
	vSQL = vSQL & "            AND oc1.�ύX�敪 = '�폜' "
	vSQL = vSQL & "            AND oc1.�I�[�\�����󒍍��v���z > 0 " '2021.01.12 GV add
	vSQL = vSQL & "         ) as �I�[�\�����󒍍��v���z "
' 2021.01.04 GV add end

	vSQL = vSQL & "      FROM "
	vSQL = vSQL & "        �� AS o1 WITH (NOLOCK)  "
	vSQL = vSQL & "      INNER JOIN �ڋq AS c WITH (NOLOCK) "
	vSQL = vSQL & "        ON c.�ڋq�ԍ� = o1.�ڋq�ԍ� "
	vSQL = vSQL & "      LEFT JOIN �󒍃J�[�h��� AS cc WITH (NOLOCK) "
	vSQL = vSQL & "        ON cc.�󒍔ԍ� = o1.�󒍔ԍ� "
	vSQL = vSQL & "      WHERE 1 = 1 "
'	vSQL = vSQL & "        o1.�ڋq�ԍ� = " & wCustomerNo

	' �ڋq�ԍ�
	If (IsNull(wCustomerNo) = False) And (wCustomerNo <> "") Then
		vSQL = vSQL & "     AND o1.�ڋq�ԍ� = " & wCustomerNo
	End If

	' �󒍔ԍ�
	If (IsNull(wOrderNo) = False) And (wOrderNo <> "") Then
		vSQL = vSQL & "     AND o1.�󒍔ԍ� = " & wOrderNo
	End If

	vSQL = vSQL & "        AND o1.�󒍌`�� IN (" & orderType & " ) "
'	vSQL = vSQL & "        AND o1.�x�����@ = '�N���W�b�g�J�[�h' AND o1.�x�����@�ڍ� = '03' " ' 03 �͓X�ܗp
	vSQL = vSQL & "        AND o1.�x�����@ = '�N���W�b�g�J�[�h' AND o1.�x�����@�ڍ� = '05' "

	' �󒍔ԍ�
'	If (IsNull(wOrderNo) = False) Or (wOrderNo <> "") Then
'		vSQL = vSQL & "        AND o1.�󒍔ԍ� = " & wOrderNo
'	End If

	'PayPay���ϔԍ�
	'If (IsNull(wPaypayPaymentId) = False) Or (wPaypayPaymentId <> "") Then
	If (wPaypayPaymentId <> "") Then
		vSQL = vSQL & "        AND cc.�J�[�h�^�M�m�F�ԍ� = '" & wPaypayPaymentId & "' "  ' card_credit_no
	End If

	vSQL = vSQL & dateTerm
	vSQL = vSQL & "    ) AS o2 "
	vSQL = vSQL & "  ) as o3 "
'	vSQL = vSQL & "  WHERE RN BETWEEN 1 AND 10 "
	vSQL = vSQL & "  WHERE RN BETWEEN " & ((PAGE_SIZE * (wIPage - 1)) + 1) & " AND " & (PAGE_SIZE * wIPage)
	vSQL = vSQL & ") AS o "
	vSQL = vSQL & "  ORDER BY "
	vSQL = vSQL & "  ���ϓ� DESC "
	vSQL = vSQL & "  ,�󒍔ԍ� desc "

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

			' �폜��
			If (IsNull(vRS("�폜��"))) Then
				deleteDate = ""
			Else
				deleteDate = CStr(Trim(vRS("�폜��")))
			End If

			'�����Җ��O
			If (IsNull(vRS("�����Җ��O"))) Then
				orderName = ""
			Else
				orderName = CStr(Trim(vRS("�����Җ��O")))
			End If

			'�ڋq��
			If (IsNull(vRS("�ڋq��"))) Then
				customerName = ""
			Else
				customerName = CStr(Trim(vRS("�ڋq��")))
			End If

			If (IsNull(vRS("���v���z"))) Then
				orderTotalAm2 = 0
			Else
				orderTotalAm2 = CDbl(vRS("���v���z"))
			End If

			If (IsNull(vRS("�󒍎��ߕs�����E���z"))) Then
				kabusokuAmAtOrder = 0
			Else
				kabusokuAmAtOrder = CDbl(vRS("�󒍎��ߕs�����E���z"))
			End If

			If (IsNull(vRS("�󒍎����p�|�C���g"))) Then
				usedPointAtOrder = 0
			Else
				usedPointAtOrder = CDbl(vRS("�󒍎����p�|�C���g"))
			End If

			If (IsNull(vRS("�󒍎����v���z"))) Then
				totalAmAtOrder = 0
			Else
				totalAmAtOrder = CDbl(vRS("�󒍎����v���z"))
			End If


			'20201.01.04 GV add start
			If (IsNull(vRS("�I�[�\�����󒍍��v���z"))) Then
				totalAmAtAuth = 0
			Else
				totalAmAtAuth = CDbl(vRS("�I�[�\�����󒍍��v���z"))
			End If
			'20201.01.04 GV add end

			' ���p�|�C���g
			If (IsNull(vRS("���p�|�C���g"))) Then
				usedPoint = 0
			Else
				usedPoint = CDbl(vRS("���p�|�C���g"))
			End If

			'���������t���O
			If (IsNull(vRS("���������t���O"))) Then
				depositFlag = ""
			Else
				depositFlag = CStr(Trim(vRS("���������t���O")))
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

			'�x�������@�ڍ�
			If (IsNull(vRS("�x�����@�ڍ�"))) Then
				wPaymentMethodDetail = ""
			Else
				wPaymentMethodDetail = CStr(vRS("�x�����@�ڍ�"))
			End If

			' �J�[�h�x�����z
			'If (IsNull(vRS("�J�[�h�x�����z"))) Then
			'	ccTotalAm = ""
			'Else
			'	ccTotalAm = CStr(CDbl(vRS("�J�[�h�x�����z")))
			'End If

			'�J�[�h�^�M�m�F�ԍ�
			If (IsNull(vRS("�J�[�h�^�M�m�F�ԍ�"))) Then
				ccCreditNo = ""
			Else
				ccCreditNo = CStr(Trim(vRS("�J�[�h�^�M�m�F�ԍ�")))
			End If

			'�J�[�h�l�b�g�`�[�ԍ�
			If (IsNull(vRS("�J�[�h�l�b�g�`�[�ԍ�"))) Then
				ccSlipNo = ""
			Else
				ccSlipNo = CStr(Trim(vRS("�J�[�h�l�b�g�`�[�ԍ�")))
			End If

			'�o�׍��v���� 2020.06.02 GV add
			If (IsNull(vRS("�o�׍��v����"))) Then
				shipSuuSum = 0
			Else
				shipSuuSum = CDbl(vRS("�o�׍��v����"))
			End If


			'--- ���׍s����
			With oJSON.data("list")
				.Add j ,oJSON.Collection()
				With .item(j)
					.Add "c_no", CStr(Trim(vRS("�ڋq�ԍ�")))
					.Add "o_no" ,CStr(Trim(vRS("�󒍔ԍ�")))
					.Add "o_dt" ,orderDate '�󒍓�
					.Add "est_dt" ,CStr(Trim(vRS("���ϓ�")))
					.Add "ship_comp_dt" , shippingCompDate  '�o�׊�����
					.Add "del_dt" ,deleteDate '�폜��
					.Add "o_nm" ,orderName '�����Җ��O
					.Add "cst_nm" ,customerName '�ڋq��
					.Add "o_type" ,CStr(Trim(vRS("�󒍌`��")))
					.Add "pay_method" ,get_paymetMethodWord(vRS("�x�����@"))
					.Add "pay_method_detail" ,wPaymentMethodDetail '�󒍕��@����
					.Add "total_order_am", CDbl(vRS("�󒍍��v���z")) 
					.Add "total_order_am2",  orderTotalAm2  ' ���v���z
					.Add "ff_charge_o", CDbl(vRS("�󒍎�����"))  '2020.06.01 GV add
					.Add "cod_charge_o", CDbl(vRS("�󒍎�����萔��")) '2020.06.01 GV add
					.Add "kabusoku_am_o", kabusokuAmAtOrder '2020.06.01 GV add
					.Add "used_pt_o", usedPointAtOrder  ' �󒍎����p�|�C���g  '2020.06.01 GV add
					.Add "total_am_o", totalAmAtOrder '�󒍎����v���z '2020.06.01 GV add
					.Add "used_pt", usedPoint  ' ���p�|�C���g
					.Add "cc_c_no", ccCreditNo '�J�[�h�^�M�m�F�ԍ�(card_credit_no)
					.Add "cc_slip", ccSlipNo '�J�[�h�l�b�g�`�[�ԍ�(card_net_slip_no)
					.Add "o_zan", CStr(Trim(vRS("���c"))) 
					.Add "tax_rate", CDbl(vRS("����ŗ�"))
					.Add "ship_suu_sum", shipSuuSum '�o�׍��v����
					.Add "deposit", depositFlag '���������t���O 2016.06.03 GV add
					.Add "total_am_auth", totalAmAtAuth '�I�[�\�����󒍍��v���z 2021.01.04 GV add
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
