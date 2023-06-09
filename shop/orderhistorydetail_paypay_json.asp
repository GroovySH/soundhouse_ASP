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
'	PayPay�x���ꗗ�y�[�W
'
'
'�ύX����
'2020.02.27 GV �V�K�쐬�B(PayPay�Ή�)(#2405)
'2020.06.01 GV �C���B
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

' 2020/.06.01 GV add start
Dim totalAmAtOrder ' �󒍎����v���z
Dim usedPointAtOrder


Dim wPaymentMethodDetail '2018.12.21 GV add

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

'�R���g���[���}�X�^���猩�ς���L���������擾 2018.01.12 GV add
call getEmaxCntlMst("��","�����m�F�҂�����","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)
If (IsNull(wItemNum1)) Then
	wDepositTerm = 10
Else
	wDepositTerm = wItemNum1
End If


' ���͒l������̏ꍇ
If (wFlg = True) Then
	'--- �Y���ڋq�̎󒍈ꗗ���o��
	vSQL = ""
	vSQL = vSQL & "SELECT "
	vSQL = vSQL & "  o.�ڋq�ԍ� "
	vSQL = vSQL & " ,o.�󒍔ԍ� "
	vSQL = vSQL & " ,o.�󒍓� "
	vSQL = vSQL & " ,o.���ϓ� "
	vSQL = vSQL & " ,o.�폜�� "
	vSQL = vSQL & " ,o.�o�׊����� "
	vSQL = vSQL & " ,o.�󒍌`�� "
	vSQL = vSQL & " ,o.�x�����@ "
	vSQL = vSQL & " ,o.����ŗ� "
	vSQL = vSQL & " ,o.�󒍍��v���z "
	vSQL = vSQL & " ,o.���v���z "
	vSQL = vSQL & " ,o.�󒍎����v���z " ' 2020.06.01 GV add
	vSQL = vSQL & " ,o.�󒍎����p�|�C���g " '2020.06.01 GV add
	vSQL = vSQL & " ,o.���p�|�C���g "
	vSQL = vSQL & " ,o.Web�����ύX�L�����Z�����t���O "
	vSQL = vSQL & " ,o.���̑����v���z "
	vSQL = vSQL & ",(SELECT "
	vSQL = vSQL & "    count(*) FROM �󒍂��̑����� other1 WITH(NOLOCK) "
	vSQL = vSQL & "  WHERE "
	vSQL = vSQL & "        other1.�󒍔ԍ� = o.�󒍔ԍ� " 
	vSQL = vSQL & "    AND �󒍂��̑��R�[�h <> 'COUPON' "
	vSQL = vSQL & ") as ���̑����׌��� "
	vSQL = vSQL & " ,ISNULL(o.�z����񖾍׎w��t���O, 'N') AS  �z����񖾍׎w��t���O "
	vSQL = vSQL & " ,o.���������t���O "
	vSQL = vSQL & " ,o.�̎����ԍ� "
	vSQL = vSQL & " ,o.�̎������s�� "
	vSQL = vSQL & " ,o.�x�����@�ڍ� "
	vSQL = vSQL & " , od.�������͐於�O "
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
	vSQL = vSQL & " ,CASE "
	vSQL = vSQL & "    WHEN i.B�i�t���O = 'Y' THEN iz.B�i�����\���� "
	vSQL = vSQL & "    ELSE iz.�����\���� "
	vSQL = vSQL & "	   END AS �݌ɐ� "
	vSQL = vSQL & " ,i.�戵���~�� "
	vSQL = vSQL & " ,i.�I���� "

	vSQL = vSQL & " , od.���[�J�[�����t���O "
	vSQL = vSQL & " , od.�󒍒P�� "
	vSQL = vSQL & " , od.�󒍋��z "
	vSQL = vSQL & " , od.�󒍐��� "
	vSQL = vSQL & " , od.�󒍎����� " '2020.06.01 GV add
	vSQL = vSQL & " , od.�o�׎w�����v���� "
	vSQL = vSQL & " , (CASE "
	vSQL = vSQL & "     WHEN o.�󒍓� IS NOT NULL AND o.�폜�� IS NULL THEN 'Y' "
	vSQL = vSQL & "     ELSE 'N' END "
	vSQL = vSQL & "   ) AS ���c "
	'vSQL = vSQL & " , iz.�K���݌ɐ��� "
	vSQL = vSQL & " , ISNULL(od.�K���݌ɐ���, 0) AS �K���݌ɐ��� "   '�������̓K���݌ɐ���
'	vSQL = vSQL & " ,od.�����ԍ� "
'	vSQL = vSQL & " ,od.�o�הԍ� "
'	vSQL = vSQL & " ,od.�o�א��� "
'	vSQL = vSQL & " ,od.�o�ד� "
'	vSQL = vSQL & " ,od.�^����ЃR�[�h "
	vSQL = vSQL & ", od.�󒍖��ה��l "


	vSQL = vSQL & " FROM "
	vSQL = vSQL & " �� AS o WITH (NOLOCK) "

	vSQL = vSQL & " INNER JOIN �󒍖��� od WITH (NOLOCK) "
	vSQL = vSQL & "    ON od.�󒍔ԍ� = o.�󒍔ԍ� "
	vSQL = vSQL & "   AND od.�Z�b�g�i�e���הԍ� = 0 "

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

	vSQL = vSQL & " WHERE "
	vSQL = vSQL & "       o.�ڋq�ԍ� =  " & wCustomerNo
	vSQL = vSQL & "   AND o.�󒍔ԍ� = " & wOrderNo

	vSQL = vSQL & " ORDER BY "
	vSQL = vSQL & "   �󒍖��הԍ� asc"


	'@@@@Response.Write(vSQL) & "<br>"


	Set vRS = Server.CreateObject("ADODB.Recordset")
	vRS.Open vSQL, ConnectionEmax, adOpenStatic, adLockOptimistic

	'���R�[�h�����݂��Ă���ꍇ
	If vRS.EOF = False Then

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

			If (IsNull(vRS("�󒍎����v���z"))) Then
				totalAmAtOrder = 0
			Else
				totalAmAtOrder = CDbl(vRS("�󒍎����v���z"))
			End If

			If (IsNull(vRS("�󒍎����p�|�C���g"))) Then
				usedPointAtOrder = 0
			Else
				usedPointAtOrder = CDbl(vRS("�󒍎����p�|�C���g"))
			End If

			' ���p�|�C���g
			If (IsNull(vRS("���p�|�C���g"))) Then
				usedPoint = 0
			Else
				usedPoint = CDbl(vRS("���p�|�C���g"))
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

			' ---------------------------------------------------


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
					.Add "total_am_o",  totalAmAtOrder  ' �󒍎����v���z 2020.06.01 GV add
					.Add "used_pt_o", usedPointAtOrder  ' �󒍎����p�|�C���g 2020.06.01 GV add
					.Add "used_pt", usedPoint  ' ���p�|�C���g
					.Add "ship_name" ,CStr(Trim(vRS("�������͐於�O")))
					.Add "web_flg", CStr(vRS("Web���i�t���O"))
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
					.Add "m_chokusou", makerChokusou
					.Add "i_tanka", CDbl(Trim(vRS("�󒍒P��")))
					.Add "i_am", CDbl(Trim(vRS("�󒍋��z")))
					.Add "i_suu", CDbl(vRS("�󒍐���")) 
					.Add "i_suu_o", CDbl(vRS("�󒍎�����")) 
					.Add "ship_inst_suu", CDbl(vRS("�o�׎w�����v����"))
					.Add "o_zan", CStr(Trim(vRS("���c"))) 
					.Add "t_zaiko_suu", CDbl(vRS("�K���݌ɐ���"))
					.Add "promote" , promote '�̑��i����
					.Add "modifying", webModCancelFlg
					.Add "deposit", depositFlag '���������t���O 2016.06.03 GV add
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