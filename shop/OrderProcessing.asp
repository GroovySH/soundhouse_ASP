<%@ LANGUAGE="VBScript" %>
<%
'�l�b�g�n�E�X�˂��ƃn�E�X�l�b�g�͂���
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
'	�I�[�_�[�������y�[�W
'
'2006/07/05 �J�[�h�I�[�\����BlueGate���Ăяo��
'2006/07/19 3D�Ăяo���ǉ�
'2006/10/24 �R���r�j���ώ�eContext�Ăяo���ǉ�
'2008/05/13 �N���X�T�C�g���N�G�X�g�t�H�W�F���[�΍� Key�p�����[�^�`�F�b�N
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2008/10/13 �V�J�[�h���͑Ή��@�i���̔F�؂ɖ߂�3D+Auth�j
'2010/03/15 hn �J�[�h�Ăяo���R�����g�A�E�g(Error.asp�ցj
'2012/02/15 GV �Z�L�����e�B�[�L�[�̓���`�F�b�N�����ʃv���V�[�W���Ŏ��s����悤�ύX (�Z�b�V�����ϐ���Skey�Ƃ̔�r���~�߁A�Z�b�V�����f�[�^�e�[�u������ Skey �Ƃ̔�r)
'2013/12/04 GV �o���f�[�V�����`�F�b�N��ǉ�
'2014/08/14 GV �����z�������ύX�Ή�
'
'========================================================================
'2013/12/04 GV add start ---
On Error Resume Next
Response.Expires = -1			' Do not cache
Response.buffer = true

Dim wMSG
Dim vRS
Dim wDateTime
Dim wDate

'---- �͐���
Dim wAfter13FL							'13���ȍ~
Dim wAfter14FL							'14���ȍ~
Dim wAfter15FL							'15���ȍ~
Dim wAfter16FL							'16���ȍ~
'Dim w94HokkaidouRitouFL				'�z���悪�@��B�E�l���E�k�C���E����   2011/06/29 an del
Dim w94ChugokuHokkaidouFL				'�z���悪�@��B�E�l���E�����E�k�C��   2011/06/29 an add
Dim wHolidayFL							'���j�Փ�
Dim wAvailableDate						'�w��\��
Dim wSatFl								'�y�j��
Dim wFriFL								'���j��
Dim wRitouFl							'�����t���O
Dim wShipPrefecture						'�͐�s���{��
Dim wDeliveryDt							'�w��[��
'2013/12/04 GV add end -----

Dim nextURL
Dim OrderTotalAm
Dim payment_method

OrderTotalAm = Trim(ReplaceInput(Request("OrderTotalAm")))
payment_method = Trim(ReplaceInput(Request("payment_method")))

'2013/12/04 GV add start ---
wDateTime = Now()
wDate = Date()
'wDateTime = "2013-12-06 1:10:30"
'wDate = "2013-12-06"

'---- DB
Dim Connection

Const w9Shuu4KokuChugokuHokkaido = "������,���茧,���ꌧ,�啪��,�F�{��,�{�茧,��������,���쌧,������,���Q��,���m��,���挧,���R��,������,�L����,�R����,�k�C��"   '2011/06/29 an add
Const cAddDaysToNyukaYoteibi = 2
'2013/12/04 GV add end -----

'---- �Z�L�����e�B�[�L�[�`�F�b�N
' 2012/02/15 GV Mod Start
'If Session("SKey") <> ReplaceInput(Request("SKey")) Then
If isLegalSecureKey(ReplaceInput(Request("SKey"))) = False Then
' 2012/02/15 GV Mod End
	Response.redirect "OrderInfoEnter.asp"
End If

'2013/12/04 GV add start ---
Call connect_db()

'---- ���̓f�[�^�[�̃`�F�b�N
Call validate_data()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "OrderProcessing.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc) 
end if                                           '2011/08/01 an add e

Call close_db()

If Err.Description <> "" Then
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

'---- �G���[�������Ƃ��͏������s�A�G���[������Β������e�w��y�[�W��
If wMSG = "" Then
Else
	Session("msg") = wMSG
	Server.Transfer "OrderInfoEnter.asp"
End If
'2013/12/04 GV add end -----

Session("BlueGate3DReturnCode") = ""		'BlueGate���^�[���R�[�h�N���A

If OrderTotalAm <> "0" Then
	If payment_method = "�N���W�b�g�J�[�h" Then
	''''	nextURL = "OrderCardAuthBG.asp"		'�I�[�\���̂ݎ擾
		''''nextURL = "OrderCard3DSecureBG.asp"		'3D+�I�[�\���擾
		''''Session("�󒍍��v���z") = OrderTotalAm
		''''nextURL = "OrderCard3DAuthSendBG.asp"		'�J�[�h����+3D+�I�[�\���擾
		''''nextURL = "OrderCard3DSecureBG2.asp"		'3D+�I�[�\���擾 NEW
		nextURL = "Error.asp"

	Else
		If payment_method = "�R���r�j�x��" Then
			nextURL = "OrderEcontext.asp"
		Else
			nextURL = "OrderSubmit.asp"
		End If
	End If
Else
	nextURL = "OrderSubmit.asp"
End If

'2013/12/04 GV add start ---
'========================================================================
'
'	Function	���̓f�[�^�[�̃`�F�b�N
'	OrderInfoInsert.asp �̓����֐��́A�z���������̂ݎ����B
'	�������ԑѕύX�̏ꍇ�́A���̃t�@�C�����ύX���邱�ƁB
'
'========================================================================
Function validate_data()

Dim vDateTimeMsg
Dim vSQL

'---- ����Recordset���o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    *"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ����"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set vRS = Server.CreateObject("ADODB.Recordset")
vRS.Open vSQL, Connection, adOpenStatic, adLockOptimistic

If vRS.EOF = True Then
	wMSG = "�J�[�g��񂪂���܂���B"
	Exit Function
End If

'�͐�s���{��
wShipPrefecture = vRS("�͐�s���{��")

'�w��[��
wDeliveryDt = vRS("�w��[��")

'---- �����t���O�̐ݒ�
Call setRitouFlag(vRS("�͐�X�֔ԍ�"))

'---- �z�B���w��(�����w�莞�̃`�F�b�N)
If vRS("�w��[��") <> "" Then

	'---- �z�B���w��\���`�F�b�N�i���c�i�œ��ח\�肪�Ȃ����A���ח\��+2�����O�͎w��NG�j
	If checkNyukaYoteibi() = False Then
		wMSG = wMSG & "�݌ɂ̂Ȃ����i���������Ɋ܂܂�Ă��邽�߁A�z�B��]���̎w��͂ł��܂���B<br>"
	Else
		'---- ���[���̏ꍇ�A�z�����w��s��
		If vRS("�x�����@") = "���[��" Then
			wMSG = wMSG & "���x�������@�����[���̏ꍇ�A�z�B��]���̎w��͂ł��܂���B<br>"
		Else
			wAfter13FL = False
			wAfter14FL = False
			wAfter15FL = False
			wAfter16FL = False
			'w94HokkaidouRitouFL = False    '2011/06/29 an del
			w94ChugokuHokkaidouFL = False   '2011/06/29 an add
			wHolidayFL = False
			wSatFL = False
			wFriFL = False

			'---- 13���ȍ~���ǂ����`�F�b�N
			If DatePart("h", wDateTime) >= 13 Then
				wAfter13FL = True
			End If

			'---- 14���ȍ~���ǂ����`�F�b�N
			If DatePart("h", wDateTime) >= 14 Then
				wAfter14FL = True
			End If

			'---- 15���ȍ~���ǂ����`�F�b�N
			If DatePart("h", wDateTime) >= 15 Then
				wAfter15FL = True
			End If

			'---- 16���ȍ~���ǂ����`�F�b�N
			If DatePart("h", wDateTime) >= 16 Then
				wAfter16FL = True
			End If

			'---- ���������j��Փ����ǂ����`�F�b�N
			If (DatePart("w",wDate) = vbSunday) Or (checkHoliday(Date()) = True) Then
				wHolidayFL = True
			End If

			'---- �������y�j�����ǂ����`�F�b�N
			If DatePart("w",wDate) = vbSaturday Then
				wSatFL = True
			End If

			'---- ���������j�����ǂ����`�F�b�N
			If DatePart("w",wDate) = vbFriday Then
				wFriFL = True
			End If

			'---- �z���悪�@��B�E�l���E�����E�k�C�����ǂ����`�F�b�N  
			'If wRitouFl = "Y" Or Instr(w9Shuu4KokuHokkaido, wShipPrefecture) > 0 Then
			if Instr(w9Shuu4KokuChugokuHokkaido, vRS("�͐�s���{��")) > 0 Then
				'w94HokkaidouRitouFL = True     '2011/06/29 an del
				w94ChugokuHokkaidouFL = True    '2011/06/29 an add
			End If

			'---- �w��\�����Z�b�g
			wAvailableDate = setAvailableDate()

			'---- �z�����`�F�b�N
			If (DateDiff("d", wDeliveryDt, wAvailableDate) > 0) Then

				'2011/02/16 hn mod s
				'---- �x��
				If wHolidayFl = True Then
					vDateTimeMsg = "�x����"
				Else
					'---- �R���r�j �y�j�� 13���ȍ~
					If payment_method = "�R���r�j�x��" AND wSatFl = True AND wAfter13Fl = true Then
						vDateTimeMsg = "�R���r�j�x���̓y�j��13���ȍ~��"
					End If

					'---- �R���r�j ���� 14���ȍ~
					If payment_method = "�R���r�j�x��" AND wSatFl = False AND wAfter14Fl = true Then
						vDateTimeMsg = "�R���r�j�x����14���ȍ~��"
					End If

					'---- ��s�U�� 14���ȍ~
					If payment_method = "��s�U��" AND wAfter14Fl = true Then
						vDateTimeMsg = "��s�U����14���ȍ~��"
					End If

					'---- ��s�U�� �y�j��
					If payment_method = "��s�U��" AND wSatFL = true Then
						vDateTimeMsg = "��s�U���̓y�j����"
					End If

					'---- ������E�J�[�h 16���ȍ~
					If (payment_method = "�����" OR payment_method = "�N���W�b�g�J�[�h") AND wAfter16Fl = true Then
						'vDateTimeMsg = payment_method & "��16���ȍ~��"	'2014/08/14 GV comment out
						vDateTimeMsg = payment_method & "��15���ȍ~��"	'2014/08/14 GV add
					End If
				End If

				'---- ��B��l��������E�k�C��
				'If w94HokkaidouRitouFL = True Then                             '2011/06/29 an del
					'wMSG = wMSG & "���͂��悪��B�E�l���E�k�C���E�����ŁA"     '2011/06/29 an del
				If w94ChugokuHokkaidouFL = True Then                            '2011/06/29 an add
					wMSG = wMSG & "���͂��悪��B�E�l���E�����E�k�C���ŁA"      '2011/06/29 an add
				End If
				
				'---- ����                      '2011/06/29 an add s
				If wRitouFl = "Y" Then
					wMSG = wMSG & "���͂��悪����E�����ŁA"
				End If                          '2011/06/29 an add e

				wMSG = wMSG & vDateTimeMsg & "�z�����w��́A" & wAvailableDate & "�ȍ~���w�肵�Ă��������B<br>"
				'2011/02/16 hn mod e

			End If

			'---- 60���ȓ�
			'2013/02/18 GV #1525 MOD START
'			If (DateDiff("d", DateAdd("d", 60,wDate), wDeliveryDt) > 0) Then
'				wMSG = wMSG & "�z�����w���60���ȓ��̓��t���w�肵�Ă��������B<br>"
			If (payment_method = "�����" AND DateDiff("d", DateAdd("d", 10,wDate), wDeliveryDt) > 0) Then
				wMSG = wMSG & "�z�����w���10���ȓ��̓��t���w�肵�Ă��������B<br>"
			ElseIf (payment_method <> "�����" AND DateDiff("d", DateAdd("d", 60,wDate), wDeliveryDt) > 0) Then
				wMSG = wMSG & "�z�����w���60���ȓ��̓��t���w�肵�Ă��������B<br>"
			'2013/02/18 GV #1525 MOD END
			End If
		End If
	End If
End If

If wMSG <> "" Then
	wMSG = "<b>�ȉ��̓��̓G���[��������Ă��������B</b><br><br>" & wMSG
End If

End Function

'========================================================================
'
'	Function	���c�̗L���Ɠ��ח\������m�F���A�z�B���w��\���`�F�b�N
'
'		return:	 �z�B���w��\�Ȃ� True
'                �z�B���w��s�Ȃ� False
'
'========================================================================
Function checkNyukaYoteibi()

Dim RSv
Dim vSQL
Dim vHikiateKanouQt
Dim vSetCount
Dim vMaxNyukaYoteibi

checkNyukaYoteibi = True

vHikiateKanouQt = ""
vSetCount = ""
vMaxNyukaYoteibi =""

'---- ���󒍖��׎��o��
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "     a.���[�J�[�R�[�h"
vSQL = vSQL & "   , a.���i�R�[�h"
vSQL = vSQL & "   , a.�F"
vSQL = vSQL & "   , a.�K�i"
vSQL = vSQL & "   , a.�󒍐���"
vSQL = vSQL & "   , b.�����\����"
vSQL = vSQL & "   , b.�����\���ח\���"
vSQL = vSQL & "   , c.�Z�b�g���i�t���O"
vSQL = vSQL & " FROM"
vSQL = vSQL & "     ���󒍖��� a WITH (NOLOCK)"
vSQL = vSQL & "   , Web�F�K�i�ʍ݌� b WITH (NOLOCK)"
vSQL = vSQL & "   , Web���i c WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "         b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "     AND b.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "     AND b.�F = a.�F"
vSQL = vSQL & "     AND b.�K�i = a.�K�i"
vSQL = vSQL & "     AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "     AND c.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "     AND b.�I���� IS NULL"
vSQL = vSQL & "     AND a.SessionID = '" & gSessionID & "'"		'2011/04/14 hn mod

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = False Then

	'---- �S���󒍖��׏��i���`�F�b�N
	Do Until RSv.EOF = True

		'�Z�b�g�i�̏ꍇ�̓Z�b�g�i�S�̂̍݌ɐ��ʁA���ח\������`�F�b�N����
		If RSv("�Z�b�g���i�t���O") = "Y" Then

			'---- �Z�b�g�i�݌ɐ��ʁAMAX�����\�\��擾
			Call GetSetCount(RSv("���[�J�[�R�[�h"), RSv("���i�R�[�h"), RSv("�F"), RSv("�K�i"), vSetCount, vMaxNyukaYoteibi)

			If vMaxNyukaYoteibi <wDate Then			'���ח\��̂Ȃ��Z�b�g�i�̏ꍇ��2000/01/01�������Ă���
				vMaxNyukaYoteibi = ""
			Else
				vMaxNyukaYoteibi = vMaxNyukaYoteibi		'�Z�b�g�i�S�̂�MAX���ח\����ŏ㏑��
			End If

			vHikiateKanouQt = vSetCount					'�Z�b�g�i�S�̂�MIN�݌ɐ��ʂŏ㏑��

		Else
			vHikiateKanouQt = RSv("�����\����")
			vMaxNyukaYoteibi = RSv("�����\���ח\���")
		End If

		If RSv("�󒍐���")  > vHikiateKanouQt Then		'���c�i�̏ꍇ�͓��ח\������`�F�b�N
			'---- ���ח\��̂Ȃ����i��1�ł�����ꍇ�͔z�B���w��s��
			If IsNULL(vMaxNyukaYoteibi) = True Or vMaxNyukaYoteibi = "" Then
				checkNyukaYoteibi = False
				Exit Do
			Else
				'---- �w��\���i���ח\���+2�ȍ~�j
				wAvailableDate = cf_FormatDate(DateAdd("d", cAddDaysToNyukaYoteibi, vMaxNyukaYoteibi), "YYYY/MM/DD")

				'---- ���ׂ��Ԃɍ���Ȃ����i��1�ł�����ꍇ�͔z�B���w��s��
				If wDeliveryDt < wAvailableDate Then
					checkNyukaYoteibi = False
					Exit Do
				End If
			End If
		End If

		RSv.MoveNext
	Loop
End If

RSv.Close

End Function

'========================================================================
'
'	Function	�Փ��`�F�b�N
'
'		input : �`�F�b�N�����(YYYY/MM/DD)
'		return:	�Փ��Ȃ�@True
'				�Փ��ȊO�@False
'
'========================================================================
Function checkHoliday(p_date)

Dim RSv
Dim vSQL

'---- �Փ��`�F�b�N
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    �N����"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    �J�����_�[ WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "        �N���� = '" & p_date & "'"
vSQL = vSQL & "    AND �x���t���O = 'Y'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	checkHoliday  = False
Else
	checkHoliday = True
End If

RSv.Close

End Function

'========================================================================
'
'	Function	�z���\���Z�b�g    2010/12/17 an mod  2011/02/16 hn mod
'
'		return:	�z���\�� (YYYY/MM/DD)
'
'   �x�����@���R���r�j
'     �y�j���ȊO
'     14���ȑO�F��t��������
'     14���ȍ~�F��t��������

'     �y�j��
'     13���ȑO�F��t��������
'     13���ȍ~�F��t��������
'
'   �x�����@����s�U���i�j���֌W�Ȃ��j
'     14���ȑO�F��t��������
'     14���ȍ~�F��t��������
'
'   �x�����@��������i�j���֌W�Ȃ��j
'     16���ȑO�F��t��������
'     16���ȍ~�F��t��������
'
'   �x�����@�����[��
'     �[���w��s��
'
'==============================
'   ��t�����A���j�A�j���A�i��s�U���̏ꍇ�͓y�j�����܂ށj�̏ꍇ�͎��̉c�Ɠ�����t���Ƃ���B
'
'   �[���w��\���́A
'   ��t��+1���@�i���L�ȊO�j
'   ��t��+2���@�i��B�A�l���A�����A�k�C���j
'   ��t��+5���@�i����A�����j
'
'========================================================================
'
Function setAvailableDate()

Dim vOrderDate

vOrderDate = cf_FormatDate(wDate, "YYYY/MM/DD")

'---- �R���r�j�x���F�y�j���ȊO�@14���ȍ~�͎�t���͗���
if vRS("�x�����@") = "�R���r�j�x��" AND wSatFl = false AND wAfter14Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1,wDate), "YYYY/MM/DD")
end if

'---- �R���r�j�x���F�y�j���@13���ȍ~�͎�t���͗���
if vRS("�x�����@") = "�R���r�j�x��" AND wSatFl = true AND wAfter13Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1,wDate), "YYYY/MM/DD")
end if

'---- ��s�U���F14���ȍ~�͎�t���͗���
if vRS("�x�����@") = "��s�U��" AND wAfter14Fl = true then
	vOrderDate = cf_FormatDate(DateAdd("d", 1,wDate), "YYYY/MM/DD")
end if

'2014/08/14 GV mod start
'15���ڍs�ɕύX
'---- ������E�J�[�h�F16���ȍ~�͎�t���͗���
'If (vRS("�x�����@") = "�����" OR vRS("�x�����@") = "�N���W�b�g�J�[�h") AND wAfter16Fl = true Then
If (vRS("�x�����@") = "�����" OR vRS("�x�����@") = "�N���W�b�g�J�[�h") AND wAfter15FL = true Then
	vOrderDate = cf_FormatDate(DateAdd("d", 1,wDate), "YYYY/MM/DD")
end if
'2014/08/14 GV mod end

'---- ��t�����A���j�A�j���A�i��s�U���̏ꍇ�͓y�j�����܂ށj�̏ꍇ�͎��̉c�Ɠ�����t���Ƃ���B
Do
	if  (DatePart("w", vOrderDate) = vbSunday) OR (checkHoliday(vOrderDate) = True) then
		vOrderDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
	else 
		if  vRS("�x�����@") = "��s�U��" AND (DatePart("w", vOrderDate) = vbSaturday) then
			vOrderDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
		else 
			exit do
		end if
	end if
Loop

'---- �w��\��
'if w94HokkaidouRitouFl = true then                '2011/06/29 an del
'---- �����F��t��+5��                             '2011/06/29 an mod s
if wRitouFl = "Y" then
	setAvailableDate = cf_FormatDate(DateAdd("d", 5, vOrderDate), "YYYY/MM/DD")
else
	'---- ��B��l��������E�k�C���F��t��+2��
	if w94ChugokuHokkaidouFL = true then
		setAvailableDate = cf_FormatDate(DateAdd("d", 2, vOrderDate), "YYYY/MM/DD")
	'---- ���̑��F��t��+1��                       '2011/06/29 an mod e
	else
		setAvailableDate = cf_FormatDate(DateAdd("d", 1, vOrderDate), "YYYY/MM/DD")
	end if
end if

End function

'========================================================================
'
'	Function	�����t���O�̐ݒ�
'
'		parm:		�z����X�֔ԍ�
'		return:	����E�����Ȃ�@wRitouFl = Y
'				�����ȊO�@�@�@�@wRitouFl = N
'
'========================================================================
Function setRitouFlag(p_zip)

Dim vZip
Dim RSv
Dim vSQL

vZip = Replace(p_zip, "-", "")

If vZip = "" Then
	wRitouFl  = "N"
	Exit Function
End If

'---- �����`�F�b�N
vSQL = ""
vSQL = vSQL & "SELECT"
vSQL = vSQL & "    �X�֔ԍ�"
vSQL = vSQL & " FROM"
vSQL = vSQL & "    ���� WITH (NOLOCK)"
vSQL = vSQL & " WHERE"
vSQL = vSQL & "    �X�֔ԍ� = '" & vZip & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF = True Then
	wRitouFl = "N"
	'---- ����͗�������               '2011/06/29 an add s
	if wShipPrefecture = "���ꌧ" then
		wRitouFl = "Y"                 '2011/06/29 an add e
	end if
Else
	wRitouFl = "Y"
End If

RSv.Close

End Function

'========================================================================
'
'	Function	Connect database
'
'========================================================================
Function connect_db()

'---- Connect database
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open g_connection

End function

'========================================================================
'
'	Function	Close database
'
'========================================================================
Function close_db()

Connection.Close
Set Connection= Nothing    '2011/08/01 an add

End Function

'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="refresh" content="0;URL=<%=nextURL%>">
<meta charset="Shift_JIS">
<title>��������t���b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/StyleOrder.css?20120629a" type="text/css">
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
      <li>���������e�̊m�F</li>
      <li class="now">��������t��</li>
    </ul>
  </div></div></div>

  <h1 class="title">��������t��</h1>
  <ol id="step">
    <li><img src="images/step01.gif" alt="1.�V���b�s���O�J�[�g" width="170" height="50"></li>
    <li><img src="images/step02.gif" alt="2.���͂���A���x�����@�̑I��" width="170" height="50"></li>
    <li><img src="images/step03_now.gif" alt="3.���������e�̊m�F" width="170" height="50"></li>
    <li><img src="images/step04.gif" alt="4.����������" width="170" height="50"></li>
  </ol>

  <p>�������̎�t�����Ă��܂��B<br>���΂炭���҂����������B</p>

<!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>
