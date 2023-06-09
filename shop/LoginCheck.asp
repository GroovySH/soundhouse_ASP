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
<!--#include file="../common/HttpsSecurity.inc"-->
<%
'========================================================================
'
'	�I�[�_�[���O�C�����`�F�b�N
'
'	�X�V����
'2004/12/20 NaviLeft���O�C���{�^������̌Ăяo�����͌Ăяo�����Ƃւ��ǂ�B
'2005/06/06 SQL�C���W�F�N�V�����΍�
'2006/08/08 ���̓f�[�^�`�F�b�N����
'2006/09/18 LoginFl�ǉ��@���O�C�����ێ�
'2007/03/22 �p�X���[�h��Trim�ǉ��@�i�X�y�[�X���͎��̑Ώ�)
'2007/08/07 Web�s�f�ڃt���O NOT= Y�@�̃f�[�^�̂ݍ̗p
'2008/04/01 �p�X���[�h�n�b�V��
'2008/04/05 �p�X���[�h���Z�b�g��1�x�����Ă��Ȃ��l�́A�����I�Ƀ��Z�b�g�y�[�W��
'2008/04/14 ���O�C�������쐬�A3�񑱂��Ď��s�����烍�b�N
'2008/05/13 ���O�C����HTTPS�p�Z�b�V����ID���Z�b�g
'2008/05/14 HTTPS�`�F�b�N�Ή�
'2009/04/30 �G���[����error.asp�ֈړ�
'2009/05/29 �p�X���[�h�n�b�V����SHA-256�ɕύX
'2010/03/17 hn FeedBack.asp�ւ̌Ăяo���̖߂�ǉ�
'2010/07/28 an ���O�C�����O��SessionID, �ڋq�ԍ��ǉ�
'2010/07/30 st RtnURL������ꍇ�͂��̂܂܌Ăяo�����փ��_�C���N�g
'2011/02/21 hn RtnURL�̃`�F�b�N�����iPCIDSS)
'2011/04/14 hn SessionID �֘A�`�F�b�N
'2011/04/20 an #843 ���O�C�����AEmail�̑���Ƀ��[�U�[ID���g�p/���O�C�����������ʊ֐���
'2011/06/13 hn ���ꃆ�[�U�[ID���L��΃G���[
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/10/01 an #722 ���O�C�����Ɍڋq����Cookie�ɃZ�b�g
'2011/12/19 hn DC�Ή�
'2011/12/25 hn �}�C�y�[�W�Ή�  member.asp�@�� mypage.asp
'2012/01/17 hn cookie��Domain�����ǉ�
'2012/02/15 GV Cookie �� LIFL �� ULIFL �ɃL�[���ύX
'2012/03/07 GV #1234 �G���[���O�o�͂����ʃv���V�[�W����p���ďo�͂���悤�ύX
'2012/03/08 GV #1234 �G���[�`�F�b�N�� �o�̓��b�Z�[�W�̕�����ύX
'2012/03/26 GV #1254 ���[�U�[ID�ƃp�X���[�h������̂��q�l�̓��O�C���o���Ȃ��悤�ɕύX
'2012/03/26 GV #1254 ���[�U�[ID�Ɉ�v����ڋq��񂠂�A�p�X���[�h�Ⴂ�̍ۂ̃��b�Z�[�W�ύX
'2012/03/26 GV �ߋ��̕s�v�ȃR�����g�A�E�g��������уR�����g���폜 (2011/8/1�ȑO��)
'2012/09/07 nt �E�B�b�V�����X�g����̃��_�C���N�g��ǉ�
'
'========================================================================
On Error Resume Next

Response.Expires = -1			' Do not cache

Dim userID
Dim MemberID
Dim LoginFl
Dim LoginCount
Dim member_email
Dim member_password
Dim called_from
Dim RtnURL

Dim Connection
Dim RS_customer
Dim RS_order_header

DIm wPasswordResetFl

Dim w_sql
Dim w_msg
Dim w_html

Dim w_userID
Dim w_userName
Dim wMemberEmail
Dim wErrDesc   '2011/08/01 an add

' CAPICOM's hash algorithm constants.
Const CAPICOM_HASH_ALGORITHM_SHA1      = 0
Const CAPICOM_HASH_ALGORITHM_MD2       = 1
Const CAPICOM_HASH_ALGORITHM_MD4       = 2
Const CAPICOM_HASH_ALGORITHM_MD5       = 3
Const CAPICOM_HASH_ALGORITHM_SHA256    = 4
Const CAPICOM_HASH_ALGORITHM_SHA384    = 5
Const CAPICOM_HASH_ALGORITHM_SHA512    = 6

'=======================================================================

userID = Session("userID")
LoginFl = Session("LoginFl")
LoginCount = Session("LoginCount")

If IsNumeric(LoginCount) = False Then
	LoginCount = 0
Else
	LoginCount = CLng(LoginCount)
End If

'---- ���̓f�[�^�[�̎��o��
MemberID = ReplaceInput_NoCRLF(Trim(Request("MemberID")))
member_password = ReplaceInput(Trim(Request("member_password")))
called_from = ReplaceInput(Request("called_from"))
RtnURL = replace(ReplaceInput(Request("RtnURL")), "��", "&")

'---- RtnURL���s���ȏꍇ�̓G���[ @@@�b��Ή� �A�J�}�C���ʑ҂�@@@ 2011/02/28 hn add
If RtnURL <> "" Then
	If InStr(LCase(RtnURL), LCase(g_HTTP)) <> 1 _
	And InStr(LCase(RtnURL), LCase(g_HTTPS)) <> 1 _
	And InStr(LCase(RtnURL), "http://hotplaza.soundhouse.co.jp") <> 1 _
	And InStr(LCase(RtnURL), "http://guide.soundhouse.co.jp") <> 1 Then
' 2012/03/07 GV Add Start
		' �G���[���O�o��
		Call fwriteErrorLog("���� RtnURL ���s�� (RtnURL=" & RtnURL & ")")
' 2012/03/07 GV Add End
		Response.Redirect g_HTTP & "shop/Error.asp"
	End If
End If

'---- ���C������
Session("msg") = ""
w_msg = ""

If userID = "" Or LoginFl <> "Y" Then
	Call connect_db()
	Call main()

	'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
	If Err.Description <> "" Then
' 2012/03/07 GV Mod Start
'		wErrDesc = "LoginCheck.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
		wErrDesc = Err.Description
		wErrDesc = Replace(wErrDesc, vbNewLine, " ")
		' �G���[���O�o��
		Call fwriteErrorLog("���O�C���`�F�b�N�����ŃG���[ " & wErrDesc)

		wErrDesc = "LoginCheck.asp" & " " & wErrDesc
' 2012/03/07 GV Mod End

		Call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
	End If                                           '2011/08/01 an add e

	Call close_db()
End If

If Err.Description <> "" Then
' 2012/03/07 GV Add Start
	' �G���[���O�o��
	wErrDesc = Err.Description
	Call fwriteErrorLog(wErrDesc)
' 2012/03/07 GV Add End
	Response.Redirect g_HTTP & "shop/Error.asp"
End If

If w_msg = "" Then

	If wPasswordResetFl = "Y" Then
		Response.Redirect g_HTTPS & "member/MemberPasswordResetRequest.asp?member_email=" & wMemberEmail
	Else
		If RtnURL <> "" Then
			Response.Redirect RtnURL
		Else
			Select Case called_from
			Case "order"
				Response.Redirect g_HTTPS & "shop/OrderInfoEnter.asp"
			Case "catalog"
				Response.Redirect g_HTTPS & "shop/CatalogRequest.asp"
			Case "present"
				Response.Redirect g_HTTPS & "shop/PresentOubo.asp"
			Case "feedback"
				Response.Redirect g_HTTPS & "shop/FeedBack.asp"
			Case "top"
				Response.Redirect g_HTTPS & "member/Mypage.asp?called_from=" & called_from
			'2012/09/07 nt add Start
			'---- �E�B�b�V�����X�g����̃��_�C���N�g��ǉ�
			Case "wishlist"
				Response.Redirect g_HTTPS & "shop/WishList.asp?called_from=" & called_from
			'2012/09/07 nt add End
			Case "navi"
				Response.Redirect g_HTTP			'Top��
			Case Else
				Response.Redirect g_HTTPS & "member/Mypage.asp?called_from=" & called_from
			End Select
		End If
	End If

Else

	If w_msg <> "NoData" Then
		Session("msg") = w_msg
	End If

	Response.Redirect g_HTTPS & "shop/Login.asp?called_from=" & called_from & "&RtnURL=" & RtnURL		'Login error at First Login

End If

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

End Function

'========================================================================
'
'	Function	Main
'
'========================================================================
'
Function main()

Dim HashedData
Dim vCustNo
Dim vHashPassword										' 2012/03/26 GV Add

Const LOGIN_FLAG_KEY = "ULIFL"							' 2012/02/15 GV Add

If MemberID = "" And member_password = "" Then
	w_msg = "NoData"
	Exit Function
End If

' 2012/03/08 GV Mod Start
'If MemberID = "" or member_password = "" Then
'	w_msg = "���͂��ꂽ���[�U�[ID�܂��́A�p�X���[�h������������܂���B"
'	Exit Function
'End If
If Len(MemberID) <= 0 Then
	w_msg = "���[�U�[ID����͂��ĉ������B"
	Exit Function
End If
If Len(member_password) <= 0 Then
	w_msg = "�p�X���[�h����͂��ĉ������B"
	Exit Function
End If
' 2012/03/08 GV Mod End

'---- �ڋq���`�F�b�N
w_sql = ""
w_sql = w_sql & "SELECT �ڋq�ԍ�"
w_sql = w_sql & "       , �ڋq��"
w_sql = w_sql & "       , �p�X���[�h2"
w_sql = w_sql & "       , Web�s�f�ڃt���O"
w_sql = w_sql & "       , ���Z�b�g�g�[�N��"
w_sql = w_sql & "       , ���Z�b�g�g�[�N���o�^��"
w_sql = w_sql & "       , �p�X���[�h���b�N��"
w_sql = w_sql & "       , �n�b�V���A���S���Y��"
w_sql = w_sql & "       , �ڋqE_mail1"
w_sql = w_sql & "  FROM Web�ڋq"
w_sql = w_sql & " WHERE ���[�U�[ID = '" & MemberID & "'"
w_sql = w_sql & "   AND Web�s�f�ڃt���O != 'Y'"
w_sql = w_sql & " ORDER BY �ڋq�ԍ� DESC"

Set RS_customer = Server.CreateObject("ADODB.Recordset")
RS_customer.Open w_sql, Connection, adOpenStatic, adLockOptimistic

If RS_customer.EOF = True Then
' 2012/03/08 GV Mod Start
'	w_msg = "���͂��ꂽ���[�U�[ID�܂��́A�p�X���[�h������������܂���B"
	w_msg = "���͂��ꂽ���[�U�[ID�̂��q�l��񂪂݂���܂���ł����B" _
	      & "<br>���o�^���ꂽ���[�U�[ID�����Y��̏ꍇ�́A���̃t�H�[����育�m�F�����܂��B"
' 2012/03/08 GV Mod End
	Call fInsertLoginHistory(MemberID, "���O�C��", "���s", gSessionID, "")
Else
	If IsNULL(RS_customer("�p�X���[�h���b�N��")) = False Then
' 2012/03/08 GV Mod Start
'		w_msg = "���̉�����̓��b�N����Ă��܂��<br>�p�X���[�h���Z�b�g���s���Ă��������B"
		w_msg = "���̂��q�l���̓��b�N����Ă��܂��B" _
		      & "<br>���b�N����������ɂ́A���̃t�H�[�����p�X���[�h�̍Đݒ���s���Ă��������B"
' 2012/03/08 GV Mod End
	Else

		'---- �d��������΃G���[
		If RS_customer.RecordCount > 1 Then
			w_msg = "�������[�U�[ID���o�^����Ă��܂��B<br>�\���󂠂�܂��񂪁A�V���ɉ���o�^�����肢�������܂��B"
		Else

			'--- �I�u�W�F�N�g�쐬
			Set HashedData = CreateObject("CAPICOM.HashedData")
			'--- �A���S���Y����SHA1���w��
			If RS_customer("�n�b�V���A���S���Y��") = "SHA1" or RS_customer("�n�b�V���A���S���Y��") = "" Then
				HashedData.Algorithm = CAPICOM_HASH_ALGORITHM_SHA1
			End If
			'--- �A���S���Y����SHA-256 ���w��
			If RS_customer("�n�b�V���A���S���Y��") = "SHA-256" Then
				HashedData.Algorithm = CAPICOM_HASH_ALGORITHM_SHA256
			End If
			'--- �n�b�V���l���v�Z
			HashedData.Hash member_password

' 2012/03/26 GV Add Start
			vHashPassword = HashedData.Value

			Set HashedData = Nothing
' 2012/03/26 GV Add End

' 2012/03/26 GV Mod Start
'			If RS_customer("�p�X���[�h2") <> HashedData.Value Then
			If RS_customer("�p�X���[�h2") <> vHashPassword Then
' 2012/03/26 GV Mod End
' 2012/03/08 GV Mod Start
'				w_msg = "���͂��ꂽ���[�U�[ID�܂��́A�p�X���[�h������������܂���B"
' 2012/03/26 GV Mod Start
'				w_msg = "���͂��ꂽ���[�U�[ID�A�܂��́A�p�X���[�h�̂��q�l��񂪂݂���܂���ł����B" _
'				      & "<br>���o�^���ꂽ���[�U�[ID�A�܂��̓p�X���[�h�����Y��̏ꍇ�́A���̃t�H�[����育�m�F�����܂��B"
				w_msg = "���͂��ꂽ�p�X���[�h�̂��q�l��񂪂݂���܂���ł����B" _
				      & "<br>�p�X���[�h�����Y��̏ꍇ�́A���̃t�H�[����育�m�F�����܂��B"
' 2012/03/26 GV Mod End
' 2012/03/08 GV Mod End
			Else
				If RS_customer("Web�s�f�ڃt���O") = "Y" Then
					w_msg = "���̉�����͍폜����Ă��܂��<br>�ēo�^�Ȃǂ̂��⍇����<a href='" & g_HTTP & "shop/Inquiry.asp'>������</a>"
				Else

' 2012/03/26 GV Add Start
					If LCase(MemberID) = LCase(member_password) Then

						' ���O�C��ID�ƃp�X���[�h������̏ꍇ�A���O�C���s��(�p�X���[�h�ύX�𑣂�)
						w_msg = "���͂��ꂽ���[�U�[ID�ƃp�X���[�h�������ł��B" _
						      & "<br>�ύX�����肢�������܂��B"

					Else
' 2012/03/26 GV Add End

						If (IsNull(RS_customer("���Z�b�g�g�[�N���o�^��")) = True) Or _
						   (IsNull(RS_customer("���Z�b�g�g�[�N���o�^��")) = False And RS_customer("���Z�b�g�g�[�N��") <> "") Then
							wPasswordResetFl = "Y"
							wMemberEmail = RS_customer("�ڋqE_mail1")
						Else
							Session("userID") = RS_customer("�ڋq�ԍ�")		'OK�̂Ƃ��ڋq�ԍ����Z�b�g
							Session("userName") = RS_customer("�ڋq��")		'OK�̂Ƃ��ڋq�����Z�b�g
							Session("LoginFl") = "Y"		'���O�C����������Y
							Session("LoginCount") = 0		'���O�C����������0

							'---- Cookie�Ɍڋq���Z�b�g										'2011/10/01 an add s
							Response.Cookies("CustName") = RS_customer("�ڋq��")			'2011/10/01 an add e
							Response.Cookies("CustName").Domain = gCookieDomain				'2012/01/17 hn add

							'---- Cookie�Ƀ��O�C���t���O�Z�b�g								'2011/12/19 hn add
' 2012/02/15 GV Mod Start
'							Response.Cookies("LIfl") = "Y"
'							Response.Cookies("LIfl").Domain = gCookieDomain					'2012/01/17 hn add
							Response.Cookies(LOGIN_FLAG_KEY) = "Y"
							Response.Cookies(LOGIN_FLAG_KEY).Domain = gCookieDomain
' 2012/02/15 GV Mod End

							'---- �Z�b�V�����f�[�^�Ɍڋq�ԍ��Z�b�g       '2011/12/19 hn add
							vCustNo = RS_customer("�ڋq�ԍ�")
							Call fSetSessionData(gSessionID, "�ڋq�ԍ�", vCustNo)

							'---- HTTPS�p�Z�b�V����ID�Z�b�g
							'Call SetSSID()                  '2011/10/01 an del
						End If

' 2012/03/26 GV Add Start
					End If
' 2012/03/26 GV Add End

				End If
			End If
		End If
	End If

	If w_msg <> "" Then

		Session("LoginCount") = LoginCount + 1

		If Session("LoginCount") >= 5 Then

			RS_customer("�p�X���[�h���b�N��") = Now()
			RS_customer.update

			' ���O�C�������쐬
			Call fInsertLoginHistory(MemberID, "�p�X���[�h���b�N", "--", gSessionID, RS_customer("�ڋq�ԍ�"))

		Else
			' ���O�C�������쐬
			Call fInsertLoginHistory(MemberID, "���O�C��", "���s", gSessionID, RS_customer("�ڋq�ԍ�"))
		End If

	Else

		' ���O�C�������쐬
		Call fInsertLoginHistory(MemberID, "���O�C��", "����", gSessionID, RS_customer("�ڋq�ԍ�"))

	End If

End If

RS_customer.Close

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
%>
