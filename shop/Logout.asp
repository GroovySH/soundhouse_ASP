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
'	���O�A�E�g
'
'	�X�V����
'
'2005/10/14 TOP�֖߂�P�[�X��inquirysend.asp, catalogrequeststore��ǉ�
'2006/09/18 LoginFl�ǉ�
'2007/12/21 English�y�[�W�ǉ��ɔ����߂肳��������t�@�C�����݂̂ōs��
'2011/10/01 an #722 Cookie�폜
'2011/12/19 hn DC�Ή�
'2012/01/18 hn Cookie�Ƀh���C�������ǉ�
'2012/02/15 GV Cookie �� LIFL �� ULIFL �ɃL�[���ύX
'2012/08/14 GV #1419 �����O�C�����E�B�b�V�����X�g���烍�O�C����ʂ�\������
'
'========================================================================

'=======================================================================

Dim wHTTP_REFERER

Const LOGIN_FLAG_KEY = "ULIFL"	' 2012/02/15 GV Add

'---- ���C������
Session("userID") = ""
Session("userName") = ""
Session("userEmail") = ""
Session("LoginFl") = ""

Response.Cookies("CustName").Expires = DateAdd("d", -1, Now())	'2011/10/01 an add �L�������؂�ŏ㏑�����č폜
Response.Cookies("CustName").Domain = gCookieDomain				'2012/01/18 hn add
' 2012/02/15 GV Mod Start
'Response.Cookies("LIfl").Expires = DateAdd("d", -1, Now())		'2011/12/19 hn add ���O�C���t���O
'Response.Cookies("LIfl").Domain = gCookieDomain					'2012/01/18 hn add
Response.Cookies(LOGIN_FLAG_KEY).Expires = DateAdd("d", -1, Now())
Response.Cookies(LOGIN_FLAG_KEY).Domain = gCookieDomain
If Len(ReplaceInput(Request.Cookies("LIfl"))) > 0 Then
	' �Â� Cookie �� LIfl �����݂���ꍇ�A��������폜
	Response.Cookies("LIfl").Expires = DateAdd("d", -1, Now())
	Response.Cookies("LIfl").Domain = gCookieDomain
End If
' 2012/02/15 GV Mod End

Session.Abandon

wHTTP_REFERER = LCase(Request.ServerVariables("HTTP_REFERER"))

'---- ���L�Y��URL���烍�O�A�E�g���ꂽ�ꍇ��TOP�֖߂�
if InStr(wHTTP_REFERER, "orderinfoenter.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "orderconfirm.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "thanks.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "presentoubo.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "catalogrequest.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "inquirysend.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "catalogrequeststore.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
if InStr(wHTTP_REFERER, "/member") > 0 then
	wHTTP_REFERER = g_HTTP
end if
' 2012/08/14 GV #1419 Add Start
if InStr(wHTTP_REFERER, "wishlist.asp") > 0 then
	wHTTP_REFERER = g_HTTP
end if
' 2012/08/14 GV #1419 Add End

Response.Redirect wHTTP_REFERER		'�Ăяo�����Ƃւ��ǂ�

%>
