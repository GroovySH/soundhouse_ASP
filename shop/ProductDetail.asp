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
<!--#include file="../3rdParty/EAgency.inc"-->
<%
'========================================================================
'
'	���i�ڍ׃y�[�W
'�X�V����
'2004/12/16 hn ASK, ���̏ꍇ�̃|�b�v�A�b�v��ʂ𒆎~�������N�ɕύX�iMAC�Ŗ�肠��̂���)
'2004/12/17 hn ���ʓ��͗��ǉ�
'2005/01/05 hn �����N����Target�w���ύX
'2005/01/14 hn �p�[�c�A�I�v�V�����f�[�^���o�ɃJ�e�S���[�ʃe�[�u�����ǉ�
'2005/01/27 hn ASK�N���b�N��ASK�P���\����ʂ��|�b�v�A�b�v����悤�ɕύX
'              ����̏��i�Ɋւ���⍇����{�^���ǉ�
'2005/02/03 hn ���[�J�[��+���i���̕\����ύX ()/�̑O��ɔ��p�X�y�[�X�ǉ�
'2005/02/09 hn ����̏��i�̂��⍇�������̃����N���̃p�����[�^��URLEncode�K�p
'2005/02/16 �����萔�ʒP�����o�����̏��������@�����萔�ʁ�0��ǉ�
'2005/02/23 ���[�J�[���A�J�e�S���[����SearchList�փ����N�ǉ�
'2005/03/15 �֘A�V���[�Y���i�\���ǉ�
'2005/05/12 �֘A�L���C�֘A�j���[�X���֘A�L��1�|4�ɂ܂Ƃ߂�
'2005/05/13 ���[�J�[�֘A�����N��ǉ�
'2005/06/27 ���i���r���[��ǉ�
'2005/07/01 �[���\�������悻�̓����ɕύX
'2005/08/18 ����̏��i�̂��⍇����p�����[�^�𒆃J�e�S���[���{��ɕύX
'2005/10/18 Web�[����\���t���O�Ή�
'2005/11/01 �����EMovie�|�b�v�A�b�v�y�[�W�ւ̃����N�ǉ�
'2005/12/01 ASK���i�\���y�[�W�ւ̃p�����[�^��Server.URLEncode����
'2006/01/10 �����A����A�֘A�L���A���i�}�j���A���A�V���i�Љ�̃����N��http���܂܂�Ă���ꍇ�͊O�������N�Ƃ���B
'2006/03/07 �P���̐������C���[�W��
'2006/04/03 �p�������ǉ�
'2006/04/07 Web���i�t���O�`�F�b�N��ǉ�
'2006/04/21 ��p�@�킪�ݒ肳��Ă��鏤�i�͌�p�@�폤�i��\��
'2006/04/25 �戵���~�A�D�Ԃō݌ɖ������i�̂݌�p�@�킪�ݒ肳��Ă��鏤�i�͌�p�@�폤�i��\��
'2006/10/18 �݌ɏ󋵁@����裂�ǉ�
'2006/12/21 �F�K�i�ʍ݌ɂ̏I�����`�F�b�N��ǉ��i�F�K�i���菤�i�ŐF�K�i�Ȃ����\������Ă��܂����߁j
'2007/01/08 em�^�O�ǉ�
'2007/03/01 CreateSpecHTML�̃p�����[�^�ύX
'2007/03/02 �݌ɏ󋵁@�\���̐F��ύX
'2007/03/13 �VNAVI�ɕύX
'2007/04/05 ���O�C�����Ă���Ώ��i���r���[��o�^�ł���悤�ɕύX(�P��̂݁j
'2007/04/18 B�i��\���ǉ�
'2007/04/20 �w�����т����胍�O�C�����Ă���Ώ��i���r���[��o�^�ł���悤�ɕύX(�P��̂݁j
'2007/04/24 �o�גʒm����̏��i���r���[�Ώ�
'2007/04/25 ���i���r���[�̕ύX�i�ڋq�ʏ��i���r���[�ꗗ�{�^���A�n���h���l�[���j
'2007/05/08 ���i���l�C���T�[�gURL1,2 �ǉ�
'2007/05/14 ���i���l�C���T�[�g�T�C�Y�w�� �ǉ�
'2007/05/14 �p�Ԃō݌ɂ��菤�i�̍݌ɏ󋵂��u�݌Ɍ���v�Ƃ���
'2007/05/15 �J�[�g���e�ɑ����\���ǉ�
'2007/05/25 �V���b�v�R�����g�t���O�Ώ�
'2007/05/30	�F�K�i�����ŌĂяo����A�Y�����i���F�K�i����̏ꍇ�́A�F�K�i�I���h���b�v�_�E����\������B
'2007/06/05 �֘A�����N��ǉ�(NaviLeftShop����ړ��j
'2007/06/15 ���r���[�Ƀ����N�������̓`�F�b�N�ǉ�
'2007/06/25 ���i���r���[�ɒ��ӃR�����g�ǉ�
'2007/07/05 �F�K�i�I�����ɕύX���ʂ��؂�ւ��܂łɃJ�[�g�{�^���������ꂽ�Ƃ��Ⴄ�F�K�i�𑗐M����G���[�̑Ώ�
'2007/07/11 �F�K�i���i�摜�t�@�C����1-4�̍l��
'2007/07/17 �F�B�Ɋ��߂�A�E�B�b�V�����X�g�ɓ����{�^����ǉ�, �J�[�g�̒��g����������ʊ֐��ɕύX
'2007/08/23 ���i�A�N�Z�X�����o�^�i�y�[�W�r���[�j�@����Z�b�V�������ꏤ�i��1��
'2007/09/10 ���i�A�N�Z�X�����o�^��N���ʂɕύX
'2007/09/12 �Y�����i�Ȃ��̂Ƃ���SearchNotFound.asp��\��
'           ��p�@�킠��̂Ƃ��́ASearchList.asp�Ō�p�@��\��(i_type=successor)
'2007/10/22 ���r���[�������݃`�F�b�N���� WriteReview=Y�ł��w���񐔃`�F�b�N
'2007/11/20 �J�[�g�{�^���̉��֏��iID�\���ǉ�
'2007/12/13 �p��+�����\�݌Ɂ�1�̎��Ɂu�݌Ɍ���v�ƕ\��
'2007/12/27 �݌ɏ󋵁@�\���̐F��ύX
'2008/01/11 ������P����B�i�Ɠ��l�̒P���\���ɕύX
'2008/01/28 �F�K�i��1��ނ����Ȃ��ꍇ�̑Ή�
'2008/05/07 ���̓f�[�^�`�F�b�N����
'2008/05/21 ���r���[���e�`�F�b�N����EOF�`�F�b�N����
'2008/07/31 �F�K�i�ʍ݌�.�I���� IS NULL�f�[�^�̈����ύX
'2008/09/13 �o�גʒm�̃����N���烌�r���[�쐬�ŌĂ΂ꂽ���͎󒍔ԍ�����UserID�����o�����r���[���͉�ʂ�\������悤�ɕύX
'2008/09/16 (�ύX�˗�#503)�����萔�ʂ̕\�������̂悤�ɕύX
'						4�ȉ��@���s�ǂ���/5-9 ����5��/10-14�@����10��/15-19�@����15��/20�ȏ�A����20��
'
'2008/12/19 ���j���[�A�� ********
'2009/04/27 �F�K�i���m�肳��Ă��Ȃ��Ƃ��́A�݌ɏ󋵂��\��
'2009/05/27 ����A�����A���[�J�[�����N�A�}�j���A���@�A�C�R���ύX
'2009/08/06 ���i�\���T�C�Y��ύX(Style�w��j
'2009/10/26 ���i���l�C���T�[�g�T�C�YH1,2�̍����������폜
'2009/12/17 hn ���R�����h�p�ύX�i���i�A�N�Z�X���O�o�́A���R�����h�\���j
'2010/01/20 an �p�������X�g�̍ŉ��w�����N�ɃJ�e�S���[�R�[�h��ǉ����Ai_tyep=cm�ɕύX�B���J�e�S���[�ւ̃����N�������Ă��Ȃ������̂Œǉ�
'2010/01/26 hn �F�K�i�w�肠��ŁA1�����Ȃ��ꍇ�̕s����C��
'2010/01/29 an B�i�������󂠂�i�ɕ\�L��ύX
'2010/02/06 if-web ���i�\�����Ɂu�����F�v�ǉ�
'2010/02/22 st �󂠂�i���킯����i�ɕ\�L��ύX
'2010/03/04 an ���R�����h���i�A�N�Z�X���O�o�^��L����
'2010/03/08 hn ���r���[�͂��E���������摜�ɂ��AJavaScript�Ŏ��s�@�iBot�΍�j
'2010/04/06 an ���R�����h��ASK���i��ASK�\���ɏC��
'2010/04/21 an ���R�����h��ASK���i�̉��i�\���̊ԈႢ���C��
'2010/05/17 ko-web �����΍�̂���HTML�^�O�ih1,h2,h3,p,strong�j�ǉ�
'2010/06/10 an SEO�΍��<link>�^�O�ǉ�
'2010/07/01 an HTML���C�A�E�g�C��
'2010/08/23 an meta descripion,keywords�ɏ��i���������Z�b�g����悤�ɏC��
'2010/08/30 an Twitter�Ԃ₭�{�^���ǉ�
'2010/09/27 an Twitter�Ԃ₭�{�^���ʒu�ύX
'2010/11/04 GV(dy) #724 ���i�ڍׂ́u�݌ɏ󋵁v�� "�������" �̉摜���\������Ȃ����̂ݕ\������悤�ɏC��
'2010/11/10 an �֘A���i�Ɍ�������AB�i�����𔽉f����悤�ɏC���B���R�����h�A�֘A�V���[�Y��B�i�����𔽉f����悤�ɏC��
'2010/11/26 an �p�ԕi�ł��F�K�i���i�̍݌ɂ�����Ό�p�@���\�����Ȃ��悤�ɏC��
'2010/12/28 hn �p�[�c�I�v�V�������@wProdTermFl��vProdTermFl �ɕύX�@wProdTermFl���Ԃ�����
'2011/02/18 GV(dy) #826 �������S�����\���̑Ή�
'2011/03/18 GV(dy) #731 Style/StyleNaviLeftShop.css ��stylesheet��`��ǉ�
'2011/06/09 hn �p�Ԃō݌ɂȂ��{�����Ȃ��@�̎��Ɋ����Ƃ���悤�ɕύX
'2011/06/15 if-web �������S�����\���ɉ���E�����������|��ǋL
'2011/08/01 an #1087 Error.asp���O�o�͑Ή�
'2011/09/09 an #816 ���r���[�����e�i���X�Ή��Ń��r���[ID�\���ǉ�
'2011/10/19 hn 1063 ASK�\�����@�ύX
'2011/11/22 an #1150 ���b�`�X�j�y�b�g�Ή�, OGP�Ή�
'2012/01/10 an ���R�����h���i�A�N�Z�X���O�o�͒�~
'2012/01/18 GV �J�X�^�}�[���r���[�p�̃f�[�^�擾���A�u���e���v��ł̕��ёւ��� �uID�v��ł̕��ёւ��ɕύX
'2012/01/18 GV �֘A�V���[�Y���i�p, �I�v�V�����p����уp�[�c�p�̃f�[�^�擾 SELECT���� LAC�N�G���[�Ă�K�p (���킹�� WITH (NOLOCK) �t��)
'2012/01/23 GV �u���i���r���[�v�e�[�u������u���i���r���[�W�v�v�e�[�u���g�p�ɕύX (CreateReviewHTML()�v���V�[�W��)
'2012/02/20 na �u���i�A�N�Z�X���v���X�|���X�΍�̂��ߒ�~
'2012/04/10 GV ���R�����h�\�����C�A�E�g�ύX
'2012/05/01 GV ��֏��i�\���@�\�ǉ�
'2012/07/10 GV ���i�ڍ׃y�[�W�f�U�C���ύX
'2012/08/01 ok �E�T�C�h�̏��i���̃��[�J�[/���i/�J�e�S���[�������������N����ɕύX
'2012/08/27 ok �֘A�p�[�c�A�I�v�V�����̃J�e�S���[�\���Ή�
'2012/10/30 nt �֘A�R���e���c�\�����ڒǉ�
'2013/05/17 GV #1507 ���r���[�ҏW�@�\
'2013/05/22 GV #1505 ���Ԃ݂��ƁI���R�����h�Ή�
'2013/08/07 if-web �����R�����h�i�`�[�����{�j���R�����g�A�E�g
'2013/08/14 GV ��֏��i�擾���\�b�h �ŁADB����擾�����l���Ȃ��ꍇ�̏�����ǉ�
'2014/03/19 GV ����ő��łɔ���2�d�\���Ή�
'
'========================================================================

On Error Resume Next

Dim wUserID

Dim maker_cd
Dim product_cd
Dim iro
Dim kikaku
Dim item
Dim item_list()
Dim item_cnt

Dim ReviewAll
Dim WriteReview
Dim OrderNo
'Dim iShop				'2013/05/17 GV #1507 comment out

Dim wMakerName
Dim wMakerNameNoKana     '2010/08/30 an add
Dim wProductName
Dim wCategoryCode
Dim wTitleWithLink
Dim wKoukeiMakerCd
Dim wKoukeiProductCd
Dim wLargeCategoryCd
Dim wMidCategoryCd
Dim wCanWriteReviewFl
Dim wPrefecture
Dim wHandleName
Dim wIroKikakuSelectMsg
Dim wLargeCategoryName   '2010/08/23 an add
Dim wMidCategoryName     '2010/08/23 an add
Dim wCategoryName        '2010/08/23 an add
Dim wTokucho             '2010/08/23 an add
Dim wFreeShippingFlag		' 2011/02/18 GV Add
Dim s_category_cd        '2011/09/09 an add For NaviLeftShop
Dim wOptionPartsTitleFlag		'2012/08/29 ok Add

Dim wIroKikakuCombo

Dim wPictureHTML
Dim wKanrenLinkHTML
Dim wTokuchoHTML
Dim wSpecHTML
Dim wOptionHtml
Dim wPartsHtml
Dim wReviewHTML

Dim wProductHTML
Dim wHyoukaHTML
Dim wCartHTML
Dim wSeriesHTML
Dim wRecommendHTML
Dim wRecommendBuyHTML	' 2012/04/10 GV Add
DIm wSubItemHTML        ' 2012/05/01 GV Add
Dim wViewHTML		' 2012/07/10 GV Add

Dim Connection
Dim RS

Dim wTitle
Dim wSalesTaxRate
Dim wProdTermFl
Dim wPrice
Dim wTaxedPrice			'2014/03/19 GV add
Dim wOptionPartsFl
Dim wIroKikakuSelectedFl
Dim wNoData
Dim wIroKikakuFl       '2010/11/26 an add
Dim wIroKikakuZaikoFl  '2010/11/26 an add
Dim wIroKikakuHacchuuFl  '2011/06/09 an add
Dim wMainProdPic        '2011/11/22 an add

Dim wItemChar1
Dim wItemChar2
Dim wItemNum1
Dim wItemNum2
Dim wItemDate1
Dim wItemDate2

Dim wNum		'�����摜

Dim wSQL
Dim wHTML
Dim wMSG
Dim wErrDesc   '2011/08/01 an add

Dim wSeriesCd		'2012/10/30 nt Add
Dim wContentsHTML	'2012/10/30 nt Add

Dim wTwPriceLabel
Dim wTwPriceData
Dim wTwInventoryData

'2013/05/17 GV #1505 add start
Dim wRecommendJS
Dim wRecommendBuyJS
Dim wEAProductDetailData
Dim wEAInventoryData
Dim wEAPrice
Dim wEAPriceExcTax
Dim wEAIroKikakuData
'2013/05/17 GV #1505 add end

'========================================================================

Response.buffer = true

wUserID = Session("UserID")

'---- Get input data
maker_cd = ReplaceInput(Trim(Request("maker_cd")))
product_cd = ReplaceInput(Trim(Request("product_cd")))
iro = ReplaceInput(Trim(Request("iro")))
kikaku = ReplaceInput(Trim(Request("kikaku")))
item = ReplaceInput(Trim(Request("item")))

if Trim(Request("parm")) <> "" then
	item = ReplaceInput(Trim(Request("parm")))
end if

iro = ""
kikaku = ""

if item <> "" then
	item_cnt = cf_unstring(item, item_list, "^")
	maker_cd = item_list(0)
	product_cd = item_list(1)
	if item_cnt > 2 then
		iro = item_list(2)
		if item_cnt > 3 then
			kikaku = item_list(3)
		end if
	end if
end if

'----���i���r���[�p�p�����[�^
ReviewAll = ReplaceInput(Request("ReviewAll"))
WriteReview = ReplaceInput(UCase(Request("WriteReview")))

'2013/07/10 GV #1507 add start
'�������N�ŃA�N�Z�X���Ă����ꍇ�A���_�C���N�g
If (WriteReview = "Y") Then
	If  gPhoneType = "SP" Then
		Response.Redirect g_HTTPS & "sp/shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "sp/shop/ReviewWrite.asp?Item=" & Server.URLEncode(item)
	Else
		Response.Redirect g_HTTPS & "shop/LoginCheck.asp?RtnURL=" & g_HTTPS & "shop/ReviewWrite.asp?Item=" & Server.URLEncode(item)
	End If
End If
'2013/07/10 GV #1507 add end

OrderNo = ReplaceInput(Request("OrderNo"))
if (OrderNo <> "" and isNumeric(OrderNo) = false) OR OrderNo = "" then
	OrderNo = 0
end if

'iShop = ReplaceInput(Trim(Request("iShop")))		'2013/05/17 GV #1507 comment out

'2013/05/22 GV #1505 add start
wRecommendJS = ""
wRecommendBuyJS = ""
wEAInventoryData = ""
wEAPrice = 0
wEAPriceExcTax = 0
'2013/05/22 GV #1505 add start

'---- Execute main
call connect_db()
call main()

'---- �G���[���b�Z�[�W���Z�b�V�����f�[�^�ɓo�^   '2011/08/01 an add s
if Err.Description <> "" then
	wErrDesc = "ProductDetail.asp" & " " & replace(replace(Err.Description, vbCR, " "), vbLF, " ")
	call fSetSessionData(gSessionID, "ErrDesc", wErrDesc)
end if                                           '2011/08/01 an add e

if Err.Description <> "" then
	Response.Redirect g_HTTP & "shop/Error.asp"
end if

'---- �Y�����i�Ȃ��̂Ƃ�
if wNoData = "Y" then
	Response.Redirect "SearchNotFound.asp"
end if

'---- ��p�@�킪����ꍇ�͂��̏��i��\��
if wKoukeiMakerCd <> "" then
	Response.Redirect "SearchList.asp?i_type=successor&s_maker_cd=" & wKoukeiMakerCd & "&s_product_cd=" & Server.URLEncode(wKoukeiProductCd)
end if

'========================================================================
'
'	Function	Connect database
'
'========================================================================
'
Function connect_db()
Dim i

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

Dim vPointer

'---- ����ŗ���o��
call getCntlMst("����","����ŗ�","1", wItemChar1, wItemChar2, wItemNum1, wItemNum2, wItemDate1, wItemDate2)			'����ŗ�
wSalesTaxRate = Clng(wItemNum1)

'---- ���i�����o��
call GetProduct()
if wMSG <> "" OR wKoukeiMakerCd <> "" OR wNoData = "Y" then
	exit function
end if

'---- ���i�摜
call CreatePictureHTML()

'---- �֘A�����N
Call CreateKanrenLinkHTML()

'---- ����
call CreateTokuchoHTML()

'---- �X�y�b�N
call CreateSpecificationHTML()

'----- �I�v�V����HTML�쐬
wOptionPartsFl = false
call CreateOptionHTML()

'---- �p�[�cHTML�쐬
call CreatePartsHTML()

wCanWriteReviewFl = "N"

'---- �J�X�^�}�[���r���[�A�]��HTML�쐬
if RS("B�i�t���O") <> "Y" then
	call CreateReviewHTML()

	'---- �o�גʒm����̃����N�Ŏ󒍔ԍ����n���ꂽ�ꍇ�́AUserID���o��
	if OrderNo <> 0 AND wUserID = "" then
		call GetUserID()
	end if

	'---- ���i���r���[�o�^�σ`�F�b�N
	if wProdTermFl <> "Y" AND wUserID <> "" then
		call CheckReview()
	end if
end if

'==== ��������E�� ================================
'---- ���[�J�[/���iHTML�쐬
call CreateProductHTML()

'---- �]��HTML�쐬
'CreateReview�ňꏏ�ɍ쐬��

'----- �J�[�g���HTML�쐬�i���ʊ֐��j
wCartHTML = fCreateCartHtml()

'----- ���i���HTML�쐬�E�\�� GV 2012/05/01
Call GetSubstituteItem()

'----- ���R�����h���ʕ\��		2009/12/17
call CreateRecommendHTML()

'---- GV Add Start 2012/04/10 
call CreateRecommendBuyHTML()
'---- GV Add End 2012/04/10 

'---- 2012/10/30 nt add Start
Call CreateContentsHTML()
'---- 2012/10/30 nt add End

'----- �֘A�V���[�Y���iHTML�쐬
'if RS("�V���[�Y�R�[�h") <> "" then
'	call CreateSeriesHTML()
'end if

'=================================================
'2013/08/07 if-web del s
'----- ���O�C�����Ă���΁A�ŋ߃`�F�b�N�������i��ǉ��쐬
'if wUserID <> "" then
'	call CreateViewedProductList()	'2012/07/10 GV Add
'	call AddViewdProduct()
'end if
'2013/08/07 if-web del e

'----- ���i�A�N�Z�X�����o�^	'2012/02/20 na ���X�|���X�΍�̂��ߒ�~
'call SetAccessCount()

'----- ���R�����h���i�A�N�Z�X���O�o�^   2009/12/17 add 2010/03/04 an �L���� 2012/01/10 an ��~
'call AddRecommendAccessLog()

RS.Close

End function

'========================================================================
'
'	Function	���i�����o��
'
'========================================================================
'
Function GetProduct()

Dim vInventoryCd
Dim vInventoryImage
Dim vSetCount
Dim vRowspan
Dim vWidth
Dim vHeight

Dim vIroCnt
Dim vKikakuCnt
Dim vProdPic(4)

Dim RSv

'---- �F�K�i���菤�i���ǂ����̃`�F�b�N
wSQL = ""
wSQL = wSQL & "SELECT a.�F"
wSQL = wSQL & "     , a.�K�i"
wSQL = wSQL & "     , a.�����\����"    '2010/11/26 an add
wSQL = wSQL & "     , a.��������"    			'2011/06/09 an add
wSQL = wSQL & "     , a.���iID"			'2013/05/22 GV #1505 add
wSQL = wSQL & "  FROM Web�F�K�i�ʍ݌� a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)
wSQL = wSQL & "   AND a.�I���� IS NULL"
wSQL = wSQL & " ORDER BY a.�F"
wSQL = wSQL & "     , a.�K�i"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wIroKikakuCombo = ""
vIroCnt = 0
vKikakuCnt = 0
wIroKikakuFl = "N"        '2010/11/26 an add
wIroKikakuZaikoFl = "N"   '2010/11/26 an add
wIroKikakuHacchuuFl = "N"   '2011/06/09 an add

if RSv.EOF = false then

	if RSv.RecordCount > 1 OR Trim(RSv("�F")) <> "" OR Trim(RSv("�K�i")) <> "" then	'2010/01/26 hn change

		wIroKikakuFl = "Y"  '�F�K�i�L   2010/11/26 an add

		'2013/05/22 GV #1505 add start
		Set wEAIroKikakuData = CreateObject("Scripting.Dictionary")
		wEAIroKikakuData.Item("code") = ""
		wEAIroKikakuData.Item("stock") = 0
		wEAIroKikakuData.Item("iro") = ""
		wEAIroKikakuData.Item("kikaku") = ""
		'2013/05/22 GV #1505 add end

		'----�F�K�i�w���1�����Ȃ��ꍇ�́A���̐F�K�i���Z�b�g����
		if RSv.RecordCount = 1 AND (Trim(RSv("�F")) <> "" OR Trim(RSv("�K�i")) <> "") then	'2010/01/26 hn add
			iro = Trim(RSv("�F"))
			kikaku = Trim(RSv("�K�i"))
		end if

		wIroKikakuCombo = wIroKikakuCombo & "                    <p class='color'>" & vbNewLine
		wIroKikakuCombo = wIroKikakuCombo & "                        <select name='IroKikaku' onChange='IroKikaku_onChange(this.form);'>" & vbNewLine
		wIroKikakuCombo = wIroKikakuCombo & "                            <option value=''>�I��" & vbNewLine

		Do until RSv.EOF = true

			'2013/05/22 GV #1505 add start
			If (wEAIroKikakuData.Item("code") = "") Then
				wEAIroKikakuData.Item("code") = Trim(RSv("���iID"))
				wEAIroKikakuData.Item("iro") = Trim(RSv("�F"))
				wEAIroKikakuData.Item("kikaku") = Trim(RSv("�K�i"))
			End If
			'2013/05/22 GV #1505 add end

			'---- ��p�@��\���`�F�b�N�Ή�  2010/11/26 an add s
			if RSv("�����\����") > 0 then
				wIroKikakuZaikoFl = "Y"

				'2013/05/22 GV #1505 add start
				If (wEAIroKikakuData.Item("stock") = 0) Then
					wEAIroKikakuData.Item("code") = Trim(RSv("���iID"))
					wEAIroKikakuData.Item("iro") = Trim(RSv("�F"))
					wEAIroKikakuData.Item("kikaku") = Trim(RSv("�K�i"))
					wEAIroKikakuData.Item("stock") = 1
				End If
				'2013/05/22 GV #1505 add end
			end if  '2010/11/26 an add s

'2011/06/09 hn add s
			if RSv("��������") > 0 then
				wIroKikakuHacchuuFl = "Y"
			end if  '2010/11/26 an add s
'2011/06/09 hn add e

			if Trim(RSv("�F")) <> "" AND Trim(RSv("�K�i")) <> "" then
				if Trim(RSv("�F")) = iro AND Trim(RSv("�K�i")) = kikaku then
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='" & Trim(RSv("�F")) & "^" & Trim(RSv("�K�i")) & "' SELECTED>" & Trim(RSv("�F")) & "/" & Trim(RSv("�K�i")) & vbNewLine

					'2013/05/22 GV #1505 add start
					wEAIroKikakuData.Item("code") = Trim(RSv("���iID"))
					wEAIroKikakuData.Item("iro") = Trim(RSv("�F"))
					wEAIroKikakuData.Item("kikaku") = Trim(RSv("�K�i"))
					'2013/05/22 GV #1505 add end
				else
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='" & Trim(RSv("�F")) & "^" & Trim(RSv("�K�i")) & "'>" & Trim(RSv("�F")) & "/" & Trim(RSv("�K�i")) & vbNewLine
				end if
				vIroCnt = vIroCnt + 1
				vKikakuCnt = vKikakuCnt + 1
			end if

			if Trim(RSv("�F")) <> "" AND Trim(RSv("�K�i")) = "" then
				if Trim(RSv("�F")) = iro AND Trim(RSv("�K�i")) = kikaku then
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='" & Trim(RSv("�F")) & "^' SELECTED>" & Trim(RSv("�F")) & vbNewLine

					'2013/05/22 GV #1505 add start
					wEAIroKikakuData.Item("code") = Trim(RSv("���iID"))
					wEAIroKikakuData.Item("iro") = Trim(RSv("�F"))
					wEAIroKikakuData.Item("kikaku") = Trim(RSv("�K�i"))
					'2013/05/22 GV #1505 add end
				else
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='" & Trim(RSv("�F")) & "^'>" & Trim(RSv("�F")) & vbNewLine
				end if
				vIroCnt = vIroCnt + 1
			end if

			if Trim(RSv("�F")) = "" AND Trim(RSv("�K�i")) <> "" then
				if Trim(RSv("�F")) = iro AND Trim(RSv("�K�i")) = kikaku then
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='^" & Trim(RSv("�K�i")) & "' SELECTED>" & Trim(RSv("�K�i")) & vbNewLine

					'2013/05/22 GV #1505 add start
					wEAIroKikakuData.Item("code") = Trim(RSv("���iID"))
					wEAIroKikakuData.Item("iro") = Trim(RSv("�F"))
					wEAIroKikakuData.Item("kikaku") = Trim(RSv("�K�i"))
					'2013/05/22 GV #1505 add end
				else
					wIroKikakuCombo = wIroKikakuCombo & "                            <option value='^" & Trim(RSv("�K�i")) & "'>" & Trim(RSv("�K�i")) & vbNewLine
				end if
				vKikakuCnt = vKikakuCnt + 1
			end if

			RSv.MoveNext
		Loop
		wIroKikakuCombo = wIroKikakuCombo & "                        </select>" & vbNewLine
		wIroKikakuCombo = wIroKikakuCombo & "                    </p>" & vbNewLine

		if vIroCnt > 0 AND vKikakuCnt > 0 then
			wIroKikakuCombo = Replace(wIroKikakuCombo, "�I��", "�F�K�i��I��")
			wIroKikakuSelectMsg = "�F�K�i��I�����Ă�������"
		end if
		if vIroCnt > 0 AND vKikakuCnt = 0 then
			wIroKikakuCombo = Replace(wIroKikakuCombo, "�I��", "�F��I��")
			wIroKikakuSelectMsg = "�F��I�����Ă�������"
		end if
		if vIroCnt = 0 AND vKikakuCnt > 0 then
			wIroKikakuCombo = Replace(wIroKikakuCombo, "�I��", "�K�i��I��")
			wIroKikakuSelectMsg = "�K�i��I�����Ă�������"
		end if

	else
		wIroKikakuCombo = wIroKikakuCombo & "                    <input type='hidden' name='IroKikaku' value='" & Trim(RSv("�F")) & "^" & Trim(RSv("�K�i")) & "'>" & vbNewLine

		iro = Trim(RSv("�F"))
		kikaku = Trim(RSv("�K�i"))

	end if

	if RSv.RecordCount <= 1 OR iro <> "" OR kikaku <> "" then
		wIroKikakuSelectedFl = true
	else
		wIroKikakuSelectedFl = false
	end if
end if

RSv.close

'---- ���i�����o��
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT a.���i�R�[�h"
wSQL = wSQL & "     , a.���i��"
wSQL = wSQL & "     , a.�̔��P��"
wSQL = wSQL & "     , a.������P��"
wSQL = wSQL & "     , a.�����萔��"
wSQL = wSQL & "     , a.������󒍍ϐ���"
wSQL = wSQL & "     , a.���i���l"
wSQL = wSQL & "     , a.���i�T��Web"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��2"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��3"
wSQL = wSQL & "     , a.���i�摜�t�@�C����_��4"
wSQL = wSQL & "     , a.ASK���i�t���O"
wSQL = wSQL & "     , a.�I����"
wSQL = wSQL & "     , a.�戵���~��"
wSQL = wSQL & "     , a.�p�ԓ�"
wSQL = wSQL & "     , a.Web���i�t���O"
wSQL = wSQL & "     , a.�J�e�S���[�R�[�h"
wSQL = wSQL & "     , a.���[�J�[�R�[�h"
wSQL = wSQL & "     , a.�󏭐���"
wSQL = wSQL & "     , a.�Z�b�g���i�t���O"
wSQL = wSQL & "     , a.���[�J�[�������敪"
wSQL = wSQL & "     , a.�V���i�Љ�URL"
wSQL = wSQL & "     , a.���i�}�j���A��URL"
wSQL = wSQL & "     , a.�����t���O"
wSQL = wSQL & "     , a.����URL"
wSQL = wSQL & "     , a.����t���O"
wSQL = wSQL & "     , a.����URL"
wSQL = wSQL & "     , a.���A���i�t���O"
wSQL = wSQL & "     , a.�����ߏ��i�R�����g"
wSQL = wSQL & "     , a.�֘A�L���^�C�g��1"
wSQL = wSQL & "     , a.�֘A�L��URL1"
wSQL = wSQL & "     , a.�֘A�L���^�C�g��2"
wSQL = wSQL & "     , a.�֘A�L��URL2"
wSQL = wSQL & "     , a.�֘A�L���^�C�g��3"
wSQL = wSQL & "     , a.�֘A�L��URL3"
wSQL = wSQL & "     , a.�֘A�L���^�C�g��4"
wSQL = wSQL & "     , a.�֘A�L��URL4"
wSQL = wSQL & "     , a.�V���[�Y�R�[�h"
wSQL = wSQL & "     , a.Web�[����\���t���O"
wSQL = wSQL & "     , a.��p�@�탁�[�J�[�R�[�h"
wSQL = wSQL & "     , a.��p�@�폤�i�R�[�h"
wSQL = wSQL & "     , a.���ח\�薢��t���O"
wSQL = wSQL & "     , a.���i�X�y�b�N�g�p�s�t���O"
wSQL = wSQL & "     , a.B�i�P��"
wSQL = wSQL & "     , a.������"
wSQL = wSQL & "     , a.B�i�t���O"
wSQL = wSQL & "     , a.���i���l�C���T�[�gURL1"
wSQL = wSQL & "     , a.���i���l�C���T�[�gURL2"
wSQL = wSQL & "     , a.���i���l�C���T�[�g�T�C�YW1"
wSQL = wSQL & "     , a.���i���l�C���T�[�g�T�C�YW2"
wSQL = wSQL & "     , a.���i���l�C���T�[�g�T�C�YH1"
wSQL = wSQL & "     , a.���i���l�C���T�[�g�T�C�YH2"
wSQL = wSQL & "     , a.�������S�������i�t���O"				' 2011/02/18 GV Add
wSQL = wSQL & "     , a.�O��P���ύX��"						' 2012/07/13 ok Add
wSQL = wSQL & "     , a.�O��̔��P��"						' 2012/07/13 ok Add
wSQL = wSQL & "     , b.���[�J�[��"
wSQL = wSQL & "     , b.���[�J�[���J�i"
wSQL = wSQL & "     , b.���[�J�[�z�[���y�[�WURL"
wSQL = wSQL & "     , b.�ڍ׏��^�C�g��1"
wSQL = wSQL & "     , b.�ڍ׏��URL1"
wSQL = wSQL & "     , b.�ڍ׏��Web�\���t���O1"
wSQL = wSQL & "     , b.�ڍ׏��^�C�g��2"
wSQL = wSQL & "     , b.�ڍ׏��URL2"
wSQL = wSQL & "     , b.�ڍ׏��Web�\���t���O2"
wSQL = wSQL & "     , b.�ڍ׏��^�C�g��3"
wSQL = wSQL & "     , b.�ڍ׏��URL3"
wSQL = wSQL & "     , b.�ڍ׏��Web�\���t���O3"
wSQL = wSQL & "     , b.�ڍ׏��^�C�g��4"
wSQL = wSQL & "     , b.�ڍ׏��URL4"
wSQL = wSQL & "     , b.�ڍ׏��Web�\���t���O4"
wSQL = wSQL & "     , c.�J�e�S���[��"

if wIroKikakuSelectedFl = true then
	wSQL = wSQL & "     , d.�F"
	wSQL = wSQL & "     , d.�K�i"
	wSQL = wSQL & "     , d.�����\����"
	wSQL = wSQL & "     , d.��������"								'2011/06/09 hn add
	wSQL = wSQL & "     , d.�����\���ח\���"
	wSQL = wSQL & "     , d.B�i�����\����"
	wSQL = wSQL & "     , d.�F�K�i���i�摜�t�@�C������"
	wSQL = wSQL & "     , d.�F�K�i���i�摜�t�@�C����1"
	wSQL = wSQL & "     , d.�F�K�i���i�摜�t�@�C����2"
	wSQL = wSQL & "     , d.�F�K�i���i�摜�t�@�C����3"
	wSQL = wSQL & "     , d.�F�K�i���i�摜�t�@�C����4"
	wSQL = wSQL & "     , d.���iID"
else
	wSQL = wSQL & "     , '' AS �F"
	wSQL = wSQL & "     , '' AS �K�i"
	wSQL = wSQL & "     , 0 AS �����\����"
	wSQL = wSQL & "     , 0 AS ��������"								'2011/06/09 hn add
	wSQL = wSQL & "     , NULL AS �����\���ח\���"
	wSQL = wSQL & "     , 0 AS B�i�����\����"
	wSQL = wSQL & "     , '' AS �F�K�i���i�摜�t�@�C������"
	wSQL = wSQL & "     , '' AS �F�K�i���i�摜�t�@�C����1"
	wSQL = wSQL & "     , '' AS �F�K�i���i�摜�t�@�C����2"
	wSQL = wSQL & "     , '' AS �F�K�i���i�摜�t�@�C����3"
	wSQL = wSQL & "     , '' AS �F�K�i���i�摜�t�@�C����4"
	wSQL = wSQL & "     , '' AS ���iID"
end if

wSQL = wSQL & "     , f.���J�e�S���[�R�[�h"
wSQL = wSQL & "     , f.���J�e�S���[�����{��"
wSQL = wSQL & "     , g.��J�e�S���[�R�[�h"
wSQL = wSQL & "     , g.��J�e�S���[��"
wSQL = wSQL & "     , g.�I�v�V�����p�[�c���o���\�L�t���O"	'2012/08/29 ok Add
wSQL = wSQL & "     , h.�V���[�Y��"				'2012/07/13 ok add
wSQL = wSQL & "     , h.�V���[�Y�摜�t�@�C����"	'2012/07/13 ok add
wSQL = wSQL & "     , h.�V���[�Y���l"			'2012/07/13 ok add
wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
wSQL = wSQL & "     LEFT JOIN �V���[�Y h WITH (NOLOCK) "		'2012/07/13 ok add
wSQL = wSQL & "     ON a.�V���[�Y�R�[�h = h.�V���[�Y�R�[�h"		'2012/07/13 ok add
wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
wSQL = wSQL & "     , �J�e�S���[ c WITH (NOLOCK)"

if wIroKikakuSelectedFl = true then
	wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d WITH (NOLOCK)"
end if

wSQL = wSQL & "     , ���J�e�S���[ f WITH (NOLOCK) "
wSQL = wSQL & "     , ��J�e�S���[ g WITH (NOLOCK) "
wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
wSQL = wSQL & "   AND f.���J�e�S���[�R�[�h = c.���J�e�S���[�R�[�h"
wSQL = wSQL & "   AND g.��J�e�S���[�R�[�h = f.��J�e�S���[�R�[�h"
wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
wSQL = wSQL & "   AND a.���[�J�[�R�[�h = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)

if wIroKikakuSelectedFl = true then
	wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
	wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
	wSQL = wSQL & "   AND d.�F = '" & iro & "'"
	wSQL = wSQL & "   AND d.�K�i = '" & kikaku & "'"
	wSQL = wSQL & "   AND d.�I���� IS NULL"
end if

'@@@@response.write(wSQL)

Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open wSQL, Connection, adOpenStatic

if RS.EOF = true then
	wPictureHTML = "<center><br><br><div class='honbun'><font color='#ff0000'>�@�@�@�@�Y�����i�͗L��܂���B</font></div></center>"
	wMSG = "no data"
	wNoData = "Y"
	exit function
end if

'---- �I���`�F�b�N
wProdTermFl = "N"

if isNull(RS("�戵���~��")) = false then		'�戵���~
	wProdTermFl = "Y"
end if

if isNull(RS("�p�ԓ�")) = false then  '2010/11/26 an mod s
	if wIroKikakuFl = "Y" then
		if wIroKikakuZaikoFl = "N" AND wIroKikakuHacchuuFl = "N" then		'�p�ԐF�K�i���菤�i�ŁA�S�F�K�i�݌ɂȂ�+�����Ȃ�	2011/06/09 hn mod
			wProdTermFl = "Y"
		end if
	else
		if RS("�����\����") <= 0 AND RS("��������") <= 0 then		'�p�Ԃō݌ɖ���+�����Ȃ�	'2011/06/09 hn mod
			wProdTermFl = "Y"
		end if
	end if
end if                                '2010/11/26 an mod e

if isNull(RS("������")) = false then		'�������i
	wProdTermFl = "Y"
end if

'---- �戵���~�܂��́A�p�ԍ݌ɖ����Ō�p�@��`�F�b�N ����Ό�p�@���\��
if wProdTermFl = "Y" AND RS("��p�@�탁�[�J�[�R�[�h") <> "" then
	wKoukeiMakerCd = RS("��p�@�탁�[�J�[�R�[�h")
	wKoukeiProductCd = RS("��p�@�폤�i�R�[�h")
	exit function
end if

'----- ���[�J�[���A���i��� �^�C�g��
wMakerName = RS("���[�J�[��")
wMakerNameNoKana = RS("���[�J�[��")  '2010/08/30 an add
if RS("���[�J�[���J�i") <> "" then
	wMakerName = wMakerName & " ( " & RS("���[�J�[���J�i") & " ) "
end if
wProductName = RS("���i��")
if trim(RS("�F")) <> "" then
	wProductName = wProductName & "/" & RS("�F")
end if
if trim(RS("�K�i")) <> "" then
	wProductName = wProductName & "/" & RS("�K�i")
end if

'---- �p�������X�g 2010/01/20 an �C��
wTitleWithLink = ""
'2012/07/10 GV Mod Start
'wTitleWithLink = "<h1><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='LargeCategoryList.asp?LargeCategoryCd=" & RS("��J�e�S���[�R�[�h") & "' class='link' itemprop='url'><span itemprop='title'>" & RS("��J�e�S���[��") & "</span>&gt;</a></span>"   '2011/11/22 an mod s
'wTitleWithLink = wTitleWithLink & "<span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='MidCategoryList.asp?MidCategoryCd=" & RS("���J�e�S���[�R�[�h") & "' class='link' itemprop='url'><span itemprop='title'>" & RS("���J�e�S���[�����{��") & "</span>&gt;</a></span>"
'wTitleWithLink = wTitleWithLink & "<span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=c&s_category_cd=" & RS("�J�e�S���[�R�[�h") & "' class='link' itemprop='url'><span itemprop='title'>" & RS("�J�e�S���[��") &  "</span>&gt;</a></span>"
'wTitleWithLink = wTitleWithLink & "<span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=cm&s_maker_cd=" & RS("���[�J�[�R�[�h") & "&s_category_cd=" & RS("�J�e�S���[�R�[�h") & "' class='link' itemprop='url'><span itemprop='title'>" & wMakerName & "</span></a></span>/" & wProductName & "</h1>"   '2011/11/22 an mod e
wTitleWithLink = "<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='LargeCategoryList.asp?LargeCategoryCd=" & RS("��J�e�S���[�R�[�h") & "' itemprop='url'><span itemprop='title'>" & RS("��J�e�S���[��") & "</span></a></span></li>"
wTitleWithLink = wTitleWithLink & "<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='MidCategoryList.asp?MidCategoryCd=" & RS("���J�e�S���[�R�[�h") & "' itemprop='url'><span itemprop='title'>" & RS("���J�e�S���[�����{��") & "</span></a></span></li>"
wTitleWithLink = wTitleWithLink & "<li><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=c&s_category_cd=" & RS("�J�e�S���[�R�[�h") & "' itemprop='url'><span itemprop='title'>" & RS("�J�e�S���[��") &  "</span></a></span></li>"
wTitleWithLink = wTitleWithLink & "<li class='now'><span itemscope itemtype='http://data-vocabulary.org/Breadcrumb'><a href='SearchList.asp?i_type=cm&s_maker_cd=" & RS("���[�J�[�R�[�h") & "&s_category_cd=" & RS("�J�e�S���[�R�[�h") & "' itemprop='url'><span itemprop='title'>" & wMakerName & "</span></a></span>/" & wProductName & "</li>"
'2012/07/10 GV Mod End

'---- �J�e�S���[�R�[�h�Z�[�u
wCategoryCode = RS("�J�e�S���[�R�[�h")
wLargeCategoryCd = RS("��J�e�S���[�R�[�h")
wMidCategoryCd = RS("���J�e�S���[�R�[�h")

'2012/08/29 ok Add
wOptionPartsTitleFlag = RS("�I�v�V�����p�[�c���o���\�L�t���O")

'2012/10/30 nt Add
wSeriesCd = RS("�V���[�Y�R�[�h")

'---- meta keywords�p    2010/08/23 an add s
wLargeCategoryName = RS("��J�e�S���[��")
wMidCategoryName = RS("���J�e�S���[�����{��")
wCategoryName = RS("�J�e�S���[��")  '2010/08/23 an add e

'---- �������S�����t���O
wFreeShippingFlag = RS("�������S�������i�t���O")			' 2011/02/18 GV Add

'2012/07/13 ok Add Start
'---- �V���[�Y���i
wSeriesHTML = ""
If RS("�V���[�Y�R�[�h") <> "" Then
	wSeriesHTML = wSeriesHTML & "      <div class='detail_side_inner01'><div class='detail_side_inner02'>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "        <div class='detail_side_inner_box'>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "          <h4 class='detail_sub'><a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("�V���[�Y�R�[�h") & "'>" & RS("�V���[�Y��") & "</a></h4>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "            <ul class='check_item'>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "              <li>" & vbNewLine
	If RS("�V���[�Y�摜�t�@�C����") <> "" Then
		wSeriesHTML = wSeriesHTML & "                <p><a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("�V���[�Y�R�[�h") & "'><img src='prod_img/" & RS("�V���[�Y�摜�t�@�C����") & "' alt='" & Replace(RS("�V���[�Y��"),"'","&#39;") & "' class='opover'></a>" & RS("�V���[�Y���l") & "</p>" & vbNewLine
	End If
	wSeriesHTML = wSeriesHTML & "              </li>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "          </ul>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "        </div>" & vbNewLine
	wSeriesHTML = wSeriesHTML & "      </div></div>" & vbNewLine
End If
'2012/07/13 ok Add End

Set wEAProductDetailData = RS	'2013/05/17 GV #1505 add

End Function

'========================================================================
'
'	Function	���i�摜 HTML�쐬
'
'========================================================================
'
Function CreatePictureHTML()

Dim vProdPic(4)

'----- ���i�摜
wHTML = ""
'wHTML = wHTML & "<table width='602' border='0' cellspacing='0' cellpadding='0' id='Shop_product_img'>" & vbNewLine	'2012/07/10 GV Del

if Trim(RS("�F�K�i���i�摜�t�@�C����1")) <> "" then
	vProdPic(1) = RS("�F�K�i���i�摜�t�@�C����1")
	vProdPic(2) = RS("�F�K�i���i�摜�t�@�C����2")
	vProdPic(3) = RS("�F�K�i���i�摜�t�@�C����3")
	vProdPic(4) = RS("�F�K�i���i�摜�t�@�C����4")
else
	vProdPic(1) = RS("���i�摜�t�@�C����_��")
	vProdPic(2) = RS("���i�摜�t�@�C����_��2")
	vProdPic(3) = RS("���i�摜�t�@�C����_��3")
	vProdPic(4) = RS("���i�摜�t�@�C����_��4")
end if

'2012/07/10 GV Mod Start
if vProdPic(1) <> "" then

	wMainProdPic = vProdPic(1)  '2011/11/22 an add

'	wHTML = wHTML & "  <tr align='left' valign='middle'>" & vbNewLine
'	wHTML = wHTML & "    <td><img name='LargeImage' src='prod_img/" & vProdPic(1) & "' alt='" & wMakerName & " / " & wProductname & "' itemprop='image' class='big'></td>" & vbNewLine     '2011/11/22 an mod
'	wHTML = wHTML & "  </tr>" & vbNewLine
	wHTML = wHTML & "  <div id='item_photo_box'>"
	wHTML = wHTML & "    <p><img src='prod_img/" & vProdPic(1) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & "' id='target' itemprop='image' class='opover'></p>"

end if

'if vProdPic(2) <> "" OR vProdPic(3) <> "" OR vProdPic(4) <> "" then
if vProdPic(1) <> "" OR vProdPic(2) <> "" OR vProdPic(3) <> "" OR vProdPic(4) <> "" then
'	wHTML = wHTML & "  <tr align='center' valign='middle'>" & vbNewLine
'	wHTML = wHTML & "    <td height='75' nowrap>" & vbNewLine
	wHTML = wHTML & "    <ul class='sub_box'>"
	if vProdPic(1) <> "" then
'		wHTML = wHTML & "    <img src='prod_img/" & vProdPic(1) & "' width='147' height='73' class='small1' alt='" & wMakerName & " / " & wProductname & " �摜1' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(1) & """);'>"
		wHTML = wHTML & "      <li><a class='modalImg' rel='fancybox' href='prod_img/" & vProdPic(1) & "'><img src='prod_img/" & vProdPic(1) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & " �摜1' class='opover'></a></li>"
	end if
	if vProdPic(2) <> "" then
'		wHTML = wHTML & "<img src='prod_img/" & vProdPic(2) & "' width='147' height='73' class='small1' alt='" & wMakerName & " / " & wProductname & " �摜2' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(2) & """);'>"
		wHTML = wHTML & "      <li><a class='modalImg' rel='fancybox' href='prod_img/" & vProdPic(2) & "'><img src='prod_img/" & vProdPic(2) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & " �摜2' class='opover'></a></li>"
	end if
	if vProdPic(3) <> "" then
'		wHTML = wHTML & "<img src='prod_img/" & vProdPic(3) & "' width='147' height='73' class='small1' alt='" & wMakerName & " / " & wProductname & " �摜3' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(3) & """);'>"
		wHTML = wHTML & "      <li><a class='modalImg' rel='fancybox' href='prod_img/" & vProdPic(3) & "'><img src='prod_img/" & vProdPic(3) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & " �摜3' class='opover'></a></li>"
	end if
	if vProdPic(4) <> "" then
'		wHTML = wHTML & "<img src='prod_img/" & vProdPic(4) & "' width='147' height='73' class='small1' alt='" & wMakerName & " / " & wProductname & " �摜4' onMouseOver='SmallImage_onMouseOver(""prod_img/" & vProdPic(4) & """);'>" & vbNewLine
		wHTML = wHTML & "      <li><a class='modalImg' rel='fancybox' href='prod_img/" & vProdPic(4) & "'><img src='prod_img/" & vProdPic(4) & "' alt='" & Replace(wMakerName & " / " & wProductname,"'","&#39;") & " �摜4' class='opover'></a></li>"
	end if
'	wHTML = wHTML & "    </td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
	wHTML = wHTML & "    </ul>"
	wHTML = wHTML & "  </div>"
end if
'2012/07/10 GV Mod End

'wHTML = wHTML & "</table>" & vbNewLine	'2012/07/10 GV Del

wPictureHTML = wHTML


End Function

'========================================================================
'
'	Function	�֘A�����N HTML�쐬
'
'========================================================================
'
Function CreateKanrenLinkHTML()

Dim vURL()
Dim vURLCount
Dim vURL2
Dim wHTMLTemp
Dim i

wHTML = ""
wHTMLTemp = ""

'----���惊���N
if RS("����t���O") = "Y" then
	vURLCount = cf_unstring(RS("����URL"), vURL, ",")

	if vURLCount > 1 then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "      <a href='JavaScript:void(0);' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('movie" & i & "','','images/movie_on.gif',1)"" onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "','SoundMovie', 'width=201 height=200 resizable=1 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><img src='images/movie_off.gif' border='0' name='movie" & i & "' alt='���������'></a>" & vbNewLine
		wHTML = wHTML & "      <li><a href='JavaScript:void(0);' onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "','SoundMovie', 'width=201 height=200 resizable=1 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');""><img src='images/btn_movie.png' alt='���������' class='opover'></a></li>" & vbNewLine
		'2012/07/10 GV Mod End
	else
		if InStr(vURL(0), "http://") > 0 then
			vURL2 = vURL(0)
		else
			vURL2 = g_HTTP & vURL(0)
		end if
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('movie" & i & "','','images/movie_on.gif',1)""><img src='images/movie_off.gif' border='0' name='movie" & i & "' alt='���������'></a>" & vbNewLine
		wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'><img src='images/btn_movie.png' alt='���������' class='opover'></a></li>" & vbNewLine
		'2012/07/10 GV Mod End
	end if

end if

'----���������N
if RS("�����t���O") = "Y" then
	vURLCount = cf_unstring(RS("����URL"), vURL, ",")

	if vURLCount > 1 then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "      <a href='JavaScript:void(0);' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('audio" & i & "','','images/audio_on.gif',1)"" onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "','SoundMovie', 'width=201 height=200 resizable=1 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><img src='images/audio_off.gif' border='0' name='audio" & i & "' alt='��������'></a>" & vbNewLine
		wHTML = wHTML & "      <li><a href='JavaScript:void(0);' onClick=""window.open('SoundMoviePopUp.asp?item=" & RS("���[�J�[�R�[�h") & "^" & Server.URLEncode(RS("���i�R�[�h")) & "','SoundMovie', 'width=201 height=200 resizable=1 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');""><img src='images/btn_view.png' alt='��������' class='opover'></a></li>" & vbNewLine
		'2012/07/10 GV Mod End
	else
		if InStr(vURL(0), "http://") > 0 then
			vURL2 = vURL(0)
		else
			vURL2 = g_HTTP & vURL(0)
		end if
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('audio" & i & "','','images/audio_on.gif',1)""><img src='images/audio_off.gif' border='0' name='audio" & i & "' alt='��������'></a>" & vbNewLine
		wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'><img src='images/btn_view.png' alt='��������' class='opover'></a></li>" & vbNewLine
		'2012/07/10 GV Mod End

	end if

end if

'---- ���[�J�[�z�[���y�[�W
if RS("���[�J�[�z�[���y�[�WURL") <> "" then
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & RS("���[�J�[�z�[���y�[�WURL") & "'target='_blank' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('maker','','images/maker_on.gif',1)""><img src='images/maker_off.gif' border='0' name='maker' alt='���[�J�[�T�C�g��'></a>" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & RS("���[�J�[�z�[���y�[�WURL") & "' target='_blank'><img src='images/btn_maker.png' alt='���[�J�[�T�C�g' class='opover'></a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

'---- ���i�}�j���A��
if RS("���i�}�j���A��URL") <> "" then
	if InStr(LCase(RS("���i�}�j���A��URL")), "http://") > 0 then
		vURL2 = Trim(RS("���i�}�j���A��URL"))
	else
		vURL2 = g_HTTP & Trim(RS("���i�}�j���A��URL"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 &  "'target='_blank' onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('manual" & i & "','','images/manual_on.gif',1)""><img src='images/manual_off.gif' border='0' name='manual" & i & "' alt='���i�}�j���A��'></a>" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'><img src='images/btn_manual.png' alt='���i�}�j���A��' class='opover'></a></li>" & vbNewLine
	'2012/07/10 GV Mod End
End if

'----
if wHTML <> "" then
'2012/07/10 GV Mod Start
'	wHTMLTemp = wHTMLTemp & "<table width='602' border='0' cellSpacing='0' cellPadding='0'>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "  <tr>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "    <td  height='40'>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & wHTML
'	wHTMLTemp = wHTMLTemp & "    </td>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "  </tr>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "  <div class='btn_box'>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "    <ul class='btn'>" & vbNewLine
	wHTMLTemp = wHTMLTemp & wHTML
	wHTMLTemp = wHTMLTemp & "    </ul>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "  </div>" & vbNewLine
'2012/07/10 GV Mod End
end if

'---- ���[�J�[�֘A�L��URL1-4
wHTML = ""

if RS("�֘A�L��URL1") <> "" then

	if InStr(LCase(RS("�֘A�L��URL1")), "http://") > 0 then
		vURL2 = Trim(RS("�֘A�L��URL1"))
	else
		vURL2 = g_HTTP & Trim(RS("�֘A�L��URL1"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank'>" & RS("�֘A�L���^�C�g��1") & "</a>&nbsp;&nbsp;|" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'>" & RS("�֘A�L���^�C�g��1") & "</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

if RS("�֘A�L��URL2") <> "" then

	if InStr(LCase(RS("�֘A�L��URL2")), "http://") > 0 then
		vURL2 = Trim(RS("�֘A�L��URL2"))
	else
		vURL2 = g_HTTP & Trim(RS("�֘A�L��URL2"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank'>" & RS("�֘A�L���^�C�g��2") & "</a>&nbsp;&nbsp;|" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'>" & RS("�֘A�L���^�C�g��2") & "</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

if RS("�֘A�L��URL3") <> "" then

	if InStr(LCase(RS("�֘A�L��URL3")), "http://") > 0 then
		vURL2 = Trim(RS("�֘A�L��URL3"))
	else
		vURL2 = g_HTTP & Trim(RS("�֘A�L��URL3"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank'>" & RS("�֘A�L���^�C�g��3") & "</a>&nbsp;&nbsp;|" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'>" & RS("�֘A�L���^�C�g��3") & "</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if


if RS("�֘A�L��URL4") <> "" then

	if InStr(LCase(RS("�֘A�L��URL4")), "http://") > 0 then
		vURL2 = Trim(RS("�֘A�L��URL4"))
	else
		vURL2 = g_HTTP & Trim(RS("�֘A�L��URL4"))
	end if

	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "      <a href='" & vURL2 & "' target='_blank'>" & RS("�֘A�L���^�C�g��4") & "</a>&nbsp;&nbsp;|" & vbNewLine
	wHTML = wHTML & "      <li><a href='" & vURL2 & "' target='_blank'>" & RS("�֘A�L���^�C�g��4") & "</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

'----
if wHTML <> "" then
	'2012/07/10 GV Del Start
'	wHTML = Left(wHTML, Len(wHTML)-3)
'	if wHTMLTemp = "" then
'		wHTMLTemp = wHTMLTemp & "<table border=0 cellSpacing=0 cellPadding=0 width=602>" & vbNewLine
'	end if
	'2012/07/10 GV Del End

	'2012/07/10 GV Mod Start
'	wHTMLTemp = wHTMLTemp & "  <tr>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "    <td  height='40'>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & wHTML & vbNewLine
'	wHTMLTemp = wHTMLTemp & "    </td>" & vbNewLine
'	wHTMLTemp = wHTMLTemp & "  </tr>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "  <div class='other_link'>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "    <ul class='link'>" & vbNewLine
	wHTMLTemp = wHTMLTemp & wHTML & vbNewLine
	wHTMLTemp = wHTMLTemp & "    </ul>" & vbNewLine
	wHTMLTemp = wHTMLTemp & "  </div>" & vbNewLine
	'2012/07/10 GV Mod End
end if

'2012/07/10 GV Del Start
'if wHTMLTemp <> "" then
'	wHTMLTemp = wHTMLTemp & "</table>" & vbNewLine
'end if
'2012/07/10 GV Del End

wKanrenLinkHTML = wHTMLTemp

End function


'========================================================================
'
'	Function	���� HTML�쐬
'
'========================================================================
'
Function CreateTokuchoHTML()

wHTML = ""

'---- ����, ���A���i
if RS("�����ߏ��i�R�����g") <> "" OR RS("���A���i�t���O") = "Y" then
'2012/07/10 GV Mod Start
'	wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='main_header'>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td><strong>����&nbsp;&nbsp;[" & RS("���[�J�[��") & "(" & RS("���[�J�[���J�i") & ")/" & RS("���i��") & "]</strong></td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine
'	wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='shop_border'>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td>" & vbNewLine
'	wHTML = wHTML & "      <p>"
	wHTML = wHTML & "<div class='inner_box_spec'>" & vbNewLine

	if RS("�����ߏ��i�R�����g") <> "" then
		wHTML = wHTML & "<span itemprop='description'>"    '2011/11/22 an add
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & RS("�����ߏ��i�R�����g") & "<br>" & vbNewLine
		wHTML = wHTML & RS("�����ߏ��i�R�����g") & vbNewLine
		'2012/07/10 GV Mod End
		wHTML = wHTML & "</span>"                          '2011/11/22 an add

		'---- meta description�p�f�[�^�擾          '2010/08/23 an add s
		wTokucho = fDeleteHTMLTag(RS("�����ߏ��i�R�����g")) 'HTML�^�O�폜
		wTokucho = replace(replace(replace(replace(wTokucho, vbCr, ""), vbLf, ""), vbTab, ""), """", "") '���s�ATab�̍폜

		if Len(wTokucho) > 97 then  '�����ꍇ��100�����ɏȗ�
			wTokucho = Left(wTokucho,97) & "..."
		end if                                      '2010/08/23 an add e

	end if

	if RS("���A���i�t���O") = "Y" then
'		wHTML = wHTML & "<a href='../information/direct_import.asp' class='link'>[���A���i]</a>" & vbNewLine
		wHTML = wHTML & "  <p><a href='../information/direct_import.asp'>[���A���i]</a></p>" & vbNewLine
	end if

'	wHTML = wHTML & "      </p>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine
	wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End
end if

wTokuchoHTML = wHTML

End Function

'========================================================================
'
'	Function	�X�y�b�N HTML�쐬
'
'========================================================================
'
Function CreateSpecificationHTML()

Dim vWidth
Dim vHeight
Dim vRemark          '2011/11/22 an add
Dim vFileExtention   '2011/11/22 an add
Dim vURL1HTML        '2011/11/22 an add
Dim vURL2HTML        '2011/11/22 an add

wHTML = ""
vURL1HTML = ""       '2011/11/22 an add
vURL2HTML = ""       '2011/11/22 an add

'---- �X�y�b�N
'2012/07/10 GV Del Start
'wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='main_header'>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td><h2>�X�y�b�N&nbsp;&nbsp;[" & RS("���[�J�[��") & "(" & RS("���[�J�[���J�i") & ")/" & RS("���i��") & "]</h2></td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
'2012/07/10 GV Del End

if RS("���i���l�C���T�[�gURL1") <> "" then

	if InStr(LCase(RS("���i���l�C���T�[�gURL1")), "http") > 0 then   '2011/11/22 an mod s
	else
		'---- txt�t�@�C���̏ꍇ
vFileExtention = LCase(Right(RS("���i���l�C���T�[�gURL1"), 3))
		if LCase(Right(RS("���i���l�C���T�[�gURL1"), 3)) = "txt" then

			'---- �t�@�C���̑��݊m�F
			vRemark = GetMapPath(RS("���i���l�C���T�[�gURL1"), vFileExtention)

			if vRemark <> "" then
'				vURL1HTML = vURL1HTML & "<div class='insert'>" & vbNewLine	'2012/07/10 GV Del
				vURL1HTML = vURL1HTML & cf_read_file_all(vRemark) & vbNewLine
'				vURL1HTML = vURL1HTML & "</div>" & vbNewLine			'2012/07/10 GV Del
			end if
		'---- txt�t�@�C���ȊO�̏ꍇ
'2012/07/19 ok Del Start txt�ȊO�̏ꍇ�A���������K�v�Ȃ��ߔ�\���Ƃ���
'		else
'
'			if RS("���i���l�C���T�[�g�T�C�YW1") <> 0 then
'				vWidth = RS("���i���l�C���T�[�g�T�C�YW1")
'				if vWidth > 600 then
'					vWidth = 600
'				end if
'			else
'				vWidth = 600
'			end if
'
'			if RS("���i���l�C���T�[�g�T�C�YH1") <> 0 then
'				vHeight = RS("���i���l�C���T�[�g�T�C�YH1")
'			else
'				vHeight = 290
'			end if
'
'			vURL1HTML = vURL1HTML & "<iframe class='insert' marginwidth='0' marginheight='0' scrolling='no' src='" & RS("���i���l�C���T�[�gURL1") & "' width='" & vWidth & "' height='" & vHeight & "' frameborder='0'></iframe>"
'2012/07/19 ok Del End
		end if

		if vURL1HTML <> "" then
			if InStr(LCase(RS("���i���l�C���T�[�gURL1")), "http") > 0 then
			else
'2012/07/10 GV Mod Start
'				wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='shop_border_insert'>" & vbNewLine
'				wHTML = wHTML & "  <tr>" & vbNewLine
'				wHTML = wHTML & "    <td>" & vURL1HTML & "</td>" & vbNewLine
'				wHTML = wHTML & "  </tr>" & vbNewLine
'				wHTML = wHTML & "</table>" & vbNewLine
				wHTML = wHTML & "<div class='insert_box'>" & vbNewLine
				wHTML = wHTML & vURL1HTML & vbNewLine
				wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End
			end if
		end if

	end if   '2011/11/22 an mod e
end if

'2012/07/10 GV Mod Start
'wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='shop_border'>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td><p>" & CreateSpecHTML(RS("�J�e�S���[�R�[�h"),RS("���[�J�[�R�[�h"),RS("���i�R�[�h"),RS("���i���l"),RS("���i�X�y�b�N�g�p�s�t���O")) & "</p></td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
wHTML = wHTML & "<div class='inner_box_spec'>" & vbNewLine
wHTML = wHTML & CreateSpecHTML(RS("�J�e�S���[�R�[�h"),RS("���[�J�[�R�[�h"),RS("���i�R�[�h"),RS("���i���l"),RS("���i�X�y�b�N�g�p�s�t���O")) & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End

if RS("���i���l�C���T�[�gURL2") <> "" then

	if InStr(LCase(RS("���i���l�C���T�[�gURL2")), "http") > 0 then   '2011/11/22 an mod s
	else
		'---- txt�t�@�C���̏ꍇ
		if LCase(Right(RS("���i���l�C���T�[�gURL2"), 3)) = "txt" then

			'---- �t�@�C���̑��݊m�F
			vRemark = GetMapPath(RS("���i���l�C���T�[�gURL2"), vFileExtention)

			if vRemark <> "" then
'				vURL2HTML = vURL2HTML & "<div class='insert'>" & vbNewLine	'2012/07/10 GV Del
				vURL2HTML = vURL2HTML & cf_read_file_all(vRemark) & vbNewLine
'				vURL2HTML = vURL2HTML & "</div>" & vbNewLine			'2012/07/10 GV Del
			end if
'2012/07/19 ok Del Start txt�ȊO�̏ꍇ�A���������K�v�Ȃ��ߔ�\���Ƃ���
'		else
'
'			if RS("���i���l�C���T�[�g�T�C�YW2") <> 0 then
'				vWidth = RS("���i���l�C���T�[�g�T�C�YW2")
'			else
'				vWidth = 600
'			end if
'
'			if RS("���i���l�C���T�[�g�T�C�YH2") <> 0 then
'				vHeight = RS("���i���l�C���T�[�g�T�C�YH2")
'			else
'				vHeight = 300
'			end if
'
'			vURL2HTML = vURL2HTML & "<iframe class='insert' marginwidth='0' marginheight='0' scrolling='no' src='" & RS("���i���l�C���T�[�gURL2") & "' width='" & vWidth & "' height='" & vHeight & "' frameborder='0'></iframe>"
'2012/07/19 ok Del End
		end if

		if vURL2HTML <> "" then
			if InStr(LCase(RS("���i���l�C���T�[�gURL2")), "http") > 0 then
			else
'2012/07/10 GV Mod Start
'				wHTML = wHTML & "<table width='602' border='0' cellpadding='0' cellspacing='0' id='shop_border_insert'>" & vbNewLine
'				wHTML = wHTML & "  <tr>" & vbNewLine
'				wHTML = wHTML & "    <td>" & vURL2HTML & "</td>" & vbNewLine
'				wHTML = wHTML & "  </tr>" & vbNewLine
'				wHTML = wHTML & "</table>" & vbNewLine
				wHTML = wHTML & "<div class='insert_box'>" & vbNewLine
				wHTML = wHTML & vURL2HTML & vbNewLine
				wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End
			end if
		end if

	end if   '2011/11/22 an mod e
end if

wSpecHTML = wHTML

End Function

'========================================================================
'
'	Function	�I�v�V���� HTML�i�f�[�^���o�j
'
'		�F�K�i�Ɋ֌W�Ȃ��Y�����i�̃I�v�V���������o��
'
'========================================================================
'
Function CreateOptionHTML()

'---- Select �I�v�V�������i
wSQL = ""
' 2012/01/18 GV Mod Start
'wSQL = wSQL & "SELECT c.�I�v�V�������[�J�[�R�[�h AS ���[�J�[�R�[�h"
'wSQL = wSQL & "     , c.�I�v�V�������i�R�[�h AS ���i�R�[�h"
'wSQL = wSQL & "     , d.�F AS �F"
'wSQL = wSQL & "     , d.�K�i AS �K�i"
'wSQL = wSQL & "     , a.���i��"
''wSQL = wSQL & "     , a.�̔��P��"     '2010/11/10 an del
'
''2010/11/10 an add s
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B�i�t���O = 'Y' THEN a.B�i�P��"
'wSQL = wSQL & "         WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS ���̔��P��"
''2010/11/10 an add e
'
'wSQL = wSQL & "     , a.���i�T��Web"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , a.�戵���~��"
'wSQL = wSQL & "     , a.�p�ԓ�"
'wSQL = wSQL & "     , a.������"
'wSQL = wSQL & "     , a.�󏭐���"
'wSQL = wSQL & "     , a.�Z�b�g���i�t���O"
'wSQL = wSQL & "     , a.���[�J�[�������敪"
'wSQL = wSQL & "     , a.Web�[����\���t���O"
'wSQL = wSQL & "     , a.���ח\�薢��t���O"
'wSQL = wSQL & "     , a.B�i�t���O"
'wSQL = wSQL & "     , a.�����萔��"
'wSQL = wSQL & "     , a.������󒍍ϐ���"
'wSQL = wSQL & "     , b.���[�J�[��"
'wSQL = wSQL & "     , d.�����\����"
'wSQL = wSQL & "     , d.��������"			'2011/06/09 hn add
'wSQL = wSQL & "     , d.�����\���ח\���"
'wSQL = wSQL & "     , d.B�i�����\����"
'wSQL = wSQL & "  FROM Web���i a"
'wSQL = wSQL & "     , ���[�J�[ b"
'wSQL = wSQL & "     , �I�v�V����2 c"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d"
'wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.���i�R�[�h = c.�I�v�V�������i�R�[�h"
'wSQL = wSQL & "   AND b.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = c.�I�v�V�������i�R�[�h"
'wSQL = wSQL & "   AND d.�F = c.�I�v�V�����F"
'wSQL = wSQL & "   AND d.�K�i = c.�I�v�V�����K�i"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = '" & maker_cd & "'"
'wSQL = wSQL & "   AND c.���i�R�[�h = '" & product_cd & "'"
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'
'wSQL = wSQL & " UNION "
'
'wSQL = wSQL & "SELECT c.�I�v�V�������[�J�[�R�[�h AS ���[�J�[�R�[�h"
'wSQL = wSQL & "     , c.�I�v�V�������i�R�[�h AS ���i�R�[�h"
'wSQL = wSQL & "     , d.�F AS �F"
'wSQL = wSQL & "     , d.�K�i AS �K�i"
'wSQL = wSQL & "     , a.���i��"
''wSQL = wSQL & "     , a.�̔��P��"     '2010/11/10 an del
'
''2010/11/10 an add s
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B�i�t���O = 'Y' THEN a.B�i�P��"
'wSQL = wSQL & "         WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS ���̔��P��"
''2010/11/10 an add e
'
'wSQL = wSQL & "     , a.���i�T��Web"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , a.�戵���~��"
'wSQL = wSQL & "     , a.�p�ԓ�"
'wSQL = wSQL & "     , a.������"
'wSQL = wSQL & "     , a.�󏭐���"
'wSQL = wSQL & "     , a.�Z�b�g���i�t���O"
'wSQL = wSQL & "     , a.���[�J�[�������敪"
'wSQL = wSQL & "     , a.Web�[����\���t���O"
'wSQL = wSQL & "     , a.���ח\�薢��t���O"
'wSQL = wSQL & "     , a.B�i�t���O"
'wSQL = wSQL & "     , a.�����萔��"
'wSQL = wSQL & "     , a.������󒍍ϐ���"
'wSQL = wSQL & "     , b.���[�J�[��"
'wSQL = wSQL & "     , d.�����\����"
'wSQL = wSQL & "     , d.��������"			'2011/06/09 hn add
'wSQL = wSQL & "     , d.�����\���ח\���"
'wSQL = wSQL & "     , d.B�i�����\����"
'wSQL = wSQL & "  FROM Web���i a"
'wSQL = wSQL & "     , ���[�J�[ b"
'wSQL = wSQL & "     , �J�e�S���[�ʃI�v�V���� c"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d"
'wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.���i�R�[�h = c.�I�v�V�������i�R�[�h"
'wSQL = wSQL & "   AND b.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = c.�I�v�V�������i�R�[�h"
'wSQL = wSQL & "   AND d.�F = c.�I�v�V�����F"
'wSQL = wSQL & "   AND d.�K�i = c.�I�v�V�����K�i"
'wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = '" & wCategoryCode & "'"
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       b.���[�J�[��"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , d.�F"
'wSQL = wSQL & "     , d.�K�i"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      c.�I�v�V�������[�J�[�R�[�h AS ���[�J�[�R�[�h "
wSQL = wSQL & "    , c.�I�v�V�������i�R�[�h AS ���i�R�[�h "
wSQL = wSQL & "    , d.�F AS �F "
wSQL = wSQL & "    , d.�K�i AS �K�i "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B�i�t���O = 'Y'                     THEN a.B�i�P�� "
wSQL = wSQL & "        WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P�� "
wSQL = wSQL & "        ELSE                                            a.�̔��P�� "
wSQL = wSQL & "      END AS ���̔��P�� "
wSQL = wSQL & "    , a.���i�T��Web "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.ASK���i�t���O "
wSQL = wSQL & "    , a.�戵���~�� "
wSQL = wSQL & "    , a.�p�ԓ� "
wSQL = wSQL & "    , a.������ "
wSQL = wSQL & "    , a.�󏭐��� "
wSQL = wSQL & "    , a.�Z�b�g���i�t���O "
wSQL = wSQL & "    , a.���[�J�[�������敪 "
wSQL = wSQL & "    , a.Web�[����\���t���O "
wSQL = wSQL & "    , a.���ח\�薢��t���O "
wSQL = wSQL & "    , a.B�i�t���O "
wSQL = wSQL & "    , a.�����萔�� "
wSQL = wSQL & "    , a.������󒍍ϐ��� "
wSQL = wSQL & "    , b.���[�J�[�� "
wSQL = wSQL & "    , d.�����\���� "
wSQL = wSQL & "    , d.�������� "
wSQL = wSQL & "    , d.�����\���ח\��� "
wSQL = wSQL & "    , d.B�i�����\���� "
wSQL = wSQL & "    , a.�J�e�S���[�R�[�h "		'2012/08/27 ok Add
wSQL = wSQL & "    , e.�J�e�S���[�� "			'2012/08/27 ok Add
wSQL = wSQL & "FROM "
wSQL = wSQL & "    �I�v�V����2                  c WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web���i         a WITH (NOLOCK) "
wSQL = wSQL & "        ON     a.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h "
wSQL = wSQL & "           AND a.���i�R�[�h     = c.�I�v�V�������i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h "
wSQL = wSQL & "           AND d.���i�R�[�h     = c.�I�v�V�������i�R�[�h "
wSQL = wSQL & "           AND d.�F             = c.�I�v�V�����F "
wSQL = wSQL & "           AND d.�K�i           = c.�I�v�V�����K�i "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O    = t1.ShohinWebY "
wSQL = wSQL & "      INNER JOIN �J�e�S���[ e WITH (NOLOCK) "			'2012/08/27 ok Add
wSQL = wSQL & "        ON     a.�J�e�S���[�R�[�h = e.�J�e�S���[�R�[�h "	'2012/08/27 ok Add
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND c.���[�J�[�R�[�h = '" & maker_cd & "' "
wSQL = wSQL & "    AND c.���i�R�[�h     = '" & Replace(product_cd, "'", "''") & "' "	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)

wSQL = wSQL & "UNION "

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      c.�I�v�V�������[�J�[�R�[�h AS ���[�J�[�R�[�h "
wSQL = wSQL & "    , c.�I�v�V�������i�R�[�h AS ���i�R�[�h "
wSQL = wSQL & "    , d.�F AS �F "
wSQL = wSQL & "    , d.�K�i AS �K�i "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B�i�t���O = 'Y'                     THEN a.B�i�P�� "
wSQL = wSQL & "        WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P�� "
wSQL = wSQL & "        ELSE                                            a.�̔��P�� "
wSQL = wSQL & "      END AS ���̔��P�� "
wSQL = wSQL & "    , a.���i�T��Web "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.ASK���i�t���O "
wSQL = wSQL & "    , a.�戵���~�� "
wSQL = wSQL & "    , a.�p�ԓ� "
wSQL = wSQL & "    , a.������ "
wSQL = wSQL & "    , a.�󏭐��� "
wSQL = wSQL & "    , a.�Z�b�g���i�t���O "
wSQL = wSQL & "    , a.���[�J�[�������敪 "
wSQL = wSQL & "    , a.Web�[����\���t���O "
wSQL = wSQL & "    , a.���ח\�薢��t���O "
wSQL = wSQL & "    , a.B�i�t���O "
wSQL = wSQL & "    , a.�����萔�� "
wSQL = wSQL & "    , a.������󒍍ϐ��� "
wSQL = wSQL & "    , b.���[�J�[�� "
wSQL = wSQL & "    , d.�����\���� "
wSQL = wSQL & "    , d.�������� "
wSQL = wSQL & "    , d.�����\���ח\��� "
wSQL = wSQL & "    , d.B�i�����\���� "
wSQL = wSQL & "    , a.�J�e�S���[�R�[�h "			'2012/08/27 ok Add
wSQL = wSQL & "    , e.�J�e�S���[�� "				'2012/08/27 ok Add
wSQL = wSQL & "FROM "
wSQL = wSQL & "    �J�e�S���[�ʃI�v�V����       c WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web���i         a WITH (NOLOCK) "
wSQL = wSQL & "        ON     a.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h "
wSQL = wSQL & "           AND a.���i�R�[�h     = c.�I�v�V�������i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���[�J�[�R�[�h = c.�I�v�V�������[�J�[�R�[�h "
wSQL = wSQL & "           AND d.���i�R�[�h     = c.�I�v�V�������i�R�[�h "
wSQL = wSQL & "           AND d.�F             = c.�I�v�V�����F "
wSQL = wSQL & "           AND d.�K�i           = c.�I�v�V�����K�i "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O    = t1.ShohinWebY "
wSQL = wSQL & "      INNER JOIN �J�e�S���[ e WITH (NOLOCK) "				'2012/08/27 ok Add
wSQL = wSQL & "        ON     a.�J�e�S���[�R�[�h = e.�J�e�S���[�R�[�h "		'2012/08/27 ok Add
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND c.�J�e�S���[�R�[�h = '" & wCategoryCode & "' "

wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      a.�J�e�S���[�R�[�h "		'2012/08/27 ok Add
wSQL = wSQL & "    , b.���[�J�[�� "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , d.�F "
wSQL = wSQL & "    , d.�K�i "
' 2012/01/18 GV Mod End

'@@@@response.write(wSQL)

'2012/07/10 GV Mod Start
'call CreateOptionPartsHTML("�I�v�V����")
call CreateOptionPartsHTML("�֘A�I�v�V����")
'2012/07/10 GV Mod End

wOptionHTML = wHTML

End Function

'========================================================================
'
'	Function	�p�[�c HTML�i�f�[�^���o�j
'
'		�F�K�i�Ɋ֌W�Ȃ��Y�����i�̃p�[�c�����o��
'
'========================================================================
'
Function CreatePartsHtml()

'---- Select �p�[�c
wSQL = ""
' 2012/01/18 GV Mod Start
'wSQL = wSQL & "SELECT c.�p�[�c���[�J�[�R�[�h AS ���[�J�[�R�[�h"
'wSQL = wSQL & "     , c.�p�[�c���i�R�[�h AS ���i�R�[�h"
'wSQL = wSQL & "     , d.�F AS �F"
'wSQL = wSQL & "     , d.�K�i AS �K�i"
'wSQL = wSQL & "     , a.���i��"
''wSQL = wSQL & "     , a.�̔��P��"   '2010/11/10 an del
'
''2010/11/10 an add s
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B�i�t���O = 'Y' THEN a.B�i�P��"
'wSQL = wSQL & "         WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS ���̔��P��"
''2010/11/10 an add e
'
'wSQL = wSQL & "     , a.���i�T��Web"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , a.�戵���~��"
'wSQL = wSQL & "     , a.�p�ԓ�"
'wSQL = wSQL & "     , a.������"
'wSQL = wSQL & "     , a.�󏭐���"
'wSQL = wSQL & "     , a.�Z�b�g���i�t���O"
'wSQL = wSQL & "     , a.���[�J�[�������敪"
'wSQL = wSQL & "     , a.Web�[����\���t���O"
'wSQL = wSQL & "     , a.���ח\�薢��t���O"
'wSQL = wSQL & "     , a.B�i�t���O"
'wSQL = wSQL & "     , a.�����萔��"
'wSQL = wSQL & "     , a.������󒍍ϐ���"
'wSQL = wSQL & "     , b.���[�J�[��"
'wSQL = wSQL & "     , d.�����\����"
'wSQL = wSQL & "     , d.��������"			'2011/06/09 hn add
'wSQL = wSQL & "     , d.�����\���ח\���"
'wSQL = wSQL & "     , d.B�i�����\����"
'wSQL = wSQL & "  FROM Web���i a"
'wSQL = wSQL & "     , ���[�J�[ b"
'wSQL = wSQL & "     , �p�[�c c"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d"
'wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.���i�R�[�h = c.�p�[�c���i�R�[�h"
'wSQL = wSQL & "   AND b.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = c.�p�[�c���i�R�[�h"
'wSQL = wSQL & "   AND d.�F = c.�p�[�c�F"
'wSQL = wSQL & "   AND d.�K�i = c.�p�[�c�K�i"
'wSQL = wSQL & "   AND c.���[�J�[�R�[�h = '" & maker_cd & "'"
'wSQL = wSQL & "   AND c.���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'
'wSQL = wSQL & " UNION "
'
'wSQL = wSQL & "SELECT c.�p�[�c���[�J�[�R�[�h AS ���[�J�[�R�[�h"
'wSQL = wSQL & "     , c.�p�[�c���i�R�[�h AS ���i�R�[�h"
'wSQL = wSQL & "     , d.�F AS �F"
'wSQL = wSQL & "     , d.�K�i AS �K�i"
'wSQL = wSQL & "     , a.���i��"
''wSQL = wSQL & "     , a.�̔��P��"   '2010/11/10 an del
'
''2010/11/10 an add s
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B�i�t���O = 'Y' THEN a.B�i�P��"
'wSQL = wSQL & "         WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS ���̔��P��"
''2010/11/10 an add e
'
'wSQL = wSQL & "     , a.���i�T��Web"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , a.�戵���~��"
'wSQL = wSQL & "     , a.�p�ԓ�"
'wSQL = wSQL & "     , a.������"
'wSQL = wSQL & "     , a.�󏭐���"
'wSQL = wSQL & "     , a.�Z�b�g���i�t���O"
'wSQL = wSQL & "     , a.���[�J�[�������敪"
'wSQL = wSQL & "     , a.Web�[����\���t���O"
'wSQL = wSQL & "     , a.���ח\�薢��t���O"
'wSQL = wSQL & "     , a.B�i�t���O"
'wSQL = wSQL & "     , a.�����萔��"
'wSQL = wSQL & "     , a.������󒍍ϐ���"
'wSQL = wSQL & "     , b.���[�J�[��"
'wSQL = wSQL & "     , d.�����\����"
'wSQL = wSQL & "     , d.��������"		'2011/06/09 hn add
'wSQL = wSQL & "     , d.�����\���ח\���"
'wSQL = wSQL & "     , d.B�i�����\����"
'wSQL = wSQL & "  FROM Web���i a"
'wSQL = wSQL & "     , ���[�J�[ b"
'wSQL = wSQL & "     , �J�e�S���[�ʃp�[�c c"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d"
'wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.���i�R�[�h = c.�p�[�c���i�R�[�h"
'wSQL = wSQL & "   AND b.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = c.�p�[�c���i�R�[�h"
'wSQL = wSQL & "   AND d.�F = c.�p�[�c�F"
'wSQL = wSQL & "   AND d.�K�i = c.�p�[�c�K�i"
'wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = '" & wCategoryCode & "'"
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       b.���[�J�[��"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , d.�F"
'wSQL = wSQL & "     , d.�K�i"
wSQL = wSQL & "SELECT "
wSQL = wSQL & "      c.�p�[�c���[�J�[�R�[�h AS ���[�J�[�R�[�h "
wSQL = wSQL & "    , c.�p�[�c���i�R�[�h AS ���i�R�[�h "
wSQL = wSQL & "    , d.�F AS �F "
wSQL = wSQL & "    , d.�K�i AS �K�i "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B�i�t���O = 'Y'                     THEN a.B�i�P�� "
wSQL = wSQL & "        WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P�� "
wSQL = wSQL & "        ELSE                                            a.�̔��P�� "
wSQL = wSQL & "      END AS ���̔��P�� "
wSQL = wSQL & "    , a.���i�T��Web "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.ASK���i�t���O "
wSQL = wSQL & "    , a.�戵���~�� "
wSQL = wSQL & "    , a.�p�ԓ� "
wSQL = wSQL & "    , a.������ "
wSQL = wSQL & "    , a.�󏭐��� "
wSQL = wSQL & "    , a.�Z�b�g���i�t���O "
wSQL = wSQL & "    , a.���[�J�[�������敪 "
wSQL = wSQL & "    , a.Web�[����\���t���O "
wSQL = wSQL & "    , a.���ח\�薢��t���O "
wSQL = wSQL & "    , a.B�i�t���O "
wSQL = wSQL & "    , a.�����萔�� "
wSQL = wSQL & "    , a.������󒍍ϐ��� "
wSQL = wSQL & "    , b.���[�J�[�� "
wSQL = wSQL & "    , d.�����\���� "
wSQL = wSQL & "    , d.�������� "
wSQL = wSQL & "    , d.�����\���ח\��� "
wSQL = wSQL & "    , d.B�i�����\���� "
wSQL = wSQL & "    , a.�J�e�S���[�R�[�h "			'2012/08/27 ok Add
wSQL = wSQL & "    , e.�J�e�S���[�� "				'2012/08/27 ok Add
wSQL = wSQL & "FROM "
wSQL = wSQL & "    �p�[�c                       c WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web���i         a WITH (NOLOCK) "
wSQL = wSQL & "        ON     a.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h "
wSQL = wSQL & "           AND a.���i�R�[�h     = c.�p�[�c���i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h "
wSQL = wSQL & "           AND d.���i�R�[�h     = c.�p�[�c���i�R�[�h "
wSQL = wSQL & "           AND d.�F             = c.�p�[�c�F "
wSQL = wSQL & "           AND d.�K�i           = c.�p�[�c�K�i "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O    = t1.ShohinWebY "
wSQL = wSQL & "      INNER JOIN �J�e�S���[ e WITH (NOLOCK) "				'2012/08/27 ok Add
wSQL = wSQL & "        ON     a.�J�e�S���[�R�[�h = e.�J�e�S���[�R�[�h "		'2012/08/27 ok Add
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND c.���[�J�[�R�[�h = '" & maker_cd & "' "
wSQL = wSQL & "    AND c.���i�R�[�h     = '" & Replace(product_cd, "'", "''") & "' "	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)

wSQL = wSQL & " UNION "

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      c.�p�[�c���[�J�[�R�[�h AS ���[�J�[�R�[�h "
wSQL = wSQL & "    , c.�p�[�c���i�R�[�h AS ���i�R�[�h "
wSQL = wSQL & "    , d.�F AS �F "
wSQL = wSQL & "    , d.�K�i AS �K�i "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B�i�t���O = 'Y'                     THEN a.B�i�P�� "
wSQL = wSQL & "        WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P�� "
wSQL = wSQL & "        ELSE                                            a.�̔��P�� "
wSQL = wSQL & "      END AS ���̔��P�� "
wSQL = wSQL & "    , a.���i�T��Web "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , a.ASK���i�t���O "
wSQL = wSQL & "    , a.�戵���~�� "
wSQL = wSQL & "    , a.�p�ԓ� "
wSQL = wSQL & "    , a.������ "
wSQL = wSQL & "    , a.�󏭐��� "
wSQL = wSQL & "    , a.�Z�b�g���i�t���O "
wSQL = wSQL & "    , a.���[�J�[�������敪 "
wSQL = wSQL & "    , a.Web�[����\���t���O "
wSQL = wSQL & "    , a.���ח\�薢��t���O "
wSQL = wSQL & "    , a.B�i�t���O "
wSQL = wSQL & "    , a.�����萔�� "
wSQL = wSQL & "    , a.������󒍍ϐ��� "
wSQL = wSQL & "    , b.���[�J�[�� "
wSQL = wSQL & "    , d.�����\���� "
wSQL = wSQL & "    , d.�������� "
wSQL = wSQL & "    , d.�����\���ח\��� "
wSQL = wSQL & "    , d.B�i�����\���� "
wSQL = wSQL & "    , a.�J�e�S���[�R�[�h "		'2012/08/27 ok Add
wSQL = wSQL & "    , e.�J�e�S���[�� "			'2012/08/27 ok Add
wSQL = wSQL & "FROM "
wSQL = wSQL & "    �J�e�S���[�ʃp�[�c           c WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN Web���i         a WITH (NOLOCK) "
wSQL = wSQL & "        ON     a.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h "
wSQL = wSQL & "           AND a.���i�R�[�h     = c.�p�[�c���i�R�[�h "
wSQL = wSQL & "      INNER JOIN ���[�J�[        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���[�J�[�R�[�h = c.�p�[�c���[�J�[�R�[�h "
wSQL = wSQL & "           AND d.���i�R�[�h     = c.�p�[�c���i�R�[�h "
wSQL = wSQL & "           AND d.�F             = c.�p�[�c�F "
wSQL = wSQL & "           AND d.�K�i           = c.�p�[�c�K�i "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O    = t1.ShohinWebY "
wSQL = wSQL & "      INNER JOIN �J�e�S���[ e WITH (NOLOCK) "				'2012/08/27 ok Add
wSQL = wSQL & "        ON     a.�J�e�S���[�R�[�h = e.�J�e�S���[�R�[�h "		'2012/08/27 ok Add
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND c.�J�e�S���[�R�[�h = '" & wCategoryCode & "' "

wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      a.�J�e�S���[�R�[�h "			'2012/08/27 ok Add
wSQL = wSQL & "    , b.���[�J�[�� "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , d.�F "
wSQL = wSQL & "    , d.�K�i "
' 2012/01/18 GV Mod End

'@@@@@response.write(wSQL)

'2012/07/10 GV Mod Start
'call CreateOptionPartsHTML("�p�[�c")
call CreateOptionPartsHTML("�֘A�p�[�c")
'2012/07/10 GV Mod End

wPartsHtml = wHTML

End Function

'========================================================================
'
'	Function	�I�v�V�����A�p�[�c HTML�쐬�i���ʁj
'
'	Parm: pTitle(�^�C�g��)
'
'========================================================================
'
Function CreateOptionPartsHTML(pTitle)

Dim RSv
Dim vInventoryCD
Dim vInventoryImage
Dim vProdTermFl		'2010/12/28 hn add
Dim i
Dim j
Dim vCategoryCode			'2012/08/27 ok Add

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wHTML = ""

if RSv.EOF = false then
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "<table width='602' border='0' cellspacing='0' cellpadding='0' id='Shop_Option_Parts_title'>" & vbNewLine
'	wHTML = wHTML & "<form name='fBuyTogether' method='post'>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td align='left'>&nbsp;<b>" & pTitle & "</b></td>" & vbNewLine
'	wHTML = wHTML & "    <td align='left'><a href='#top'><img src='images/goes_up.gif' width='18' height='18' border='0' align='right'></a></td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine
'	wHTML = wHTML & "<table width='602' border='1' cellpadding='0' cellspacing='0' id='Shop_Option_Parts_Frame'>" & vbNewLine
	wHTML = wHTML & "<h2 class='detail_title'>" & pTitle & "</h2>" & vbNewLine
	wHTML = wHTML & "<form name='fBuyTogether' method='post'>" & vbNewLine
	'2012/07/10 GV Mod End

	vCategoryCode = ""		'2012/08/27 ok Add
	Do While RSv.EOF = false
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "  <tr>" & vbNewLine
'2012/08/27 ok Add Start
		if vCategoryCode = "" And wOptionPartsTitleFlag = "Y" Then
			wHTML = wHTML & "  <div class='headline'>" & vbNewLine
			wHTML = wHTML & "    <h3>" & RSv("�J�e�S���[��") & "</h3>" & vbNewLine
			wHTML = wHTML & "  </div>" & vbNewLine
			vCategoryCode = RSv("�J�e�S���[�R�[�h")
		End If
'2012/08/27 ok Add End

		wHTML = wHTML & "<ul class='relation'>" & vbNewLine
		'2012/07/10 GV Mod End

		'2012/07/10 GV Mod Start
'		For i=1 To 5
		For i=1 To 4
		'2012/07/10 GV Mod End
			'---- �p�ԃ`�F�b�N
			if  (isNull(RSv("�戵���~��")) = true AND isNull(RSv("�p�ԓ�")) = true) _
			 OR (isNull(RSv("�p�ԓ�")) = false AND (RSv("�����\����") > 0 OR RSv("��������") > 0)) _
			 OR (isNull(RSv("������")) = false) then		'2011/06/09 hn mod
				vProdTermFl = "N"		'2010/12/28 hn mod
			else
				vProdTermFl = "Y"		'2010/12/28 hn mod
			end if

			'2012/07/10 GV Del Start
'			wHTML = wHTML & "    <td>" & vbNewLine
'			wHTML = wHTML & "      <table border='0' cellspacing='0' cellpadding='0' id='Shop_Option_Parts_product'>" & vbNewLine
'			wHTML = wHTML & "        <tr>" & vbNewLine
			'2012/07/10 GV Del End

			'---- ���i�摜�A���i��
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "          <td><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h") & "^" & Trim(RSv("�F")) & "^" & Trim(RSv("�K�i"))) & "'><img src='prod_img/" & RSv("���i�摜�t�@�C����_��") & "' width='100' height='50' border='0'><br>" & RSv("���[�J�[��") & "<br>" & RSv("���i��") & "</a><br></td>" & vbNewLine
'			wHTML = wHTML & "        </tr>" & vbNewLine
			wHTML = wHTML & "  <li>" & vbNewLine
			wHTML = wHTML & "    <p><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h") & "^" & Trim(RSv("�F")) & "^" & Trim(RSv("�K�i"))) & "'>"
			If RSv("���i�摜�t�@�C����_��") <> "" Then
				wHTML = wHTML & "<img src='prod_img/" & RSv("���i�摜�t�@�C����_��") & "' alt='" & Replace(RSv("���[�J�[��") & " " & RSv("���i��"),"'","&#39;") & "' class='opover'>"
			Else
				wHTML = wHTML & "<img src=""prod_img/n/nopict-.jpg"" alt="""">"
			End If
			wHTML = wHTML & RSv("���[�J�[��") & " / " & RSv("���i��") & "</a></p>" & vbNewLine
			'2012/07/10 GV Mod End

			wHTML = wHTML & "    <div class='box'>" & vbNewLine	'2012/07/10 GV Add
			'----- �̔��P��
			wPrice = calcPrice(RSv("���̔��P��"), wSalesTaxRate)  '2010/11/10 an mod
'			wHTML = wHTML & "        <tr>" & vbNewLine	'2012/07/10 GV Del
			if RSv("ASK���i�t���O") = "Y" then
'2011/10/19 hn mod s
'				wHTML = wHTML & "          <td>ASK</td>" & vbNewLine
				'2012/07/10 GV Mod Start
'				wHTML = wHTML & "          <td><a class='tip'>ASK<span>"  & FormatNumber(wPrice,0) & "�~(�ō�)</span></a></td>" & vbNewLine
'2014/03/19 GV mod start ---->
'				wHTML = wHTML & "      <p><a class='tip'>ASK<span>"  & FormatNumber(wPrice,0) & "�~(�ō�)</span></a></p>" & vbNewLine
				wHTML = wHTML & "      <p><a class='tip'>ASK<span class='exc-tax'>"  & FormatNumber(RSv("���̔��P��"),0) & "�~(�Ŕ�)</span><br>"
				wHTML = wHTML & "      <span class='inc-tax'>(�ō�&nbsp;"  & FormatNumber(wPrice,0) & "�~)</span></a></p>" & vbNewLine
'2014/03/19 GV mod end <-----
				'2012/07/10 GV Mod End
'2011/10/19 hn mod e

			else
				'2012/07/10 GV Mod Start
'				wHTML = wHTML & "          <td>" & FormatNumber(wPrice,0) & "�~(�ō�)</td>" & vbNewLine
'2014/03/19 GV mod start ---->
'				wHTML = wHTML & "      <p>"  & FormatNumber(wPrice,0) & "�~(�ō�)</p>" & vbNewLine
				wHTML = wHTML & "      <p>"  & FormatNumber(RSv("���̔��P��"),0) & "�~(�Ŕ�)</p>" & vbNewLine
				wHTML = wHTML & "      <p>(�ō�&nbsp;"  & FormatNumber(wPrice,0) & "�~)</p>" & vbNewLine
'2014/03/19 GV mod end <-----
				'2012/07/10 GV Mod End
			end if
'			wHTML = wHTML & "        </tr>" & vbNewLine	'2012/07/10 GV Del

			'----- �݌ɏ�
			vInventoryCd = GetInventoryStatus(RSv("���[�J�[�R�[�h"),RSv("���i�R�[�h"),RSv("�F"),RSv("�K�i"),RSv("�����\����"),RSv("�󏭐���"),RSv("�Z�b�g���i�t���O"),RSv("���[�J�[�������敪"),RSv("�����\���ח\���"),vProdTermFl)  		'2010/12/28 hn mod

			'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
			call GetInventoryStatus2(RSv("�����\����"), RSv("Web�[����\���t���O"), RSv("���ח\�薢��t���O"), RSv("�p�ԓ�"), RSv("B�i�t���O"), RSv("B�i�����\����"), RSv("�����萔��"), RSv("������󒍍ϐ���"), vProdTermFl, vInventoryCd, vInventoryImage)		'2010/12/28 hn mod

			'----
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "        <tr>" & vbNewLine
'			wHTML = wHTML & "          <td><img src='images/" & vInventoryImage & "' width='10' height='10'> " & vInventoryCd & "</td>" & vbNewLine
'			wHTML = wHTML & "        </tr>" & vbNewLine
			wHTML = wHTML & "      <p class='stock'><img src='images/" & vInventoryImage & "' alt='" & vInventoryCd & "'>" & vInventoryCd & "</p>" & vbNewLine
			'2012/07/10 GV Mod End

		'----- �ꏏ�ɍw������
'			wHTML = wHTML & "        <tr>" & vbNewLine	'2012/07/10 GV Del

			if vInventoryCd = "�戵���~" then
				'2012/07/10 GV Mod Start
'				wHTML = wHTML & "          <td class='prod_cart'>&nbsp;</td>" & vbNewLine
				wHTML = wHTML & "      <p class='together'>&nbsp;</p>" & vbNewLine
				'2012/07/10 GV Mod End
			else
				'2012/07/10 GV Mod Start
'				wHTML = wHTML & "          <td class='prod_cart'><input type='checkbox' name='iBuyTogether' value='" & RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h") & "^" & Trim(RSv("�F")) & "^" & Trim(RSv("�K�i")) & "' id='checkbox' onClick='BuyTogether_onClick(this);'>�ꏏ�ɍw������</td>" & vbNewLine
				wHTML = wHTML & "      <p class='together'><input type='checkbox' name='iBuyTogether' value='" & RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h") & "^" & Trim(RSv("�F")) & "^" & Trim(RSv("�K�i")) & "' onClick='BuyTogether_onClick(this);'>�ꏏ�ɍw��</p>" & vbNewLine
				'2012/07/10 GV Mod End

			end if

			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "        </tr>" & vbNewLine
'			wHTML = wHTML & "      </table>" & vbNewLine
'			wHTML = wHTML & "    </td>" & vbNewLine
			wHTML = wHTML & "    </div>" & vbNewLine
			wHTML = wHTML & "  </li>" & vbNewLine
			'2012/07/10 GV Mod End

			RSv.MoveNext

			'---- 1�s5���׈ȓ��̎��͋󖾍ׂ����
			if RSv.EOF = true then
				'2012/07/10 GV Del Start
'				For j=i+1 to 5
'					wHTML = wHTML & "    <td>" & vbNewLine
'					wHTML = wHTML & "      <table border='0' cellspacing='0' cellpadding='0' id='Shop_Option_Parts_product'>" & vbNewLine
'					wHTML = wHTML & "        <tr>" & vbNewLine
'					wHTML = wHTML & "          <td>&nbsp;</td>" & vbNewLine
'					wHTML = wHTML & "        </tr>" & vbNewLine
'					wHTML = wHTML & "      </table>" & vbNewLine
'					wHTML = wHTML & "    </td>" & vbNewLine
'				Next
				'2012/07/10 GV Del End
				i = 5
'2012/08/27 ok Add Start
			else
				if vCategoryCode <> RSv("�J�e�S���[�R�[�h") And wOptionPartsTitleFlag = "Y" Then
					vCategoryCode = ""
					i = 5
				end if
'2012/08/27 ok Add End
			end if
		Next

		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "  </tr>" & vbNewLine
		wHTML = wHTML & "</ul>" & vbNewLine
		'2012/07/10 GV Mod End
	Loop

	wHTML = wHTML & "</form>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine	'2012/07/10 GV Del

	wOptionPartsFl = true
end if

RSv.Close

End Function

'========================================================================
'
'	Function	�J�X�^�}�[���r���[�A�]�� HTML�쐬
'
'========================================================================
'
Function CreateReviewHTML()

Dim vAvgRating
Dim v1Cnt
Dim v0Cnt
Dim vHalfCnt
Dim vTotalCnt
Dim vOnpu
Dim RSv
Dim i

'---- Select ���i���r���[ ���ρC���� �擾
wSQL = ""
' 2012/01/23 GV Mod Start
'wSQL = wSQL & "SELECT SUM(a.�]��) AS �]�����v"
'wSQL = wSQL & "     , COUNT(a.ID) AS ���r���[��"
'wSQL = wSQL & "  FROM ���i���r���[ a WITH (NOLOCK) "				' 2012/01/18 GV Mod  WITH (NOLOCK)�t��
'wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = '" & maker_cd & "'"
'wSQL = wSQL & "   AND a.���i�R�[�h = '" & product_cd & "'"
'
''@@@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'if RSv("���r���[��") = 0 then
'	exit function
'end if
'
'vAvgRating = Round(RSv("�]�����v")/RSv("���r���[��"), 1)
'v1Cnt = Fix(vAvgRating)
'if (vAvgRating - v1Cnt) >= 0.5 then
'	vHalfCnt = 1
'else
'	vHalfCnt = 0
'end if
'v0Cnt = 5 - v1Cnt - vHalfCnt
'
'vTotalCnt = RSv("���r���[��")
'Rsv.Close

wSQL = wSQL & "SELECT "
wSQL = wSQL & "      a.���r���[�]������ "
wSQL = wSQL & "    , a.���r���[���� "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ���i���r���[�W�v a WITH (NOLOCK) "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        a.���[�J�[�R�[�h = '" & maker_cd & "' "
wSQL = wSQL & "    AND a.���i�R�[�h     = '" & Replace(product_cd, "'", "''") & "' "

'@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

If RSv.EOF Then
	RSv.Close
	Set RSv = Nothing
	Exit Function
End If

vAvgRating = Round(RSv("���r���[�]������"), 1)
vTotalCnt = RSv("���r���[����")

Rsv.Close
Set RSv = Nothing

If vTotalCnt = 0 Then
	Exit Function
End If

v1Cnt = Fix(vAvgRating)
If (vAvgRating - v1Cnt) >= 0.5 Then
	vHalfCnt = 1
Else
	vHalfCnt = 0
End If
v0Cnt = 5 - v1Cnt - vHalfCnt
' 2012/01/23 GV Mod End

'---- �����摜�쐬
vOnpu = ""
For i=1 to v1Cnt
	'2012/07/10 GV Mod Start
'	vOnpu = vOnpu & "<img src='images/onpu1.jpg' width='20' height='18'>"
	vOnpu = vOnpu & "<img src='images/review_icon10.png' alt='1'>"
	'2012/07/10 GV Mod End
Next
if vHalfcnt = 1 then
	'2012/07/10 GV Mod Strat
'	vOnpu = vOnpu & "<img src='images/onpuHalf.jpg' width='20' height='18'>"
	vOnpu = vOnpu & "<img src='images/review_icon05.png' alt='0.5'>"
	'2012/07/10 GV Mod End
end if
For i=1 to v0Cnt
	'2012/07/10 GV Mod Start
'	vOnpu = vOnpu & "<img src='images/onpu0.jpg' width='20' height='18'>"
	vOnpu = vOnpu & "<img src='images/review_icon00.png' alt='0'>"
	'2012/07/10 GV Mod End
Next

wHTML = ""

'---- �]���ҏW
'2012/07/10 GV Mod Start
'wHTML = wHTML & "<table width='188' border='0' cellspacing='0' cellpadding='0' id='Shop_right'>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td class='head'>�]��</td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td class='base' itemprop='review' itemscope itemtype='http://data-vocabulary.org/Review-aggregate'>" & vbNewLine        '2011/11/22 an mod
'wHTML = wHTML & "      <table width='180' border='0' cellspacing='0' cellpadding='0' id='Shop_right_product'>" & vbNewLine
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left' width='80'>�������ߓx�F</td>" & vbNewLine
'wHTML = wHTML & "          <td align='left' width='100'><span itemprop='rating'>" & FormatNumber(vAvgRating,1) & "</span></td>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "        </tr>" & vbNewLine
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td colspan='2' align='center' height='26'>" & vOnpu & "</td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left'>���r���[���F</td>" & vbNewLine
'wHTML = wHTML & "          <td align='left'><span itemprop='count'>" & vTotalCnt & "</span></td>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "        </tr>" & vbNewLine
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td colspan='2' align='center' height='26'><a href='#review'><img src='images/Reviews.gif' border='0'></a></td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'wHTML = wHTML & "      </table>" & vbNewLine

'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
wHTML = wHTML & "<div class='review'>" & vbNewLine
wHTML = wHTML & "  <p><strong>�]���F</strong>" & vOnpu & "</p>" & vbNewLine
wHTML = wHTML & "  <p><a href='#review'>���r���[���F" & vTotalCnt & "</a></p>" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Mod End

wHyoukaHTML = wHTML

'----��������J�X�^�}�[���r���[ ===========
'---- �����]���ҏW
wHTML = ""
'2012/07/10 GV Mod Start
'wHTML = wHTML & "<table width='602' height='50' border='0' cellspacing='0' cellpadding='0' id='Shop_review_head'>" & vbNewLine

'---- �������ߓx
'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td width='80' align='center'>�����]��</td>" & vbNewLine
'wHTML = wHTML & "    <td width='110' nowrap>" & vOnpu & "</td>" & vbNewLine
'wHTML = wHTML & "    <td width='50' nowrap><b>(" & FormatNumber(vAvgRating,1) & ")</b></td>" & vbNewLine
wHTML = wHTML & "<div class='comment_box'>" & vbNewLine
wHTML = wHTML & "  <ul id='totalreview' itemprop='review' itemscope itemtype='http://data-vocabulary.org/Review-aggregate'>" & vbNewLine
wHTML = wHTML & "    <li><span class='review_icon'>�����]���F" & vOnpu & "<span itemprop='rating'>(" & FormatNumber(vAvgRating,1) & ")</span></span></li>" & vbNewLine

'---- ���r���[��
'wHTML = wHTML & "    <td nowrap><b>���r���[���F " & vTotalCnt & "</b></td>" & vbNewLine
'wHTML = wHTML & "    <td align='left'><a href='#top'><img src='images/goes_up.gif' width='18' height='18' border='0' align='right'></a></td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
wHTML = wHTML & "    <li>���r���[���F<span itemprop='count'>" & vTotalCnt & "</span></li>" & vbNewLine
wHTML = wHTML & "  </ul>" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine

'wHTML = wHTML & "</table>" & vbNewLine
'2012/07/10 GV Mod End

'---- Select �ʏ��i���r���[
wSQL = ""
if ReviewAll = "Y" then
	wSQL = wSQL & "SELECT "
else
	wSQL = wSQL & "SELECT TOP 5 "
end if
wSQL = wSQL & "      a.ID "
wSQL = wSQL & "    , a.���e�� "
wSQL = wSQL & "    , a.�]�� "
wSQL = wSQL & "    , a.�^�C�g�� "
wSQL = wSQL & "    , a.���O "
wSQL = wSQL & "    , a.���r���[���e "
wSQL = wSQL & "    , a.�Q�l�� "
wSQL = wSQL & "    , a.�s�Q�l�� "
wSQL = wSQL & "    , a.�ڋq�ԍ� "
wSQL = wSQL & "    , a.�V���b�v�R�����g�� "
wSQL = wSQL & "    , a.�V���b�v�R�����g�^�C�g�� "
wSQL = wSQL & "    , a.�V���b�v�R�����g "
wSQL = wSQL & "    , b.�ڋq�s���{�� "
' 2012/01/18 GV Mod Start (WITH (NOLOCK)�t��)
'wSQL = wSQL & "  FROM ���i���r���[ a LEFT JOIN Web�ڋq�Z�� b"
'wSQL = wSQL & "                             ON b.�ڋq�ԍ� = a.�ڋq�ԍ�"
'wSQL = wSQL & "                            AND b.�Z���A�� = 1"
wSQL = wSQL & "FROM "
wSQL = wSQL & "    ���i���r���[            a WITH (NOLOCK) "
wSQL = wSQL & "      LEFT JOIN Web�ڋq�Z�� b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.�ڋq�ԍ� = a.�ڋq�ԍ� "
wSQL = wSQL & "           AND b.�Z���A�� = 1 "
' 2012/01/18 GV Mod End
wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)
wSQL = wSQL & " ORDER BY"
' 2012/01/18 GV Mod Start
'wSQL = wSQL & "       a.���e�� DESC"
wSQL = wSQL & "       a.ID DESC"
' 2012/01/18 GV Mod End

'@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

'wHTML = wHTML & "<table width='602' border='0' cellspacing='0' cellpadding='5' id='shop_border'>" & vbNewLine

'---- �ʃ��r���[�ҏW
Do While RSv.EOF = false

	'2012/07/10 GV Mod Start
'---- �������ߓx
'	wHTML = wHTML & "  <tr class='honbun'>" & vbNewLine
'	wHTML = wHTML & "    <td width='130' nowrap class='Shop_review_th Shop_review_th_onpu'>"
'	For i=1 to RSv("�]��")
'		wHTML = wHTML & "<img src='images/onpu1.jpg' width='20' height='18'>"
'	Next
'	For i=RSv("�]��")+1 to 5
'		wHTML = wHTML & "<img src='images/onpu0.jpg' width='20' height='18'>"
'	Next
'	wHTML = wHTML & " (" & FormatNumber(RSv("�]��"), 1) & ")" & vbNewLine
'	wHTML = wHTML & "</td>" & vbNewLine

'---- �^�C�g��, ���e��
'	wHTML = wHTML & "    <td width='400' class='Shop_review_th'><h3>" & RSv("�^�C�g��") & "</h3></td>" & vbNewLine
'	wHTML = wHTML & "    <td align='right' nowrap class='Shop_review_th Shop_review_th_right'><span>���r���[ID�F" & RSv("ID") & "</span><br>" & cf_FormatDate(RSv("���e��"), "YYYY/MM/DD") & "</td>" & vbNewLine   '2011/09/09 an mod
'	wHTML = wHTML & "  </tr>" & vbNewLine

'---- ���e���C�������ߓx�C�^�C�g��
	wHTML = wHTML & "<div class='comment_box'>" & vbNewLine
	wHTML = wHTML & "  <p>" & cf_FormatDate(RSv("���e��"), "YYYY/MM/DD") & "</p>" & vbNewLine
	wHTML = wHTML & "  <p class='subject'><span class='review_icon'>"
	For i=1 to RSv("�]��")
		wHTML = wHTML & "<img src='images/review_icon10.png' alt='1'>"
	Next
	wHTML = wHTML & "</span>" & RSv("�^�C�g��") & "</p>" & vbNewLine
	'2012/07/10 GV Mod End
	'2012/07/10 GV Mod Start
'---- ���e�Җ��C�s���{���A���̐l�̃��r���[�����N�A�Q�l�ɂȂ����l��
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td colspan='3' nowrap>" & vbNewLine
'	wHTML = wHTML & "      <table width='100%' cellspacing='0' cellpadding='0'>" & vbNewLine
'	wHTML = wHTML & "        <tr class='honbun'>" & vbNewLine
'	wHTML = wHTML & "          <td nowrap>���e�Җ��F" & RSv("���O")

'	if IsNull(RSv("�ڋq�ԍ�")) = false then
'		if RSv("�ڋq�ԍ�") <> 0 then
'			wHTML = wHTML & " �y" & RSv("�ڋq�s���{��") & "�z <a href='ReviewAllByCustomer.asp?CNo=" & RSv("�ڋq�ԍ�") & "' class='link'><b>���r���[������</b></a>"
'		end if
'	end if

'	wHTML = wHTML & "</td>" & vbNewLine

'	wHTML = wHTML & "          <td align='right' nowrap>�Q�l�ɂȂ����l���F" & RSv("�Q�l��") & "�l(" & RSv("�Q�l��") + RSv("�s�Q�l��") & "�l��)</td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine

'--- ���e�Җ�(���̐l�̃��r���[�����N)�C�s���{��
	if IsNull(RSv("�ڋq�ԍ�")) = false then
		if RSv("�ڋq�ԍ�") <> 0 then
			wHTML = wHTML & "  <p class='postname'>���e�Җ��F<a href='ReviewAllByCustomer.asp?CNo=" & RSv("�ڋq�ԍ�") & "'>" & RSv("���O") & "</a><span>"
			wHTML = wHTML & " �y" & RSv("�ڋq�s���{��") & "�z" & vbNewLine
		end if
	end if
	wHTML = wHTML & "</span></p>"
	'2012/07/10 GV Mod End

'---- ���r���[���e
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "  <tr class='honbun'>" & vbNewLine
'	wHTML = wHTML & "    <td colspan='3' class='Shop_review_text'><p>" & Replace(RSv("���r���[���e"), vbNewline, "<br>") & "</p></td>" & vbNewLine
	wHTML = wHTML & "  <p>" & Replace(RSv("���r���[���e"), vbNewline, "<br>") & "</p>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
	'2012/07/10 GV Mod End

'---- ���r���[���e  2010/03/08 hn changed
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "    <tr>" & vbNewLine
'	wHTML = wHTML & "      <td colspan='3' align='right' class='Shop_review_btn_td'>" & vbNewLine
'	wHTML = wHTML & "        <div class='Shop_review_yn_wrap'>" & vbNewLine
'	wHTML = wHTML & "          <div class='Shop_review_yntxt'>�Q�l�ɂȂ�܂������H</div>" & vbNewLine
'	wHTML = wHTML & "          <div class='Shop_review_ynbtn'>" & vbNewLine
'	wHTML = wHTML & "            <img src='images/btn_yes20.jpg' alt='YES' width='34' height='20' border='0' onMouseOver='this.src=""images/btn_yes20-.jpg"";' onMouseOut='this.src=""images/btn_yes20.jpg"";' onClick='ReviewSankou_onClick(""" & RSv("ID") & """,""" & item & """,""Y"");'>" & vbNewLine
'	wHTML = wHTML & "          </div>" & vbNewLine
'	wHTML = wHTML & "          <div class='Shop_review_slash'>/</div>" & vbNewLine
'	wHTML = wHTML & "          <div class='Shop_review_ynbtn'>" & vbNewLine
'	wHTML = wHTML & "            <img src='images/btn_no20.jpg' alt='NO' width='34' height='20' border='0' onMouseOver='this.src=""images/btn_no20-.jpg"";' onMouseOut='this.src=""images/btn_no20.jpg"";' onClick='ReviewSankou_onClick(""" & RSv("ID") & """,""" & item & """,""N"");'>" & vbNewLine
'	wHTML = wHTML & "          </div>" & vbNewLine
'	wHTML = wHTML & "        </div>" & vbNewLine
'	wHTML = wHTML & "      </td>" & vbNewLine
'	wHTML = wHTML & "    </tr>" & vbNewLine

'2013/05/17 GV #1507 add start
'---- �����̃R�����g�͕ҏW
If (Trim(RSv("�ڋq�ԍ�")) = CStr(wUserID)) Then
	wHTML = wHTML & "  <p id='review_edit'><a href='" & g_HTTPS & "shop/ReviewWrite.asp?Item=" & Server.URLEncode(Item) & "'><img src='images/btn_review_edit.png' alt='���̃��r���[��ҏW����' class='opover'></a></p>"
End If
'2013/05/17 GV #1507 add end

	wHTML = wHTML & "  <div class='review_other'>"	& vbNewLine
	wHTML = wHTML & "    <p class='review_id'>���r���[ID�F" & RSv("ID") & "</p>"	& vbNewLine
	wHTML = wHTML & "    <p>�Q�l�ɂȂ����l���F" & RSv("�Q�l��") & "�l(" & RSv("�Q�l��") + RSv("�s�Q�l��") & "�l��)</p>"	& vbNewLine
	wHTML = wHTML & "    <dl>"	& vbNewLine
	wHTML = wHTML & "      <dt>�Q�l�ɂȂ�܂������H</dt>"	& vbNewLine
	wHTML = wHTML & "      <dd><img src='images/btn_yes20.jpg' alt='Yes' class='opover' onClick='ReviewSankou_onClick(""" & RSv("ID") & """,""" & item & """,""Y"");'></dd>" & vbNewLine
	wHTML = wHTML & "      <dd><img src='images/btn_no20.jpg' alt='NO' class='opover' onClick='ReviewSankou_onClick(""" & RSv("ID") & """,""" & item & """,""N"");'></dd>" & vbNewLine
	wHTML = wHTML & "    </dl>"	& vbNewLine
	wHTML = wHTML & "  </div>"	& vbNewLine
	'2012/07/10 GV Mod End

'---- �V���b�v�R�����g 2010/03/08 an changed
	if IsNull(RSv("�V���b�v�R�����g��")) = false then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td colspan='3' class='Shop_review_text'>" & vbNewLine
'		wHTML = wHTML & "      <div class='Shop_review_sh_res'>" & vbNewLine
'		wHTML = wHTML & "        <div class='Shop_review_sh_res_head'><span>" & RSv("�V���b�v�R�����g�^�C�g��") & "</span> " & cf_FormatDate(RSv("�V���b�v�R�����g��"), "YYYY/MM/DD") & "</div>" & vbNewLine
'		wHTML = wHTML & "        <div class='Shop_review_text'><p>" & Replace(RSv("�V���b�v�R�����g"), vbNewline, "<br>") & "</p></div>" & vbNewLine
'		wHTML = wHTML & "      </div>" & vbNewLine
'		wHTML = wHTML & "    </td>" & vbNewLine
'		wHTML = wHTML & "  </tr>" & vbNewLine
		wHTML = wHTML & "  <div class='reply_box'>" & vbNewLine
		wHTML = wHTML & "    <p>" & cf_FormatDate(RSv("�V���b�v�R�����g��"), "YYYY/MM/DD") & "</p><br>" & vbNewLine
'		wHTML = wHTML & "    <p class='ansewr'>" & RSv("�V���b�v�R�����g�^�C�g��") & "</p>" & vbNewLine
		wHTML = wHTML & "    <p>" & Replace(RSv("�V���b�v�R�����g"), vbNewline, "<br>") & "</p>" & vbNewLine
		wHTML = wHTML & "  </div>" & vbNewLine
		'2012/07/10 GV Mod End
	end if

'2013/05/17 GV #1507 Mod Start
'�g�p���ĂȂ��̂ŃR�����g�A�E�g
'---- �V���b�v�R�����g��������
'	if iShop = "Y" then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "  <tr>" & vbNewLine
'		wHTML = wHTML & "    <td colspan='3'><a href='ReviewShopComment.asp?id=" & RSv("ID") & "' class='link'><b>���V���b�v�R�����g������</b></a></td>" & vbNewLine
'		wHTML = wHTML & "    <a href='ReviewShopComment.asp?id=" & RSv("ID") & "' class='link'><b>���V���b�v�R�����g������</b></a>" & vbNewLine
'		wHTML = wHTML & "  <tr>" & vbNewLine
		'2012/07/10 GV Mod End
'	end if
'2013/05/17 GV #1507 Mod End
	wHTML = wHTML & "</div>" & vbNewLine	'class='comment_box'

'---- ��؂��
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td colspan='3'><hr width='99%' size='1'></td>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine & vbNewLine


	RSv.MoveNext
Loop

'wHTML = wHTML & "</table>" & vbNewLine	'2012/07/10 GV Del

RSv.Close

wHTML = wHTML & "<ul class='btn_review'>" & vbNewLine
'---- �S�Ẵ��r���[������
if ReviewAll <> "Y" AND vTotalCnt > 5 then
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "<table width='602' border='0' cellspacing='0' cellpadding='3'>" & vbNewLine
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td align='right'><a href='ProductDetail.asp?ReviewAll=Y&Item=" & item & "'><img src='images/ReviewAll.gif' border='0'></a></td>" & vbNewLine
	wHTML = wHTML & "  <li><a href='ProductDetail.asp?ReviewAll=Y&Item=" & item & "'><img src='images/btn_review.png' alt='���i���r���[�������ƌ���' class='opover'></a></li>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "</table>" & vbNewLine
	'2012/07/10 GV Mod End
end if

wReviewHTML = wHTML

End Function

'========================================================================
'
'	Function	�o�גʒm����̃����N�Ŏ󒍔ԍ����n���ꂽ�ꍇ�́AUserID���o��
'
'========================================================================
'
Function GetUserID()

Dim RSv

'---- Select Web��
wSQL = ""
wSQL = wSQL & "SELECT a.�ڋq�ԍ�"
wSQL = wSQL & "  FROM Web�� a WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.�󒍔ԍ� = " & OrderNo

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

if RSv.EOF = false then
	wUserID = RSv("�ڋq�ԍ�")
end if

RSv.Close

End Function

'========================================================================
'
'	Function	Check Review �Y���ڋq���w�����т�����A���̏��i�̃��r���[�𓊍e���Ă��邩�ǂ����`�F�b�N
'
'========================================================================
'
Function CheckReview()

Dim RSv

'---- Select ���i���r���[
wSQL = ""
wSQL = wSQL & "SELECT a.�w���� "
wSQL = wSQL & "     , a.�n���h���l�[�� "
wSQL = wSQL & "     , c.�ڋq�s���{�� "
wSQL = wSQL & "     , b.ID "
' 2012/01/18 GV Mod Start ( WITH (NOLOCK)�t�� )
'wSQL = wSQL & "  FROM Web�ڋq a LEFT JOIN ���i���r���[ b"
'wSQL = wSQL & "                        ON b.�ڋq�ԍ� = a.�ڋq�ԍ�"
'wSQL = wSQL & "                       AND b.���[�J�[�R�[�h = '" & maker_cd & "'"
'wSQL = wSQL & "                       AND b.���i�R�[�h = '" & product_cd & "'"
'wSQL = wSQL & "     , Web�ڋq�Z�� c"
'wSQL = wSQL & " WHERE c.�ڋq�ԍ� = a.�ڋq�ԍ�"
'wSQL = wSQL & "   AND c.�Z���A�� = 1"
'wSQL = wSQL & "   AND a.�ڋq�ԍ� = " & wUserID
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web�ڋq                   a WITH (NOLOCK) "
wSQL = wSQL & "      LEFT JOIN  ���i���r���[ b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.�ڋq�ԍ�       = a.�ڋq�ԍ�"
wSQL = wSQL & "           AND b.���[�J�[�R�[�h = '" & maker_cd & "' "
wSQL = wSQL & "           AND b.���i�R�[�h     = '" & Replace(product_cd, "'", "''") & "' "	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)
wSQL = wSQL & "      INNER JOIN Web�ڋq�Z��  c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.�ڋq�ԍ�       = a.�ڋq�ԍ� "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        c.�Z���A�� = 1 "
wSQL = wSQL & "    AND a.�ڋq�ԍ� = " & wUserID
' 2012/01/18 GV Mod End

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

'@@@@@response.write(wSQL)

wPrefecture = ""
wHandleName = ""

if RSv.EOF = false then
	wPrefecture = RSv("�ڋq�s���{��")

	if IsNull(RSv("�n���h���l�[��")) = false then
		wHandleName = Trim(RSv("�n���h���l�[��"))
	end if

	if RSv("�w����") > 0 AND IsNull(RSv("ID")) = true then
		wCanWriteReviewFl = "Y"
	end if
end if

RSv.Close

End Function

'========================================================================
'
'	Function	���[�J�[/���iHTML�쐬
'
'========================================================================
'
Function CreateProductHTML()

Dim vInventoryCd
Dim vInventoryImage
Dim vFreeShippingHTML				' 2011/02/18 GV Add
Dim v_price					' 2012/07/10 GV Add
Dim v_exprice					' 2012/07/10 GV Add
Dim vUrl					' 2012/07/20 GV Add
Dim RSv

wHTML = ""

'---- �^�C�g��
'2012/07/10 GV Mod Start
'wHTML = wHTML & "<table width='188' border='0' cellspacing='0' cellpadding='0' id='Shop_right_Detail'>" & vbNewLine

'wHTML = wHTML & "  <tr>" & vbNewLine
'wHTML = wHTML & "    <td class='head'>���[�J�[/���i��</td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine

'wHTML = wHTML & "    <td class='base' itemprop='offerDetails' itemscope itemtype='http://data-vocabulary.org/Offer'>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "      <table width='180' border='0' cellspacing='0' cellpadding='0' class='ProductDetail'>" & vbNewLine
wHTML = wHTML & "<div id='detail_side_inner01'><div id='detail_side_inner02'>" & vbNewLine

wHTML = wHTML & "  <div id='detail_pp' itemprop='offerDetails' itemscope itemtype='http://data-vocabulary.org/Offer'>" & vbNewLine
'2012/07/10 GV End

'---- ���[�J�[���A���i���A�J�e�S���[��
'2012/07/10 GV Mod Start
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left'><a href='SearchList.asp?i_type=m&s_maker_cd=" & RS("���[�J�[�R�[�h") & "' class='link'>" & RS("���[�J�[��") & " (" & RS("���[�J�[���J�i") & ")</a></td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine

'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left'><span itemprop='itemreviewed'>" & RS("���i��") & "</span></td>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "        </tr>" & vbNewLine

'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='left'><a href='SearchList.asp?i_type=c&s_category_cd=" & RS("�J�e�S���[�R�[�h") & "' class='link'><span itemprop='category'>" & RS("�J�e�S���[��") & "</span></a></td>" & vbNewLine   '2011/11/22 an mod
'wHTML = wHTML & "        </tr>" & vbNewLine

' 2011/02/18 GV Add Start
'---- �������S�������i�̏ꍇ�ɑ}���o�͂��� �y���������L�����y�[�����z �𐶐�
'vFreeShippingHTML = ""
'If wFreeShippingFlag = "Y" Then
'	vFreeShippingHTML = "<br/><strong class='freeshipping'>�y���������L�����y�[�����z<span>������E����������</span></strong>"	'2011/06/15 if-web mod
'End If
' 2011/02/18 GV Add End
' 2012/08/01 ok Mod Start
'wHTML = wHTML & "    <h3 class='item_name'>" & RS("���[�J�[��") & " (" & RS("���[�J�[���J�i") & ")<br>" & RS("���i��") & "<br>" & RS("�J�e�S���[��") & "</h3>" & vbNewLine
wHTML = wHTML & "    <h3 class='item_name'><a href='SearchList.asp?i_type=m&s_maker_cd=" & RS("���[�J�[�R�[�h") & "'>" & RS("���[�J�[��") & " (" & RS("���[�J�[���J�i") & ")</a><br>" & RS("���i��") & "<br><a href ='SearchList.asp?i_type=c&s_category_cd=" & RS("�J�e�S���[�R�[�h") & "'>" & RS("�J�e�S���[��") & "</a></h3>" & vbNewLine
' 2012/08/01 ok Mod End
wHTML = wHTML & "    <p>���iID:" & RS("���iID") & "</p>" & vbNewLine
wHTML = wHTML & "    <ul class='icon_list'>" & vbNewLine
If wFreeShippingFlag = "Y" Then
	wHTML = wHTML & "    <li><img src='images/icon_free.gif' alt='��������'></li>" & vbNewLine
End If
'2012/07/10 GV Mod End
'2012/07/10 GV Add Start
'---- �v���C�X�_�E���̏ꍇ�ɑ}��
If isNULL(RS("�O��P���ύX��")) = False Then
	If DateAdd("d", 60, RS("�O��P���ύX��")) >= Date() AND RS("�O��̔��P��") > RS("�̔��P��") AND RS("�O��̔��P��") <> 0 Then
	wHTML = wHTML & "    <li><img src='images/icon_discount.gif' alt='�l�������܂���'></li>" & vbNewLine
	End If
End If
wHTML = wHTML & "    </ul>" & vbNewLine
'2012/07/10 GV Add End
'---- �̔��P��
v_price = calcPrice(RS("�̔��P��"), wSalesTaxRate)
v_exprice = calcPrice(RS("�O��̔��P��"), wSalesTaxRate)
'1�s�ڂ̕\���iASK���i�ł͂Ȃ��l�����i�̋����i�j
If RS("ASK���i�t���O") <> "Y" Then
	If RS("B�i�t���O") = "Y" OR (RS("�����萔��") > RS("������󒍍ϐ���") AND RS("�����萔��") > 0) OR ( isNULL(RS("�O��P���ύX��")) = False AND DateAdd("d", 60, RS("�O��P���ύX��")) >= Date() AND RS("�O��̔��P��") > RS("�̔��P��") AND RS("�O��̔��P��") <> 0) Then
		'�l�����i�̋����i��\��
		If isNULL(RS("�O��P���ύX��")) = False AND DateAdd("d", 60, RS("�O��P���ύX��")) >= Date() AND RS("�O��̔��P��") > RS("�̔��P��") Then
			wHTML = wHTML & "<p class='cancel'>" & FormatNumber(v_exprice,0) & "�~</p>" & vbNewLine
		'B�i�A����i�͔̔����i�������i�Ƃ��ĕ\��
		Else
			wHTML = wHTML & "<p class='cancel'>" & FormatNumber(v_price,0) & "�~</p>" & vbNewLine
		End If
	End If
End If
'---- �̔��P��
wPrice = calcPrice(RS("�̔��P��"), wSalesTaxRate)
wEAPriceExcTax = FormatNumber(RS("�̔��P��"),0)

'wHTML = wHTML & "        <tr>" & vbNewLine	'2012/07/10 GV Del

if RS("ASK���i�t���O") = "Y" then
	wTwPriceData = "ASK"
'2011/10/19 hn mod s
'	wHTML = wHTML & "          <td>�����F<a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RS("���[�J�[��")) & "&ProductName=" & Server.URLEncode(wProductName) & "&Price=" & wPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'>ASK</a>" & vFreeShippingHTML & "</td>" & vbNewLine
	if RS("B�i�t���O") = "Y" OR (RS("�����萔��") > RS("������󒍍ϐ���") AND RS("�����萔��") > 0) then

		if RS("B�i�t���O") = "Y" then
			wPrice = calcPrice(RS("B�i�P��"), wSalesTaxRate)
			wEAPriceExcTax = FormatNumber(RS("B�i�P��"),0)
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "          <td>�킯����i�����F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a></td>" & vbNewLine
'2014/03/19 GV mod start --->
'			wHTML = wHTML & "          <p class='price'>�킯����i�����F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a></p>" & vbNewLine
			wHTML = wHTML & "          <p class='price'>�킯����i�����F<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("B�i�P��"),0) & "�~(�Ŕ�)</span>"
			wHTML = wHTML & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)</span></a></p>" & vbNewLine
'2014/03/19 GV mod end <-----
			'2012/07/10 GV Mod End
			wTwPriceLabel = "�킯����i����"
		end if

		if (RS("�����萔��") > RS("������󒍍ϐ���") AND RS("�����萔��") > 0) then
			wPrice = calcPrice(RS("������P��"), wSalesTaxRate)
			wEAPriceExcTax = FormatNumber(RS("������P��"),0)
			'2012/07/10 GV Mod
'			wHTML = wHTML & "          <td>��������F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a></td>" & vbNewLine
'2014/03/19 GV mod start --->
'			wHTML = wHTML & "          <p class='price'>��������F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a></p>" & vbNewLine
			wHTML = wHTML & "          <p class='price'>��������F<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("������P��"),0) & "�~(�Ŕ�)</span>"
			wHTML = wHTML & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)</span></a></p>" & vbNewLine
'2014/03/19 GV mod end <-----
			'2012/07/10 GV Mod
			wTwPriceLabel = "�������"
		end if
	else
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "          <td>�����F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a></td>" & vbNewLine
'2014/03/19 GV start ---->
'		wHTML = wHTML & "          <p class='price'>�����F<a class='tip'>ASK<span>" & FormatNumber(wPrice,0) & "�~(�ō�)</span></a></p>" & vbNewLine
		wHTML = wHTML & "          <p class='price'>�����F<a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RS("�̔��P��"),0) & "�~(�Ŕ�)</span>"
		wHTML = wHTML & "<span class='inc-tax'>(�ō�&nbsp;" & FormatNumber(wPrice,0) & "�~)</span></a></p>" & vbNewLine
'2014/03/19 GV end <-----
		'2012/07/10 GV Mod End
		wTwPriceLabel = "�Ռ�����"
	end if

else

	if RS("B�i�t���O") = "Y" OR (RS("�����萔��") > RS("������󒍍ϐ���") AND RS("�����萔��") > 0) then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "            <td><div class='price_table'><del>" & FormatNumber(wPrice,0) & "�~(�ō�)</del><br>" & vbNewLine
'		wHTML = wHTML & "            <p class='price'><strong class='red_large'><meta itemprop='currency' content='JPY' /><span itemprop='price'>"
		wHTML = wHTML & "            <p class='price'><strong class='red_large'>"
		'2012/07/10 GV Mod End

		'---- B�i����
		if RS("B�i�t���O") = "Y" then
			wPrice = calcPrice(RS("B�i�P��"), wSalesTaxRate)
			wEAPriceExcTax = FormatNumber(RS("B�i�P��"),0)

'2014/03/19 GV start ---->
'			wHTML = wHTML & FormatNumber(wPrice,0) & "�~</span></strong><span>(�ō�)</span></p>" & vbNewLine
			wHTML = wHTML & FormatNumber(RS("B�i�P��"),0) & "�~</strong><span>(�Ŕ�)</span></p>" & vbNewLine
			wHTML = wHTML & "<p>(�ō�&nbsp;<meta itemprop='currency' content='JPY' /><span itemprop='price'>" & FormatNumber(wPrice,0) & "�~</span>)</p>" & vbNewLine
'2014/03/19 GV end <-----
' 2011/02/18 GV Mod Start
'			wHTML = wHTML & "            <span class='price'>" & FormatNumber(wPrice,0) & "�~</span><span class='tax'>(�ō�)</span><br><b>�킯����i����</b></div></td>" & vbNewLine '2010/01/26 an �C�� 2010/02/06 if-web �C�� 2010/02/22 st �C��
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "            <meta itemprop='currency' content='JPY' /><span class='price' itemprop='price'>" & FormatNumber(wPrice,0) & "�~</span><span class='tax'>(�ō�)</span><br><b>�킯����i����</b>" & vFreeShippingHTML & "</div></td>" & vbNewLine
			wHTML = wHTML & "            <p class='deals'>�킯����i����</p>" & vbNewLine
			'2012/07/10 GV Mod End
			wTwPriceLabel = "�킯����i����"
' 2011/02/18 GV Mod End   '2011/11/22 an mod

		else
		'---- ������P��
			wPrice = calcPrice(RS("������P��"), wSalesTaxRate)
			wEAPriceExcTax = FormatNumber(RS("������P��"),0)

'2014/03/19 GV start ---->
'			wHTML = wHTML & FormatNumber(wPrice,0) & "�~</span></strong><span>(�ō�)</span></p>" & vbNewLine
			wHTML = wHTML & FormatNumber(RS("������P��"),0) & "�~</strong><span>(�Ŕ�)</span></p>" & vbNewLine
			wHTML = wHTML & "<p>(�ō�&nbsp;<meta itemprop='currency' content='JPY' /><span itemprop='price'>" & FormatNumber(wPrice,0) & "�~</span>)</p>" & vbNewLine
'2014/03/19 GV end <-----
' 2011/02/18 GV Mod Start
'			wHTML = wHTML & "            <span class='price'>" & FormatNumber(wPrice,0) & "�~</span><span class='tax'>(�ō�)</span><br><b>�������</b></div></td>" & vbNewLine
			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "            <meta itemprop='currency' content='JPY' /><span class='price' itemprop='price'>" & FormatNumber(wPrice,0) & "�~</span><span class='tax'>(�ō�)</span><br><b>�������</b>" & vFreeShippingHTML & "</div></td>" & vbNewLine
			wHTML = wHTML & "           <p class='deals'>�������</p>" & vbNewLine
			'2012/07/10 GV Mod End
			wTwPriceLabel = "�������"
' 2011/02/18 GV Mod End   '2011/11/22 an mod
		end if
	else
' 2011/02/18 GV Mod Start
'		wHTML = wHTML & "            <td><div class='price_table'>�����F<span class='price'>" & FormatNumber(wPrice,0) & "�~</span><span class='tax'>(�ō�)</span></div></td>" & vbNewLine '2010/02/06 if-web �C��
		'2012/07/12 GV Mod Start
'		wHTML = wHTML & "            <td><div class='price_table'>�����F<meta itemprop='currency' content='JPY' /><span class='price' itemprop='price'>" & FormatNumber(wPrice,0) & "�~</span><span class='tax'>(�ō�)</span>" & vFreeShippingHTML & "</div></td>" & vbNewLine '2010/02/06 if-web �C��

'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "            <p class='price'><strong class='red_large'><meta itemprop='currency' content='JPY' /><span itemprop='price'>" & FormatNumber(wPrice,0) & "�~</span></strong><span>(�ō�)</span></p>" & vbNewLine '2010/02/06 if-web �C��
		wHTML = wHTML & "            <p class='price'><strong class='red_large'>" & FormatNumber(RS("�̔��P��"),0) & "�~</strong><span>(�Ŕ�)</span></p>" & vbNewLine
		wHTML = wHTML & "            <p>(�ō�&nbsp;<meta itemprop='currency' content='JPY' /><span itemprop='price'>" & FormatNumber(wPrice,0) & "�~</span>)</p>" & vbNewLine
'2014/03/19 GV mod end <----

		'2012/07/12 GV Mod End
		wTwPriceLabel = "�Ռ�����"
' 2011/02/18 GV Mod End   '2011/11/22 an mod
	end if
	wTwPriceData = FormatNumber(wPrice,0) & "�~(�ō�)"
end if

'wHTML = wHTML & "        </tr>" & vbNewLine	'2012/07/12 GV Del

if wIroKikakuSelectedFl = true then
	'----- �݌ɏ�
	vInventoryCd = GetInventoryStatus(RS("���[�J�[�R�[�h"),RS("���i�R�[�h"),RS("�F"),RS("�K�i"),RS("�����\����"),RS("�󏭐���"),RS("�Z�b�g���i�t���O"),RS("���[�J�[�������敪"),RS("�����\���ח\���"),wProdTermFl)

	'---- �݌ɏ󋵁A�F���ŏI�Z�b�g
	call GetInventoryStatus2(RS("�����\����"), RS("Web�[����\���t���O"), RS("���ח\�薢��t���O"), RS("�p�ԓ�"), RS("B�i�t���O"), RS("B�i�����\����"), RS("�����萔��"), RS("������󒍍ϐ���"), wProdTermFl, vInventoryCd, vInventoryImage)

	'----
'2010/11/04 GV Mod Start
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td>�݌ɏ󋵁F<img src='images/" & vInventoryImage & "' width='10' height='10'> " & vInventoryCd & "</td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
	'---- �������łȂ��ꍇ�̂݁A�݌ɏ󋵂�\��
	if (IsNull(RS("�戵���~��")) = false) OR (IsNull(RS("������")) = false) OR (RS("B�i�t���O") = "Y" AND RS("B�i�����\����") <= 0) OR (IsNull(RS("�p�ԓ�")) = false AND RS("�����\����") <= 0 AND RS("��������") <= 0 AND wIroKikakuSelectedFl = true) then	'2011/06/09 hn mod
		wTwInventoryData = "�������܂���"
	else
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "        <tr>" & vbNewLine
'		wHTML = wHTML & "          <td>�݌ɏ󋵁F<img src='images/" & vInventoryImage & "' width='10' height='10'> "    '2011/11/22 an mod s
		wHTML = wHTML & "          <p class='stock'><img src='images/" & vInventoryImage & "' alt='" & vInventoryCd & "'> "    '2011/11/22 an mod s
		'2012/07/10 GV Mod End

		if vInventoryCd = "�݌ɂ���" OR vInventoryCd = "�݌ɋ͏�" OR vInventoryCd = "�݌Ɍ���" OR Left(vInventoryCd, 2) = "����" then
			wHTML = wHTML & "<span itemprop='availability' content='in_stock'>" & vInventoryCd & "</span>"
		else
			wHTML = wHTML & vInventoryCd
		end if                                                                                                          '2011/11/22 an mod e

		wTwInventoryData = vInventoryCd

		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "          </td>" & vbNewLine
'		wHTML = wHTML & "        </tr>" & vbNewLine
		wHTML = wHTML & "          </p>" & vbNewLine
		'2012/07/10 GV Mod End
	end if
'2010/11/04 GV Mod End
end if


'2012/07/10 GV Del Start
'---- �����ɂ���
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td><a href='../guide/kaimono.asp#souryou' class='link'>�����ɂ���</a></td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'2012/07/10 GV Del End

'---- �J�[�g�{�^��
'2012/07/10 GV Mod Start
'wHTML = wHTML & "        <tr>" & vbNewLine
wHTML = wHTML & "          <form name='f_data' method='post' action='OrderPreInsert.asp' onSubmit='return order_onClick(this);'>" & vbNewLine
'wHTML = wHTML & "          <td>" & vbNewLine
'2012/07/10 GV Mod End
wHTML = wHTML & wIroKikakuCombo

if (IsNull(RS("�戵���~��")) = false) OR (IsNull(RS("������")) = false) OR (RS("B�i�t���O") = "Y" AND RS("B�i�����\����") <= 0) OR (IsNull(RS("�p�ԓ�")) = false AND RS("�����\����") <= 0 AND RS("��������") <= 0 AND wIroKikakuSelectedFl = true) then	'2011/06/09 hn mod
'2012/07/10 GV Mod Start
'  wHTML = wHTML & "<img src='images/Kanbai2.jpg'>" & vbNewLine
  wHTML = wHTML & "<p class='sold'><img src='images/icon_sold.png' alt='�������܂���'></p>" & vbNewLine
'2012/07/10 GV Mod End

else
	wHTML = wHTML & "            <div id='cart'>" & vbNewLine
	wHTML = wHTML & "                <span>��<input type='text' name='qt' value='1'></span><input type='image' src='images/btn_cart_productdetail.png' alt='�J�[�g�ɓ����' class='opover'>" & vbNewLine
	wHTML = wHTML & "                <input type='hidden' name='Item' value='" & RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "^" & Trim(RS("�F")) & "^" & Trim(RS("�K�i")) & "'>" & vbNewLine
	wHTML = wHTML & "            </div>" & vbNewLine
end if

'wHTML = wHTML & "          </td>" & vbNewLine	'2012/07/10 GV Del

'---- �ꏏ�ɍw������`�F�b�N���̒ǉ����i�o�^�p(���[�J�[�R�[�h^���i�R�[�h^�F^�K�i ,��؂�Ŋi�[�j
wHTML = wHTML & "          <input type='hidden' name='AdditionalItem' value=''>" & vbNewLine

wHTML = wHTML & "          </form>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine	'2012/07/10 GV Del

vUrl = ""
vUrl = vUrl & RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "^" & Trim(RS("�F")) & "^" & Trim(RS("�K�i"))

'----- �E�B�b�V�����X�g
if wProdTermFl = "Y" OR wIroKikakuSelectedFl = false then
else
	if wIroKikakuSelectedFl = true then
		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "        <tr>"
'		wHTML = wHTML & "          <td align='right'><a href='WishListAdd.asp?Item=" & Server.URLEncode(RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h") & "^" & RS("�F") & "^" & RS("�K�i")) & "' class='link'>�E�B�b�V�����X�g</a></td>" & vbNewLine
'		wHTML = wHTML & "        </tr>" & vbNewLine
	
		wHTML = wHTML & "          <p class='btn_wish'><a href='"
		if wUserID = "" Then
			wHTML = wHTML & g_HTTPS & "shop/LoginCheck.asp?RtnURL=" & g_HTTP & "shop/WishListAdd.asp?Item=" & Server.URLEncode(vUrl)
		Else
			wHTML = wHTML & "WishListAdd.asp?Item=" & Server.URLEncode(vUrl)
		End If
			wHTML = wHTML & "' ><img src='images/btn_wish.png' alt='�E�B�b�V�����X�g�ɒǉ�' class='opover' width='200' height='25'></a></p>" & vbNewLine
		'2012/07/10 GV Mod End
	end if
end if

'----- ���iID
'2012/07/10 GV Del Start
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td align='right'>���iID:" & RS("���iID") & "</td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'2012/07/10 GV Del End

'---- Twitter,Facebook
wHTML = wHTML & "          <ul class='sns'>" & vbNewLine
wHTML = wHTML & "            <li><a href='http://twitter.com/share' class='twitter-share-button' data-via='soundhouse_jp' data-lang='ja'>�c�C�[�g</a></li>" & vbNewLine
'wHTML = wHTML & "            <li><div class='fb-like' data-href='http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=" & Server.URLEncode(maker_cd & "^" & product_cd & "^" & iro & "^" & kikaku) & "' data-send='false' data-layout='button_count' data-width='145' data-show-faces='false'></div></li>" & vbNewLine
wHTML = wHTML & "            <li><iframe src='//www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.soundhouse.co.jp%2Fshop%2FProductDetail.asp%3FItem%3D" & Server.URLEncode(vUrl) & "&amp;send=false&amp;layout=button_count&amp;width=100&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font&amp;height=21&amp;appId=191447484218062' scrolling='no' frameborder='0' style='border:none; overflow:hidden; width:120px; height:21px;' allowTransparency='true'></iframe></li>" & vbNewLine
wHTML = wHTML & "          </ul>" & vbNewLine

'2012/07/10 GV Add Start
'---- �]��
wHTML = wHTML & wHyoukaHTML

wHTML = wHTML & "         <ul class='info'>" & vbNewLine
'---- ���̏��i�̖⍇��
wHTML = wHTML & "            <li><a href='" & g_HTTPS & "shop/Inquiry.asp?MakerNm=" & Server.URLEncode(RS("���[�J�[��")) & "&ProductCd=" & Server.URLEncode(RS("���i�R�[�h")) & "&CategoryNm=" & Server.URLEncode(RS("�J�e�S���[��")) & "'>���̏��i�ւ̂��₢���킹</a></li>" & vbNewLine

'---- �����ɂ���
wHTML = wHTML & "            <li><a href='../guide/kaimono.asp#souryou'>�����ɂ���</a></li>" & vbNewLine
'2012/07/10 GV Add End

'---- �F�B�Ɋ��߂�
if wUserID <> "" AND wProdTermFl <> "Y" then
	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td colspan='2' align='center' height='26'><a href='TellaFriend.asp?Item=" & Server.URLEncode(RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h")) & "' class='link'><img src='images/TomodachiNiSusumeru.gif' border='0'></a></td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
	wHTML = wHTML & "            <li><a href='TellaFriend.asp?Item=" & Server.URLEncode(RS("���[�J�[�R�[�h") & "^" & RS("���i�R�[�h")) & "'>�F�B�ɂ����߂�</a></li>" & vbNewLine
	'2012/07/10 GV Mod End
end if

'---- ���̏��i�̖⍇��
'2012/07/10 GV Del Start
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td colspan='2' align='center' height='26'><a href='" & g_HTTPS & "shop/Inquiry.asp?MakerNm=" & Server.URLEncode(RS("���[�J�[��")) & "&ProductCd=" & Server.URLEncode(RS("���i�R�[�h")) & "&CategoryNm=" & Server.URLEncode(RS("�J�e�S���[��")) & "' class='link'><img src='images/ShouhinNoToiawase.gif' border='0'></a></td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine
'2012/07/10 GV Del End

'---- Twitter�����N  2010/09/27 an add s
'2012/07/10 GV Del Start
'wHTML = wHTML & "        <tr>" & vbNewLine
'wHTML = wHTML & "          <td colspan='2' align='center' height='26'>" & vbNewLine
'wHTML = wHTML & "            <ul class='smbtn'>" & vbNewLine
'wHTML = wHTML & "              <li><a href='http://twitter.com/share' class='twitter-share-button' data-count='horizontal' data-via='soundhouse_jp' data-lang='ja'>Tweet</a></li>" & vbNewLine
'wHTML = wHTML & "              <li><a name='fb_share'>�V�F�A����</a></li>" & vbNewLine
'wHTML = wHTML & "            </ul>" & vbNewLine
'wHTML = wHTML & "          </td>" & vbNewLine
'wHTML = wHTML & "        </tr>" & vbNewLine    '2010/09/27 an add e

'wHTML = wHTML & "      </table>" & vbNewLine

'wHTML = wHTML & "    </td>" & vbNewLine
'wHTML = wHTML & "  </tr>" & vbNewLine
'wHTML = wHTML & "</table>" & vbNewLine
'2012/07/10 GV Del End

'2012/07/10 GV Add Start
wHTML = wHTML & "    </ul>" & vbNewLine
wHTML = wHTML & "  </div>" & vbNewLine
wHTML = wHTML & "</div></div>" & vbNewLine
'2012/07/10 GV Add End

wProductHTML = wHTML

'2013/05/17 GV #1505 add start
wEAInventoryData = vInventoryCd
wEAPrice         = wPrice
'2013/05/17 GV #1505 add end

End Function

'========================================================================
'
'	Function	���R�����h����	2009/12/17
'
'========================================================================
'
Function CreateRecommendHTML()

'2013/05/17 GV #1505 add start
wRecommendJS = fEAgency_CreateRecommendJS(wEAProductDetailData, wEAIroKikakuData)
'2013/05/17 GV #1505 add end

'2013/08/07 if-web del s
'Dim RSv
'Dim vPrice
'
'
'---- ���R�����h����
'wSQL = ""
'wSQL = wSQL & "SELECT DISTINCT TOP 5"
'wSQL = wSQL & "       a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B�i�t���O = 'Y' THEN a.B�i�P��"    '2010/11/12 an add
'wSQL = wSQL & "         WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS ���̔��P��"    '2010/11/12 an mod
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , b.���[�J�[��"
'wSQL = wSQL & "     , e.�ގ��x"
'wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d WITH (NOLOCK)"
'
'if wUserID = "" then
'	wSQL = wSQL & "     , ���R�����h���ʃA�N�Z�X e WITH (NOLOCK)"
'else
'	wSQL = wSQL & "     , ���R�����h���ʍw�� e WITH (NOLOCK)"
'end if
'
'wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = e.���R�����h���[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.���i�R�[�h = e.���R�����h���i�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND a.�戵���~�� IS NULL"
'wSQL = wSQL & "   AND ((a.�p�ԓ� IS NULL) OR (a.�p�ԓ� IS NOT NULL AND d.�����\���� > 0 AND d.�������� > 0))"	'2011/06/09 hn mod
'wSQL = wSQL & "   AND e.���[�J�[�R�[�h = '" & maker_cd & "'"
'wSQL = wSQL & "   AND e.���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       e.�ގ��x DESC"
'
'@@@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'wHTML = ""
'
'if RSv.EOF = true then
'	RSv.close
'	exit function
'end if
'
'----- ���R�����h���iHTML�ҏW
'2012/07/10 GV Del Start
'wHTML = wHTML & "<table id=Shop_right_relation cellSpacing=0 cellPadding=0 width=188 border=0>" & vbNewLine
'wHTML = wHTML & "  <tr>" & vbNewLine
'2012/07/10 GV Del End
'
'if wUserID = "" then
'	wHTML = wHTML & "    <td style='padding:5px; border:#999999 solid 1px;' bgcolor='#FFCC66'>���̃A�C�e���������l��<br>����ȃA�C�e�������Ă��܂��B</td>" & vbNewLine	'2012/07/10 GV Del
'else
'	wHTML = wHTML & "    <td style='padding:5px; border:#999999 solid 1px;' bgcolor='#CCFF00'>���̃A�C�e���𔃂����l��<br>����ȃA�C�e���������Ă��܂��B</td>" & vbNewLine
'end if
'2012/07/10 GV Add Start
'wHTML = wHTML & "<div class='detail_side_inner01'><div class='detail_side_inner02'>" & vbNewLine
'wHTML = wHTML & "  <div class='detail_side_inner_box'>" & vbNewLine
'wHTML = wHTML & "    <!--���̃A�C�e���������l�� -->" & vbNewLine
'wHTML = wHTML & "    <h4 class='detail_sub'>���̃A�C�e���������l��<br>����ȃA�C�e�������Ă��܂��B</h4>" & vbNewLine
'wHTML = wHTML & "    <ul class='check_item'>" & vbNewLine
'2012/07/10 GV Add End
'
'wHTML = wHTML & "  </tr>" & vbNewLine	'2012/07/10 GV Del
'
'Do Until RSv.EOF = true
'	vPrice = calcPrice(RSv("���̔��P��"), wSalesTaxRate) '2010/11/12 an mod �̔��P�������̔��P��
'
'	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "  <tr>" & vbNewLine
'	wHTML = wHTML & "    <td class=base align=middle>" & vbNewLine
'	wHTML = wHTML & "      <table id=Shop_right_product cellSpacing=0 cellPadding=0 width=180 border=0>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h")) & "'><img src='prod_img/" & RSv("���i�摜�t�@�C����_��") & "' width='170' height='85' border='0'></a></td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "          <td>" & RSv("���[�J�[��") & " " & RSv("���i��") & "</td>" & vbNewLine
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "        <tr>" & vbNewLine
'	wHTML = wHTML & "      <li>" & vbNewLine
'
'	wHTML = wHTML & "        <p><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h")) & "'>"
'	If RSv("���i�摜�t�@�C����_��") <> "" Then
'		wHTML = wHTML & "<img src='prod_img/" & RSv("���i�摜�t�@�C����_��") & "' alt='" & Replace(RSv("���[�J�[��") & " / " & RSv("���i��"),"'","&#39;") & "' class='opover'>"
'	End If
'	wHTML = wHTML & RSv("���[�J�[��") & " / " & RSv("���i��") & "</a></p>" & vbNewLine
'	'2012/07/10 GV Mod End
'
'	if RSv("ASK���i�t���O") <> "Y" then  '2010/04/06 an changed start
'		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "          <td>" & FormatNumber(vPrice, 0) & "�~(�ō�)</td>" & vbNewLine
'		wHTML = wHTML & "        <p>" & FormatNumber(vPrice, 0) & "�~(�ō�)</p>" & vbNewLine
'		'2012/07/10 GV Mod End
'	else
'2011/10/19 hn mod s
'		wHTML = wHTML & "          <td><a href='JavaScript:void(0);' onClick=""askWin=window.open('AskPrice.asp?MakerName=" & Server.URLEncode(RSv("���[�J�[��")) & "&ProductName=" & Server.URLEncode(RSv("���i��")) & "&Price=" & vPrice & "' ,'ask', 'width=250 height=80 scrollbars=false memubar=false toolbar=false statusbar=false personalbar=false locationbar=false');"" class='link'><strong>ASK</b></strong></td>" & vbNewLine
'
'		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "          <td><a class='tip'>ASK<span>" & FormatNumber(vPrice, 0) & "�~(�ō�)</span></a></td>" & vbNewLine
'		wHTML = wHTML & "        <p><a class='tip'>ASK<span>" & FormatNumber(vPrice, 0) & "�~(�ō�)</span></a></p>" & vbNewLine
'		'2012/07/10 GV Mod End
'2011/10/19 hn mod e
'
'	end if       '2010/04/06 an changed end,  2010/04/21 an changed
'
'	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "        </tr>" & vbNewLine
'	wHTML = wHTML & "      </table>" & vbNewLine
'	wHTML = wHTML & "    </td>" & vbNewLine
'	wHTML = wHTML & "  </tr>" & vbNewLine
'	wHTML = wHTML & "      </li>" & vbNewLine
'	'2012/07/10 GV Mod End
'
'	RSv.MoveNext
'Loop
'
'2012/07/10 GV Add Start
'wHTML = wHTML & "    </ul>" & vbNewLine
'wHTML = wHTML & "  </div>" & vbNewLine
'wHTML = wHTML & "</div></div>" & vbNewLine
'2012/07/10 GV Add End
'
'wHTML = wHTML & "</table>" & vbNewLine	'2012/07/10 GV Del
'
'wRecommendHTML = wHTML
'
'RSv.close
'2013/08/07 if-web del e

End function

'========================================================================
'
'	Function	���R�����h���i�擾  '2012/04/10
'
'========================================================================
Function CreateRecommendBuyHTML()

'2013/05/17 GV #1505
wRecommendBuyJS = fEAgency_CreateRecommendBuyJS(wEAProductDetailData, wEAInventoryData, wEAPrice, wEAPriceExcTax, wEAIroKikakuData)

'2013/08/07 if-web del s
'Dim RSv
'Dim iCnt	'2012/07/10 GV Add
'
'
'---- ���R�����h���i�擾(�ގ��x���傫��5���i)
'wSQL = ""
'
'1�s�ɕ\�����錏����5������4���ɕύX 2012/07/20 ok Mod
'wSQL = wSQL & "SELECT DISTINCT TOP 5"
'wSQL = wSQL & "SELECT DISTINCT TOP 4"
'wSQL = wSQL & "       a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS �̔��P��"
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "     , b.���[�J�[��"
'wSQL = wSQL & "     , e.�ގ��x"
'wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d WITH (NOLOCK)"
'wSQL = wSQL & "     , ���R�����h���ʍw�� e WITH (NOLOCK)"
'wSQL = wSQL & " WHERE a.���[�J�[�R�[�h = e.���R�����h���[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.���i�R�[�h = e.���R�����h���i�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND a.�戵���~�� IS NULL"
'wSQL = wSQL & "   AND ((a.�p�ԓ� IS NULL) OR (a.�p�ԓ� IS NOT NULL AND d.�����\���� > 0))"
'wSQL = wSQL & "   AND e.���[�J�[�R�[�h = '" & maker_cd & "'"
'wSQL = wSQL & "   AND e.���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       e.�ގ��x DESC"
'wSQL = wSQL & "     , a.�J�e�S���[�R�[�h"
'
'@@@@response.write(wSQL)
'
'Set RSv = Server.CreateObject("ADODB.Recordset")
'RSv.Open wSQL, Connection, adOpenStatic
'
'wHTML = ""
'iCnt = 0	'2012/07/10 GV Add
'
'if RSv.EOF = false then
'
'	wHTML = ""
'	'2012/07/10 GV Mod Start
'	wHTML = wHTML & "<h2 id=""recommend_h"">���̃A�C�e���𔃂����l�͂���ȃA�C�e���������Ă��܂��B</h2>" & vbNewLine
'	wHTML = wHTML & "<ul id=""recommend_box"">" & vbNewLine
'	wHTML = wHTML & "<h2 class='detail_title'>���̃A�C�e���𔃂����l�͂���ȃA�C�e���������Ă��܂�</h2>" & vbNewLine
'	'2012/07/10 GV Mod End
'
'	Do Until RSv.EOF = True
'
'		'2012/07/10 GV Add Start
'		if iCnt mod 4 = 0 then
'			wHTML = wHTML & "<ul class='relation other'>" & vbNewLine
'		end if
'		'2012/07/10 GV Add End
'
'		wPrice = calcPrice(RSv("�̔��P��"), wSalesTaxRate)
'
'		wHTML = wHTML & "  <li>" & vbNewLine	'2012/07/10 GV Add
'		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "    <a href=""ProductDetail.asp?Item=" & RSv("���[�J�[�R�[�h") & "%5E" & RSv("���i�R�[�h") & """>"
'		wHTML = wHTML & "    <p><a href='ProductDetail.asp?Item=" & RSv("���[�J�[�R�[�h") & "%5E" & RSv("���i�R�[�h") & "'>"
'		'2012/07/10 GV Mod End
'		if RSv("���i�摜�t�@�C����_��") <> "" then
'			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "<img src=""prod_img/" & RSv("���i�摜�t�@�C����_��") & """ alt=""" & RSv("���[�J�[��") & " " & RSv("���i��") & """>"
'			wHTML = wHTML & "<img src='prod_img/" & RSv("���i�摜�t�@�C����_��") & "' alt='" & Replace(RSv("���[�J�[��") & " " & RSv("���i��"),"'","&#39;") & "' class='opover'></a></p>"
'			'2012/07/10 GV Mod End
'		else
'			wHTML = wHTML & "<img src=""prod_img/n/nopict-.jpg"" alt="""">"
'		end if
'		wHTML = wHTML & "</a><br>" & vbNewLine	'2012/07/10 GV Del
'		'2012/07/10 GV Mod Start
'		wHTML = wHTML & "    " & RSv("���[�J�[��") & " " & RSv("���i��") & "<br>" & vbNewLine
'		wHTML = wHTML & "    <p><a href='ProductDetail.asp?Item=" & RSv("���[�J�[�R�[�h") & "%5E" & RSv("���i�R�[�h") & "'>"
'		wHTML = wHTML & "    " & RSv("���[�J�[��") & " / " & RSv("���i��") & "</a></p>" & vbNewLine
'		'2012/07/10 GV Mod End
'
'		wHTML = wHTML & "    <div class='box'>" & vbNewLine	'2012/07/10 GV Add
'		If RSv("ASK���i�t���O") <> "Y" Then
'			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "    " & FormatNumber(wPrice,0) & "�~(�ō�)" & vbNewLine
'			wHTML = wHTML & "      <p>" & FormatNumber(wPrice,0) & "�~(�ō�)</p>" & vbNewLine
'			'2012/07/10 GV Mod End
'		Else
'			'2012/07/10 GV Mod Start
'			wHTML = wHTML & "    <a class='tip'>ASK<span>"& FormatNumber(wPrice,0) & "�~(�ō�)</span></a>" & vbNewLine
'			wHTML = wHTML & "      <p><a class='tip'>ASK<span>"& FormatNumber(wPrice,0) & "�~(�ō�)</span></a></p>" & vbNewLine
'			'2012/07/10 GV Mod End
'
'		End If
'		wHTML = wHTML & "    </div>" & vbNewLine	'2012/07/10 GV Add
'
'		wHTML = wHTML & "  </li>" & vbNewLine		'2012/07/10 GV Add
'
'		RSv.MoveNext
'
'		'2012/07/10 GV Add Start
'		if (iCnt mod 4 = 3) Or (RSv.RecordCount = iCnt+1) then
'			wHTML = wHTML & "</ul>" & vbNewLine
'		end if
'		iCnt = iCnt + 1
'		'2012/07/10 GV Add End
'	Loop
'
'End if
'
'RSv.Close
'
'wRecommendBuyHTML = wHTML
'2013/08/07 if-web del e

End function

'========================================================================
'
'	Function	�֘A�V���[�Y���i
'
'========================================================================
'
Function CreateSeriesHTML()

Dim RSv
Dim vPrice
Dim vRecordCount
Dim vCount

'---- �֘A�V���[�Y���i
wSQL = ""
' 2012/01/18 GV Mod Start
'wSQL = wSQL & "SELECT DISTINCT "
'wSQL = wSQL & "       a.���[�J�[�R�[�h"
'wSQL = wSQL & "     , a.���i�R�[�h"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , a.���i�摜�t�@�C����_��"
'wSQL = wSQL & "     , CASE"
'wSQL = wSQL & "         WHEN a.B�i�t���O = 'Y' THEN a.B�i�P��"   '2010/11/12 an add
'wSQL = wSQL & "         WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P��"
'wSQL = wSQL & "         ELSE a.�̔��P��"
'wSQL = wSQL & "       END AS ���̔��P��"   '2010/11/12 an mod
'wSQL = wSQL & "     , a.ASK���i�t���O"
'wSQL = wSQL & "     , b.���[�J�[��"
'wSQL = wSQL & "     , c.�J�e�S���[��"
'wSQL = wSQL & "     , c.�\����"
'wSQL = wSQL & "     , d.�F"
'wSQL = wSQL & "     , d.�K�i"
'wSQL = wSQL & "  FROM Web���i a WITH (NOLOCK)"
'wSQL = wSQL & "     , ���[�J�[ b WITH (NOLOCK)"
'wSQL = wSQL & "     , �J�e�S���[ c WITH (NOLOCK)"
'wSQL = wSQL & "     , Web�F�K�i�ʍ݌� d WITH (NOLOCK)"
'wSQL = wSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND c.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h"
'wSQL = wSQL & "   AND d.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
'wSQL = wSQL & "   AND d.���i�R�[�h = a.���i�R�[�h"
'wSQL = wSQL & "   AND a.Web���i�t���O = 'Y'"
'wSQL = wSQL & "   AND a.�戵���~�� IS NULL"
'wSQL = wSQL & "   AND ((a.�p�ԓ� IS NULL) OR (a.�p�ԓ� IS NOT NULL AND d.�����\���� > 0 AND d.�������� > 0))"		'2011/06/09 hn mod
'wSQL = wSQL & "   AND NOT (a.���[�J�[�R�[�h = '" & RS("���[�J�[�R�[�h") & "'"
'wSQL = wSQL & "       AND a.���i�R�[�h = '" & RS("���i�R�[�h") & "')"
'wSQL = wSQL & "   AND a.�V���[�Y�R�[�h = '" & RS("�V���[�Y�R�[�h") & "'"
'wSQL = wSQL & " ORDER BY"
'wSQL = wSQL & "       c.�\����"
'wSQL = wSQL & "     , b.���[�J�[��"
'wSQL = wSQL & "     , a.���i��"
'wSQL = wSQL & "     , d.�F"
'wSQL = wSQL & "     , d.�K�i"
wSQL = wSQL & "SELECT DISTINCT "
wSQL = wSQL & "      a.���[�J�[�R�[�h "
wSQL = wSQL & "    , a.���i�R�[�h "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , a.���i�摜�t�@�C����_�� "
wSQL = wSQL & "    , CASE "
wSQL = wSQL & "        WHEN a.B�i�t���O = 'Y'                     THEN a.B�i�P�� "
wSQL = wSQL & "        WHEN a.�����萔�� > a.������󒍍ϐ��� THEN a.������P�� "
wSQL = wSQL & "        ELSE                                            a.�̔��P�� "
wSQL = wSQL & "      END AS ���̔��P�� "
wSQL = wSQL & "    , a.ASK���i�t���O "
wSQL = wSQL & "    , b.���[�J�[�� "
wSQL = wSQL & "    , c.�J�e�S���[�� "
wSQL = wSQL & "    , c.�\���� "
wSQL = wSQL & "    , d.�F "
wSQL = wSQL & "    , d.�K�i "
wSQL = wSQL & "FROM "
wSQL = wSQL & "    Web���i                      a WITH (NOLOCK) "
wSQL = wSQL & "      INNER JOIN ���[�J�[        b WITH (NOLOCK) "
wSQL = wSQL & "        ON     b.���[�J�[�R�[�h   = a.���[�J�[�R�[�h "
wSQL = wSQL & "      INNER JOIN �J�e�S���[      c WITH (NOLOCK) "
wSQL = wSQL & "        ON     c.�J�e�S���[�R�[�h = a.�J�e�S���[�R�[�h "
wSQL = wSQL & "      INNER JOIN Web�F�K�i�ʍ݌� d WITH (NOLOCK) "
wSQL = wSQL & "        ON     d.���[�J�[�R�[�h   = a.���[�J�[�R�[�h "
wSQL = wSQL & "           AND d.���i�R�[�h       = a.���i�R�[�h "
wSQL = wSQL & "      LEFT JOIN ( SELECT 'Y' AS 'ShohinWebY' )   t1 "
wSQL = wSQL & "        ON     a.Web���i�t���O    = t1.ShohinWebY  "
wSQL = wSQL & "WHERE "
wSQL = wSQL & "        t1.ShohinWebY IS NOT NULL "
wSQL = wSQL & "    AND a.�戵���~�� IS NULL "
wSQL = wSQL & "    AND (    (a.�p�ԓ� IS NULL) "
wSQL = wSQL & "         OR  (    a.�p�ԓ� IS NOT NULL "
wSQL = wSQL & "              AND d.�����\���� > 0 "
wSQL = wSQL & "              AND d.�������� > 0)) "
wSQL = wSQL & "    AND NOT  (    a.���[�J�[�R�[�h = '" & RS("���[�J�[�R�[�h") & "' "
wSQL = wSQL & "              AND a.���i�R�[�h = '" & Replace(RS("���i�R�[�h"), "'", "''") & "') "	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)
wSQL = wSQL & "    AND a.�V���[�Y�R�[�h = '" & RS("�V���[�Y�R�[�h") & "' "
wSQL = wSQL & "ORDER BY "
wSQL = wSQL & "      c.�\���� "
wSQL = wSQL & "    , b.���[�J�[�� "
wSQL = wSQL & "    , a.���i�� "
wSQL = wSQL & "    , d.�F "
wSQL = wSQL & "    , d.�K�i "
' 2012/01/18 GV Mod End

'@@@@@@response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic
vRecordCount = RSv.RecordCount

wHTML = ""
vCount = 0

if RSv.EOF = true then
	RSv.close
	exit function
end if

'2012/07/10 GV Add Start
wHTML = wHTML & "<div id='detail_side'>" & vbNewLine
wHTML = wHTML & "  <div id='detail_side_inner01'><div id='detail_side_inner02'>" & vbNewLine
wHTML = wHTML & "    <div class='detail_side_inner_box'>" & vbNewLine
wHTML = wHTML & "      <!-- �֘A�V���[�Y -->" & vbNewLine
wHTML = wHTML & "      <h4 class='detail_sub'><a href='SearchList.asp'>CLASSIC PRO CPAHW�V���[�Y</a></h4>" & vbNewLine
wHTML = wHTML & "      <ul class='check_item'>" & vbNewLine
wHTML = wHTML & "        <li>" & vbNewLine
wHTML = wHTML & "          <p><a href='SearchList.asp'><img src='prod_img/f/fender_jbesquireb-.jpg' alt='PLAYTECH / PAUL REED SMITH PRS Guitar Strings' class='opover'></a>�ړ��ɕ֗��ȃL�������O�n���h���E�L���X�^�[�t���̎��������b�N�P�[�X�B�d���@�ނ̈ړ��A�^���ɍœK�ł��B</p>" & vbNewLine
wHTML = wHTML & "        </li>" & vbNewLine
wHTML = wHTML & "      </ul>" & vbNewLine
wHTML = wHTML & "    </div>" & vbNewLine
wHTML = wHTML & "  </div></div>" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine
'2012/07/10 GV Add End
'----- �֘A�V���[�Y���iHTML�ҏW
wHTML = wHTML & "<table width='188' border='0' cellspacing='0' cellpadding='0' id='Shop_right_relation'>" & vbNewLine
wHTML = wHTML & "  <tr>" & vbNewLine
wHTML = wHTML & "    <td align='left' class='head'>�֘A�V���[�Y���i</td>" & vbNewLine
wHTML = wHTML & "  </tr>" & vbNewLine

Do Until (RSv.EOF = true OR vCount >= 5)
	vPrice = calcPrice(RSv("���̔��P��"), wSalesTaxRate)   '2010/11/12 an mod �̔��P�������̔��P��

  wHTML = wHTML & "  <tr>" & vbNewLine
  wHTML = wHTML & "    <td align='center' class='base'>" & vbNewLine
  wHTML = wHTML & "      <table width='180' border='0' cellpadding='0' cellspacing='0' id='Shop_right_product'>" & vbNewLine
  wHTML = wHTML & "        <tr>" & vbNewLine

  wHTML = wHTML & "          <td><a href='ProductDetail.asp?Item=" & Server.URLEncode(RSv("���[�J�[�R�[�h") & "^" & RSv("���i�R�[�h") & "^" & RSv("�F") & "^" & RSv("�K�i")) & "'>"
  If RSv("���i�摜�t�@�C����_��") <> "" Then
    wHTML = wHTML & "<img src='prod_img/" & RSv("���i�摜�t�@�C����_��") & "' width='170' height='85' border='0'>"
  End If
  wHTML = wHTML & "</a></td>" & vbNewLine

  wHTML = wHTML & "        </tr>" & vbNewLine
  wHTML = wHTML & "        <tr>" & vbNewLine
  wHTML = wHTML & "          <td>" & RSv("���[�J�[��") & " " & RSv("���i��") & "</td>" & vbNewLine
  wHTML = wHTML & "        </tr>" & vbNewLine
  wHTML = wHTML & "        <tr>" & vbNewLine

'2011/10/19 hn add s
	if RSv("ASK���i�t���O") <> "Y" then
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "          <td>" & FormatNumber(vPrice, 0) & "�~(�ō�)</td>" & vbNewLine
		wHTML = wHTML & "          <td>" & FormatNumber(RSv("���̔��P��"), 0) & "�~(�Ŕ�)</td>" & vbNewLine
		wHTML = wHTML & "          <td><strong>(�ō�)" & FormatNumber(vPrice, 0) & "�~</strong></td>" & vbNewLine
'2014/03/19 GV mod end <-----
	else
'2014/03/19 GV mod start ---->
'		wHTML = wHTML & "          <td><a class='tip'>ASK<span>" & FormatNumber(vPrice, 0) & "�~(�ō�)</span></a></td>" & vbNewLine
		wHTML = wHTML & "          <td><a class='tip'>ASK<span class='exc-tax'>" & FormatNumber(RSv("���̔��P��"), 0) & "�~(�Ŕ�)</span>"
		wHTML = wHTML & "<span class='inc-tax'>(�ō�)" & FormatNumber(vPrice, 0) & "�~</span></a></td>" & vbNewLine
'2014/03/19 GV mod end <-----
	end if
'2011/10/19 hn add e

  wHTML = wHTML & "        </tr>" & vbNewLine
  wHTML = wHTML & "      </table>" & vbNewLine
  wHTML = wHTML & "    </td>" & vbNewLine
  wHTML = wHTML & "  </tr>" & vbNewLine

	RSv.MoveNext
	vCount = vCount + 1
Loop

if vRecordCount > vCount then
	wHTML = wHTML & "  <tr>" & vbNewLine
	wHTML = wHTML & "    <td><a href='SearchList.asp?i_type=se&sSeriesCd=" & RS("�V���[�Y�R�[�h") & "' class='link'>���̑��֘A�V���[�Y���i>></a></td>" & vbNewLine
	wHTML = wHTML & "  </tr>" & vbNewLine
end if

wHTML = wHTML & "</table>" & vbNewLine

wSeriesHTML = wHTML

RSv.close

End function

'========================================================================
'
'	Function	�ŋ߃`�F�b�N�������i�ɒǉ�
'
'========================================================================
'
Function AddViewdProduct()

Dim RSv

'----
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM �ŋ߃`�F�b�N�������i"
wSQL = wSQL & " WHERE �ڋq�ԍ� = " & wUserID
wSQL = wSQL & "   AND ���[�J�[�R�[�h = '" & maker_cd & "'"
wSQL = wSQL & "   AND ���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)
wSQL = wSQL & "   AND �F = '" & iro & "'"
wSQL = wSQL & "   AND �K�i = '" & kikaku & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	RSv.AddNew

	RSv("�ڋq�ԍ�") = wUserID
	RSv("���[�J�[�R�[�h") = maker_cd
	RSv("���i�R�[�h") = product_cd
	RSv("�F") = iro
	RSv("�K�i") = kikaku
end if

RSv("�`�F�b�N��") = Now()

RSv.Update
RSv.close

End function

'========================================================================
'
'	Function	���i�A�N�Z�X�J�E���g�o�^�i�y�[�W�r���[�j
'
'========================================================================
'
Function SetAccessCount()

Dim vYYYYMM
Dim RSv

'---- ����Z�b�V������1��ڂ��ǂ����`�F�b�N
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM �Z�b�V�����f�[�^"
wSQL = wSQL & " WHERE SessionID = '" & gSessionID & "'"			'2011/04/14 hn mod
wSQL = wSQL & "   AND ���ږ� = '" & maker_cd & "^" & product_cd & "'"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

if RSv.EOF = true then
	'---- �Z�b�V�����f�[�^�o�^
	RSv.AddNew

	RSv("SessionID") = gSessionID		'2011/04/14 hn mod
	RSv("���ږ�") = maker_cd & "^" & product_cd
	RSv("���e") = "�y�[�W�r���[�`�F�b�N�p"
	RSv("�ŏI�X�V��") = Now()

	RSv.Update
	RSv.close

	'---- �y�[�W�r���[�o�^
	vYYYYMM = Year(Now()) & Right("0" & Month(Now()),2)

	wSQL = ""
	wSQL = wSQL & "SELECT *"
	wSQL = wSQL & "  FROM ���i�A�N�Z�X����"
	wSQL = wSQL & " WHERE ���[�J�[�R�[�h = '" & maker_cd & "'"
	wSQL = wSQL & "   AND ���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"	' 2012/01/23 GV Mod (�R�[�h���ɃV���O���N�I�[�e�[�V���������݂����ꍇ�̑Ή�)
	wSQL = wSQL & "   AND �N�� = '" & vYYYYMM & "'"

	Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

	if RSv.EOF = true then
		RSv.AddNew

		RSv("���[�J�[�R�[�h") = maker_cd
		RSv("���i�R�[�h") = product_cd
		RSv("�N��") = vYYYYMM
		RSv("�y�[�W�r���[����") = 1
	else
		RSv("�y�[�W�r���[����") = RSv("�y�[�W�r���[����") + 1
	end if

	RSv.Update
	RSv.close
end if

End function

'========================================================================
'
'	Function	���R�����h���i�A�N�Z�X���O	2009/12/17
'
'========================================================================
'
Function AddRecommendAccessLog()

Dim RSv

'----
wSQL = ""
wSQL = wSQL & "SELECT *"
wSQL = wSQL & "  FROM ���R�����h���i�A�N�Z�X���O"
wSQL = wSQL & " WHERE 1 = 2"

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic

'---- ���R�����h���i�A�N�Z�X���O�o�^
RSv.AddNew

RSv("���R�����h���[�U�[ID") = gSessionID		'2011/04/14 hn mod
RSv("���[�J�[�R�[�h") = maker_cd
RSv("���i�R�[�h") = product_cd
RSv("���[�U�[�G�[�W�F���g") = Request.ServerVariables("HTTP_USER_AGENT")
RSv("�A�N�Z�X��") = Now()

RSv.Update
RSv.close

End function

'========================================================================
'
'	Function	��֏��i�擾���\�b�h GV 2012/05/01
'
'========================================================================
Function GetSubstituteItem()
    
    Dim RSv
    Dim RSvSub
    Dim vSql

    wSQL = ""
	wSQL = wSQL & "SELECT "
    wSQL = wSQL & "    a.���[�J�[�R�[�h, "
    wSQL = wSQL & "    a.���i�R�[�h, "
    wSQL = wSQL & "    a.���i��, "
    wSQL = wSQL & "    a.��p�@�탁�[�J�[�R�[�h, "
    wSQL = wSQL & "    a.��p�@�폤�i�R�[�h, "
    wSQL = wSQL & "    b.�����\���� "
	wSQL = wSQL & "FROM "
    wSQL = wSQL & "    Web���i a WITH (NOLOCK) "
    wSQL = wSQL & "INNER JOIN  "
    wSQL = wSQL & "    Web�F�K�i�ʍ݌� b WITH (NOLOCK) ON  "
    wSQL = wSQL & "    a.���[�J�[�R�[�h = b.���[�J�[�R�[�h AND "
    wSQL = wSQL & "    a.���i�R�[�h = b.���i�R�[�h "
	wSQL = wSQL & "WHERE a.���[�J�[�R�[�h = '" & maker_cd & "'"
	wSQL = wSQL & "    AND a.���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"
	
    Set RSv = Server.CreateObject("ADODB.Recordset")
	RSv.Open wSQL, Connection, adOpenStatic, adLockOptimistic
    
	'2013/08/14 GV add start
	If RSv.EOF = True Then
		'//�f�[�^�����݂��Ȃ��ꍇ�I��
		RSv.Close
		Exit Function
	End If
	'2013/08/14 GV add end

    '//�݌ɂ����݂��Ȃ��ꍇ�ő�֏��i�����݂���ꍇ
	If RSv("�����\����") <= 0 And (RSv("��p�@�탁�[�J�[�R�[�h") <> "" And RSv("��p�@�폤�i�R�[�h") <> "") Then
        
        '//��p�@���i�R�[�h�̏��i���A�݌ɂ��擾����
        'calcPrice(RS("�̔��P��"), wSalesTaxRate)
        vSql = ""
        vSql = vSql & " SELECT "
	    vSql = vSql & " a.���[�J�[�R�[�h,"
	    vSql = vSql & " c.���[�J�[��,"
	    vSql = vSql & " a.���i�R�[�h,"
	    vSql = vSql & " a.���i��,"
		vSql = vSql & " CASE"
		vSql = vSql & "   WHEN (a.�����萔�� > a.������󒍍ϐ��� AND a.�����萔�� > 0) THEN a.������P��"
		vSql = vSql & "   WHEN (a.B�i�t���O = 'Y') THEN a.B�i�P��"
		vSql = vSql & "   ELSE a.�̔��P��"
		vSql = vSql & " END AS �̔��P��,"
	    vSql = vSql & " a.���i�摜�t�@�C����_��,"
	    vSql = vSql & " a.���i�摜�t�@�C����_��,"
		vSql = vSql & " a.ASK���i�t���O"
        vSql = vSql & " FROM "
        vSql = vSql & " 	Web���i a WITH (NOLOCK) "
        vSql = vSql & " INNER JOIN "
        vSql = vSql & " 	Web�F�K�i�ʍ݌� b WITH (NOLOCK) ON "
        vSql = vSql & " 		a.���[�J�[�R�[�h = b.���[�J�[�R�[�h AND "
        vSql = vSql & " 		a.���i�R�[�h = b.���i�R�[�h "
        vSql = vSql & " INNER JOIN ���[�J�[ c WITH (NOLOCK) ON "
        vSql = vSql & " 	c.���[�J�[�R�[�h = a.���[�J�[�R�[�h "
        vSql = vSql & " WHERE  "
        vSql = vSql & " 	a.���[�J�[�R�[�h = '" & RSv("��p�@�탁�[�J�[�R�[�h") &  "' AND "
        vSql = vSql & " 	a.���i�R�[�h = '" & RSv("��p�@�폤�i�R�[�h") & "' AND "
        vSql = vSql & " 	a.Web���i�t���O = 'Y' AND "
        vSql = vSql & " 	((b.�����\���� > 0 ) OR (a.B�i�t���O = 'Y') AND (b.B�i�����\���� > 0))"
        '@@@@@@Response.Write(vSql)
        
        Set RSvSub = Server.CreateObject("ADODB.Recordset")
	    RSvSub.Open vSql, Connection, adOpenStatic, adLockOptimistic
        
        If RSvSub.EOF = True Then
            '//�f�[�^�����݂��Ȃ��ꍇ�I��
            RSv.close
	        RSvSub.close
	        Exit Function
        End If

        wSubItemHTML = ""
	'2012/07/10 GV Mod Start
'        wSubItemHTML = wSubItemHTML & "<div id='alt-item'>" & vbNewLine
'        wSubItemHTML = wSubItemHTML & "<p class='head'>���̃A�C�e����<br>�����ɂ��͂��ł��܂��B</p>" & vbNewLine
'        wSubItemHTML = wSubItemHTML & "<p><a href='ProductDetail.asp?Item=" & RSvSub("���[�J�[�R�[�h") & "%5E" & RSvSub("���i�R�[�h") & "'>" & vbNewLine
'        wSubItemHTML = wSubItemHTML & "<img src='prod_img/" & RSvSub("���i�摜�t�@�C����_��") & "' width='170' height='85' border='0'></a><br>" & vbNewLine
'        wSubItemHTML = wSubItemHTML & RSvSub("���[�J�[��") & " " & RSvSub("���i��") & "<br>" & vbNewLine
'		If RSvSub("ASK���i�t���O") <> "Y" Then
'	        wSubItemHTML = wSubItemHTML & FormatNumber(calcPrice(RSvSub("�̔��P��"), wSalesTaxRate),0) & "�~(�ō�)</p>" & vbNewLine
'		Else
'	        wSubItemHTML = wSubItemHTML & "<a class='tip'>ASK<span>" & FormatNumber(calcPrice(RSvSub("�̔��P��"), wSalesTaxRate),0) & "�~(�ō�)</span></a></p>" & vbNewLine
'		End If
'        wSubItemHTML = wSubItemHTML & "</div>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "<div class='detail_side_inner01'><div class='detail_side_inner02'>"
	wSubItemHTML = wSubItemHTML & "<div class='detail_side_inner_box'>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "  <!-- ���̃A�C�e���͂����ɂ��͂��ł��܂� -->" & vbNewLine
	wSubItemHTML = wSubItemHTML & "  <h4 class='detail_sub truck'>���̃A�C�e����<br>�����ɂ��͂��ł��܂�</h4>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "  <ul class='check_item'>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "    <li>" & vbNewLine

	wSubItemHTML = wSubItemHTML & "      <p><a href='ProductDetail.asp?Item=" & RSvSub("���[�J�[�R�[�h") & "%5E" & RSvSub("���i�R�[�h") & "'>"
	If RSvSub("���i�摜�t�@�C����_��") <> "" Then
		wSubItemHTML = wSubItemHTML & "<img src='prod_img/" & RSvSub("���i�摜�t�@�C����_��") & "' alt='" & Replace(RSvSub("���[�J�[��") & " / " & RSvSub("���i��"),"'","&#39;") & "' class='opover'>"
	End If		
	wSubItemHTML = wSubItemHTML & RSvSub("���[�J�[��") & " / " & RSvSub("���i��") & "</a></p>" & vbNewLine

	wSubItemHTML = wSubItemHTML & "    </li>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "  </ul>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "</div>" & vbNewLine
	wSubItemHTML = wSubItemHTML & "</div></div>" & vbNewLine
	'2012/07/10 GV Mod End

    Else
        RSv.close
        Exit Function
    End If

End Function


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

'========================================================================
'
'	Function	�����p�X���畨���p�X���擾����i�t�@�C���̑��݊m�F�L��j
'               2011/11/22 an FH����R�s�[
'
'========================================================================
Function GetMapPath(pTargetPath, pFileExtention)

Dim vFSO
Dim vTarget

pFileExtention = ""
GetMapPath = ""

Set vFSO = CreateObject("Scripting.FileSystemObject")
vTarget = Server.MapPath(pTargetPath)

pFileExtention = vFSO.GetExtensionName(vTarget)

' �g���q��txt�̏ꍇ�A"/shop"����̑��΃A�h���X���w�肳��Ă���
If LCase(pFileExtention) = "txt" Then
	vTarget = Server.MapPath(pTargetPath)
End If

If vFSO.FileExists(vTarget) = True Then
	GetMapPath = vTarget
End If

Set vFSO = Nothing

End Function

'2012/07/10 GV Add Start
'========================================================================
'
'	Function	�ŋ߃`�F�b�N�������i�ꗗ
'
'========================================================================
'
Function CreateViewedProductList()

Dim RSv
Dim vHTML
Dim vSQL

Dim vPrice
'Dim vCnt
Dim vName

NAVIViewedProductListHTML = ""

'---- �ŋ߃`�F�b�N�������i ���o��
vSQL = ""
'�\��������10������5���ɕύX	'2012/07/20 ok Mod
'vSQL = vSQL & "SELECT TOP 10"
vSQL = vSQL & "SELECT TOP 5"
vSQL = vSQL & "       a.���[�J�[�R�[�h"
vSQL = vSQL & "     , a.���i�R�[�h"
vSQL = vSQL & "     , a.�F"
vSQL = vSQL & "     , a.�K�i"
vSQL = vSQL & "     , b.���i�摜�t�@�C����_��"
vSQL = vSQL & "     , b.���i��"
vSQL = vSQL & "     , c.���[�J�[��"
vSQL = vSQL & "  FROM �ŋ߃`�F�b�N�������i a WITH (NOLOCK)"
vSQL = vSQL & "     , Web���i b WITH (NOLOCK)"
vSQL = vSQL & "     , ���[�J�[ c WITH (NOLOCK)"
vSQL = vSQL & " WHERE b.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "   AND b.���i�R�[�h = a.���i�R�[�h"
vSQL = vSQL & "   AND c.���[�J�[�R�[�h = a.���[�J�[�R�[�h"
vSQL = vSQL & "   AND a.�ڋq�ԍ� = " & wUserID
vSQL = vSQL & " ORDER BY"
vSQL = vSQL & "       a.�`�F�b�N�� DESC"

'@@@@@@@@@@response.write(vSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open vSQL, Connection, adOpenStatic

If RSv.EOF Then
	Exit Function
End If

	vHTML = vHTML & "    <div class='detail_side_inner01'><div class='detail_side_inner02'>" & vbNewLine
	vHTML = vHTML & "      <div class='detail_side_inner_box'>" & vbNewLine
	vHTML = vHTML & "        <h4 class='detail_sub'>�ŋ߃`�F�b�N�������i</h4>" & vbNewLine
	vHTML = vHTML & "        <ul class='check_item'>" & vbNewLine

'vCnt = 0

Do Until RSv.EOF

'	vCnt = vCnt  + 1
	vName = ""
	
	
	vName = vName & RSv("���[�J�[��") & " / " & RSv("���i��")
	If Trim(RSv("�F")) <> "" AND Trim(RSv("�K�i")) <> ""then
		vName = vName & " / " & Trim(RSv("�F")) & " / " & Trim(RSv("�K�i"))
	End If

	If Trim(RSv("�F")) <> "" AND Trim(RSv("�K�i")) = ""then
		vName = vName & " / " & Trim(RSv("�F"))
	End If

	If Trim(RSv("�F")) = "" AND Trim(RSv("�K�i")) <> ""then
		vName = vName & " / " & Trim(RSv("�K�i"))
	End If
		
	'---- ���[�J�[���A���i��/�F/�K�i
	vHTML = vHTML & "    <li>" & vbNewLIne
	vHTML = vHTML & "      <p><a href='ProductDetail.asp?Item=" & RSv("���[�J�[�R�[�h") & "^" & Server.URLEncode(RSv("���i�R�[�h")) & "^" & Trim(RSv("�F")) & "^" & Trim(RSv("�K�i")) & "'>"

'	If vCnt <= 5 Then		'5�܂ł͉摜�\��
		'---- ���i�摜
	If RSv("���i�摜�t�@�C����_��") <> "" Then
		vHTML = vHTML & "<img src='prod_img/" & RSv("���i�摜�t�@�C����_��") & "' alt='" & Replace(vName,"'","&#39;")  & "' class='opover'>" 
	End If
'	End If

	vHTML = vHTML & vName & "</a></p>" & vbNewLine
	vHTML = vHTML & "    </li>" & vbNewLIne

	RSv.MoveNext

Loop

	vHTML = vHTML & "      </ul>" & vbNewLine
	vHTML = vHTML & "    </div>" & vbNewLine
	vHTML = vHTML & "  </div></div>" & vbNewLine

RSv.close

wViewHTML = vHTML

End function
'2012/07/10 GV Add End

'2012/10/30 nt Add Start
'========================================================================
'
'	Function	�֘A�R���e���c�擾
'
'========================================================================
Function CreateContentsHTML()

Dim RSv
Dim vCnt

'---- �R���e���c�擾
wSQL = ""
wSQL = wSQL & "SELECT DISTINCT TOP 3"
wSQL = wSQL & "       b.�R���e���c��"
wSQL = wSQL & "     , b.URL"
wSQL = wSQL & "     , b.����"
wSQL = wSQL & "     , b.�摜�t�@�C����"
wSQL = wSQL & "     , a.�֘A�敪"
wSQL = wSQL & "     , b.�D�揇��"
wSQL = wSQL & "  FROM ���i�R���e���c a WITH (NOLOCK)"
wSQL = wSQL & "     , �R���e���c b WITH (NOLOCK)"
wSQL = wSQL & " WHERE a.�R���e���c�ԍ� = b.�R���e���c�ԍ�"
'wSQL = wSQL & "   AND b.�����N�t���O = 'Y'"
wSQL = wSQL & "   AND ((a.���[�J�[�R�[�h = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.���i�R�[�h = '" & Replace(product_cd, "'", "''") & "'"
wSQL = wSQL & "   AND a.�֘A�敪='1')"
'wSQL = wSQL & "   OR  ('" & Replace(product_cd, "'", "''") & "' LIKE a.���i�R�[�h + '%'"
'wSQL = wSQL & "   AND a.�֘A�敪='2')"
wSQL = wSQL & "   OR  (a.�V���[�Y�R�[�h='" & wSeriesCd & "'"
wSQL = wSQL & "   AND a.�֘A�敪='3')"
wSQL = wSQL & "   OR  (a.���[�J�[�R�[�h='" & maker_cd & "'"
wSQL = wSQL & "   AND a.�֘A�敪='4')"
wSQL = wSQL & "   OR  (a.�J�e�S���[�R�[�h='" & wCategoryCode & "'"
wSQL = wSQL & "   AND a.�֘A�敪='5')"
wSQL = wSQL & "   OR  (a.���J�e�S���[�R�[�h='" & wMidCategoryCd & "'"
wSQL = wSQL & "   AND a.�֘A�敪='6')"
wSQL = wSQL & "   OR  (a.��J�e�S���[�R�[�h='" & wLargeCategoryCd & "'"
wSQL = wSQL & "   AND a.�֘A�敪='7')"
wSQL = wSQL & "   OR  (a.�J�e�S���[�R�[�h='" & wCategoryCode & "'"
wSQL = wSQL & "   AND a.���[�J�[�R�[�h = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.�֘A�敪='8')"
wSQL = wSQL & "   OR  (a.���J�e�S���[�R�[�h='" & wMidCategoryCd & "'"
wSQL = wSQL & "   AND a.���[�J�[�R�[�h = '" & maker_cd & "'"
wSQL = wSQL & "   AND a.�֘A�敪='9'))"
wSQL = wSQL & " ORDER BY"
wSQL = wSQL & "       b.�D�揇�� ASC, a.�֘A�敪 ASC"

'@@@@@@@ Debug
'response.write(wSQL)

Set RSv = Server.CreateObject("ADODB.Recordset")
RSv.Open wSQL, Connection, adOpenStatic

wContentsHTML = ""

If RSv.EOF = True Then
	RSv.close
	Exit Function
End If

wHTML = ""
wHTML = wHTML & "<div class='detail_side_inner01'><div class='detail_side_inner02'>" & vbNewLine
wHTML = wHTML & "<div class='detail_side_inner_box'>" & vbNewLine
wHTML = wHTML & "  <h4 class='detail_sub'>���̏��i�Ɋ֘A����Z���N�V����</h4>" & vbNewLine
wHTML = wHTML & "  <ul class='special'>" & vbNewLine

Do Until RSv.EOF = True
	wHTML = wHTML & "    <li>" & vbNewLine

	If RSv("�摜�t�@�C����") <> "" Then
		wHTML = wHTML & "        <p class='photo'><a href='" & RSv("URL") & "'><img src='../" & RSv("�摜�t�@�C����") & "' alt='" & RSv("�R���e���c��") & "' class='opover' width='120' height='90' /></a></p>" & vbNewLine
	End If
	wHTML = wHTML & "          <p class='txt'><a href='" & RSv("URL") & "'>" & RSv("�R���e���c��") & "</a></p>" & vbNewLine

	wHTML = wHTML & "    </li>" & vbNewLine

	RSv.MoveNext
Loop

wHTML = wHTML & "  </ul>" & vbNewLine
wHTML = wHTML & "</div>" & vbNewLine
wHTML = wHTML & "</div></div>" & vbNewLine


wContentsHTML = wContentsHTML & wHTML

RSv.close

End Function
'2012/10/30 nt Add End

'========================================================================
%>
<!DOCTYPE html>
<html lang="ja">
<head prefix="og: http://ogp.me/ns# fb: http://ogp.me/ns/fb#">
<meta charset="Shift_JIS">
<meta name="robots" content="noindex,nofollow">
<link rel="canonical" href="http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=<%=Server.URLEncode(maker_cd & "^" & product_cd & "^" & iro & "^" & kikaku)%>">
<title><%=wMakerName%>&gt;<%=wProductName%>�b�T�E���h�n�E�X</title>
<% if wTokucho <> "" then%><meta name="description" content="<%=wTokucho%>"><% end if %>
<meta name="keywords" content="<%=wLargeCategoryName%>,<%=wMidCategoryName%>,<%=wCategoryName%>,<%=wMakerName%>,<%=wProductName%>">
<meta name="twitter:card" content="product">
<meta name="twitter:url" content="http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=<%=Server.URLEncode(maker_cd & "^" & product_cd & "^" & iro & "^" & kikaku)%>">
<meta name="twitter:site" content="@soundhouse_jp">
<meta name="twitter:image:width" content="600">
<meta name="twitter:image:height" content="300">
<meta name="twitter:label1" content="<%=wTwPriceLabel%>">
<meta name="twitter:data1" content="<%=wTwPriceData%>">
<meta name="twitter:label2" content="�݌ɏ�">
<% if wTwInventoryData <> "" Then%><meta name="twitter:data2" content="<%=wTwInventoryData%>"><% Else %><meta name="twitter:data2" content="�T�C�g��������������"><% End If %>
<meta property="og:title" content="<%=wMakerName%>&gt;<%=wProductName%>�b�T�E���h�n�E�X">
<meta property="og:type" content="article">
<meta property="og:url" content="http://www.soundhouse.co.jp/shop/ProductDetail.asp?Item=<%=Server.URLEncode(maker_cd & "^" & product_cd & "^" & iro & "^" & kikaku)%>">
<meta property="og:image" content="<%=g_HTTP%>shop/prod_img/<%=wMainProdPic%>">
<% if wTokucho <> "" Then%><meta property="og:description" content="<%=wTokucho%>"><% Else %><meta property="og:description" content="<%=wMakerName%>&gt;<%=wProductName%>"><% End If %>
<meta property="og:site_name" content="�T�E���h�n�E�X">
<meta property="og:locale" content="ja_JP">
<meta property="fb:app_id" content="191447484218062">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/shop.css?20140401a" type="text/css">
<link rel="stylesheet" href="style/jquery.fancybox-1.3.4.css" type="text/css">
<link rel="stylesheet" href="Style/ProductDetail.css?20140401" type="text/css">
<link rel="stylesheet" href="style/ask.css?20140401a" type="text/css">
<script type="text/javascript">
//
// ====== 	Function:	order_onClick
//
function order_onClick(pForm){
	if (pForm.qt.value == ""){
		pForm.qt.value = 0;
	}else{
		if (numericCheck(pForm.qt.value) == false){
			pForm.qt.value = 0;
		}
	}

	if (pForm.qt.value <= 0){
		alert("���ʂ���͂��Ă���J�[�g�{�^���������Ă��������B");
		return false;
	}

	if (pForm.IroKikaku.length > 1){
		if (pForm.IroKikaku.selectedIndex == 0){
			alert("<%=wIroKikakuSelectMsg%>");
			return false;
		}else{			//�F�K�i�𑗐M�G���A�փZ�b�g
			pForm.Item.value = pForm.Item.value + "^" + pForm.IroKikaku.options[pForm.IroKikaku.selectedIndex].value;
		}
	}

	return true;

}
//
// ====== 	Function:	BuyTogether_onClick
//
function BuyTogether_onClick(pProd){

	if (pProd.checked == true){
		document.f_data.AdditionalItem.value = document.f_data.AdditionalItem.value + "," + pProd.value;	}else{
		document.f_data.AdditionalItem.value = document.f_data.AdditionalItem.value.replace("," + pProd.value,"");
	}
}

//
// ====== 	Function:	IroKikaku_onChange
//
function IroKikaku_onChange(pForm){

	var i;

	i = pForm.IroKikaku.selectedIndex;

	document.fIroKikaku.Item.value = document.fIroKikaku.Item.value + "^" + pForm.IroKikaku.options[i].value;
	document.fIroKikaku.submit();
}

//
// ====== 	Function:	review_onSubmit
//
function review_onSubmit(pForm){

	if (pForm.Title.value == ""){
		alert("�^�C�g������͂��Ă��������");
		return false;
	}
	if (pForm.HandleName.value == ""){
		alert("�����O����͂��Ă��������");
		return false;
	}
	if (pForm.Review.value == ""){
		alert("���r���[����͂��Ă��������");
		return false;
	}
	if (pForm.Review.value.length > 1000){
		alert("���r���[��������1000�����𒴂��Ă��܂���@1000�����ȓ��ł��肢���܂��B");
		return false;
	}
	if (pForm.Review.value.indexOf("ttp://",0)  > 0){
		alert("�����N���܂܂�Ă��܂��B���r���[�ւ̓����N�͓o�^�ł��܂���B");
		return false;
	}
	return true;
}

//
//	ReviewSankou_onClick		'2010/03/08 hn add
//
function ReviewSankou_onClick(pID, pItem, pSankou){

var vAction;

		vAction = "Review" + "Sankou.asp";
		document.fReviewSankou.ID.value = pID;
		document.fReviewSankou.Item.value = pItem;
		document.fReviewSankou.Sankou.value = pSankou;
		document.fReviewSankou.action = vAction;
    	document.fReviewSankou.submit();
}

</script>

</head>
<body>
<div id="fb-root"></div>
<script>(function(d, s, id) {
  var js, fjs = d.getElementsByTagName(s)[0];
  if (d.getElementById(id)) return;
  js = d.createElement(s); js.id = id;
  js.src = "//connect.facebook.net/ja_JP/all.js#xfbml=1&appId=191447484218062";
  fjs.parentNode.insertBefore(js, fjs);
}(document, 'script', 'facebook-jssdk'));</script>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
	<span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
	<!-- �R���e���cstart -->
	<div id="globalContents" itemscope itemtype="http://data-vocabulary.org/Product">
    	<div id="path_box"><div id="path_box_inner01"><div id="path_box_inner02">
			<p class="home"><a href="../"><img src="../images/icon_home.gif" alt="HOME"></a></p>
			<ul id="path">
				<!-- �p�������X�g -->
				<%=wTitleWithLink%>
			</ul>
		</div></div></div>
        
        <!-- ���i�ڍ� -->
        <h1 class="title"><span itemprop="brand"><%=wMakerName%></span> / <span itemprop="name"><%=wProductname%></span></h1>
		<div id="productdetail">
        
        <div id="detail">
        	
			<div id="detail_inner01"><div id="detail_inner02">
<!-- ���i�摜 -->
<%=wPictureHTML%>

<!-- �֘A�����N -->
<%=wKanrenLinkHTML%>

<div id="inner_box">

<!-- ���� -->
<%=wTokuchoHTML%>

<!-- �X�y�b�N -->
<%=wSpecHTML%>

<!-- ���R�����h2 -->
<%= wRecommendBuyJS %>

<!-- �I�v�V���� -->
<%if wOptionHTML <> "" then %>

<%=wOptionHTML%>

<% end if %>

<!-- �p�[�c -->
<%if wPartsHTML <> "" then %>

<%=wPartsHTML%>

<% end if %>

<!-- ���i���r���[ -->
<% if wReviewHTML <> "" then %>

<h2 id="review" class="detail_title">���i���r���[</h2>
<%=wReviewHTML%>

<% elseif wCanWriteReviewFl = "Y" and WriteReview <> "Y" then %>

<h2 class="detail_title">���r���[�𓊍e����</h2>
<ul class="btn_review">

<% end if %>

<% if wCanWriteReviewFl = "Y" then %>
<% '2013/05/17 GV #1507
'�����r���[�ҏW������\�����Ȃ��悤�Aif���Ő���
'if WriteReview = "Y" then %>
	<% if 1=0 then %>
  </ul>
					<h2 class="detail_title">���r���[�𓊍e����</h2>
					<div class="comment_box no_line">
						<form name="f_review" method="post" action="ReviewStore.asp" onSubmit="return review_onSubmit(this);">
						<table class="comment">
							<tr>
							 	<th><span class="pp">�]��</span></th>
							 	<td>
							 	<select name="Rating">
									<option value="5">��~5</option>
									<option value="4">��~4</option>
									<option value="3">��~3</option>
									<option value="2">��~2</option>
									<option value="1">��~1</option>
								</select>
								<span>�ő�5�܂�</span>
							</td>
							</tr>
							<tr>
								<th><span class="pp">�^�C�g��</span></th><td><input type="text" name="Title" id="Title" maxsize="50"></td>
							</tr>
							<tr>
		<% if wHandleName = "" then %>
								<th><span class="pp">�n���h���l�[��</span></th><td><input type="text" name="HandleName" id="HandleName" maxlength="30"></td>
		<% else %>
								<th><span class="pp">�n���h���l�[��</span></th><td><%=wHandleName%><input type="hidden" name="HandleName" id="HandleName" value="<%=wHandleName%>"></td>
		<% end if %>
							</tr>
							<tr>
								<th><span class="pp">�Z��</span></th><td class="address"><%=wPrefecture%>�i����o�^�Z�����\������܂��j</td>
							</tr>
							<tr>
								<th><span class="pp">���r���[</span><br>�i1000�����܂Łj</th><td><textarea name="Review" rows="8" cols="70" style="width:325px;ime-mode:auto;"></textarea></td>
							</tr>
						 </table>
                         <div class="submit">
                         	<div class="review_attention">
                            	<h4>�����r���[�́A���i�Ɋւ���R�����g�݂̂����肢���܂��B</h4>
                                <p>�ȉ��ɊY������ꍇ�A���Ђ̔��f�ɂč폜�A���������s�Ȃ��ꍇ���������܂��̂ŁA���炩���߂��������������B</p>
                                <ul>
                                	<li>���i�ɑ΂��Ă̕]���Ƃ͊֌W�̖����R�����g</li>
                                    <li>���̃��r���[�ɑ΂��Ă̈ӌ��A�R�����g</li>
                                    <li>��排�����A��������Ǝv����L�q</li>
                                </ul>
                            </div>
                            <input type="image" src="images/btn_review_submit.png" alt="���e����"><p>��U���e���ꂽ���r���[�͕ύX�ł��܂���B</p>
                            <input type="hidden" name="OrderNo" value="<%=OrderNo%>">
                            <input type="hidden" name="Item" value="<%=item%>">
                        </div>
                        </form>
					</div>

	<% else %> 
<%
'2013/05/17 GV #1507 modified start
'  <li><a href="ProductDetail.asp?Item=<%=item%'>&WriteReview=Y"><img src="images/btn_review_write.png" alt="���̏��i�̃��r���[������" class="opover"></a></li>
'  </ul>
Dim UrlEncodeItem
UrlEncodeItem = Server.URLEncode(Item)
%>
  <li><a href="<%=g_HTTPS%>Shop/ReviewWrite.asp?Item=<%=UrlEncodeItem%>"><img src="images/btn_review_write.png" alt="���̏��i�̃��r���[������" class="opover"></a></li>
</ul>
<%
'2013/05/17 GV #1507 modified end
%>
	<% end if %>
<% else %>
	<% if wReviewHTML <> "" Or (wCanWriteReviewFl = "Y" and WriteReview <> "Y")then %>
</ul>
	<% end if %>
<% end if %>
				 </div>
			</div></div>
		<!--/#detail --></div>
		
<!-- ��������E�� ====================================================== -->
    <div id="detail_side">
<!-- ���[�J�[/���i�� -->
<%=wProductHTML%>

<!-- �J�[�g��� -->
<%=wCartHTML%>
<!-- ��֏��i�\�� -->
<%=wSubItemHTML%>

<!-- �֘A�R���e���c -->
<%=wContentsHTML%>

<!--�@�֘A�}�b�v�ւ̃����N�@-->
<!--
    <div class="detail_side_inner01"><div class="detail_side_inner02">
      <div class="detail_side_inner_box">
        <ul class="check_item">
          <li><a href="../recommend/RecommendMap.asp?item=<%=item%>"><img src="images/btn_recommendmap.png" alt="�֘A�A�C�e�����}�b�v�ŕ\��" class="opover">�֘A�A�C�e�����}�b�v�ŕ\��</a></li>
        </ul>
      </div>
    </div></div>
-->

<!-- �֘A�V���[�Y���i -->
<%=wSeriesHTML%>

<!-- ���R�����h���� -->
<%=wRecommendJS%>

<!-- �ŋ߃`�F�b�N�������i -->
<%=wViewHTML%>

<!-- �F�K�i�I�����ꂽ�Ƃ���ProductDetail.asp���ČĂяo��-->
<form name="fIroKikaku" method="get" action="ProductDetail.asp">
	<input type="hidden" name="Item" value="<%=maker_cd%>^<%=product_cd%>">
</form>
<!-- ���r���[�@�͂�/�������@��ReviewSankou.asp�Ăяo���p 2010/03/08 hn add -->
<form name="fReviewSankou" method="post" action="">
	<input type="hidden" name="ID" value="">
	<input type="hidden" name="Item" value="">
	<input type="hidden" name="Sankou" value="">
</form>
        </div>
        </div>
      <!--/#contents --></div>
	<div id="globalSide">
	<!--#include file="../Navi/NaviSide.inc"-->
	<!--/#globalSide --></div>
<!--/#main --></div>
<!--#include file="../Navi/Navibottom.inc"-->
<div class="tooltip"><p>ASK</p></div>
<!--#include file="../Navi/NaviScript.inc"-->
<script type="text/javascript" src="jslib/jquery.fancybox-1.3.4.pack.js"></script>
<script type="text/javascript" src="jslib/jquery.easing.1.3.js"></script>
<script type="text/javascript" src="../jslib/jquery.carouFredSel-5.5.0-packed.js"></script>
<script type="text/javascript" src="jslib/ask.js?20140401a"></script>
<script type="text/javascript" src="jslib/ProductDetail.js?20130709"></script>
<script type="text/javascript" src="http://platform.twitter.com/widgets.js" charset="utf-8"></script>
</body>
</html>
