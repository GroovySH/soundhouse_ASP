<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<!--#include file="../3rdParty/EAgency.inc"-->
<%
'========================================================================
'
'	�u���i���݂���܂���v�y�[�W
'�X�V����
'2013/05/20 GV #1505 ���Ԃ݂��ƁI���R�����h�@�\
'
'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<meta name="Description" content="�y��APA�����@��ADJ�EDTM�A�Ɩ��@��A�J���I�P�@�ނ��ǂ������y���������z�ł��񋟂���S���ʔ̃T�C�g�ł��B�y��A�����@��̂��ƂȂ�T�E���h�n�E�X�ɂ��C�����������I">
<meta name="keywords" content="�y��,����,�@��,DJ,DTM,�T�E���h�n�E�X">
<title>���i��������܂���b�T�E���h�n�E�X</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/otherlist.css" type="text/css">
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="��������{���ł�"></a></span>
  <!-- �R���e���cstart -->
  <div id="globalContents">
    <h1 class="title">���i��������܂���</h1>
    <p>�ǂ̃J�e�S���[�̏��i�����T���ł����H<br>���L�J�e�S���[�ꗗ���炨�T���̏��i��������Ȃ��ꍇ�́A���C�y��<a href="<%=g_HTTPS%>shop/Inquiry.asp"><strong>���₢���킹</strong></a>���������B</p>
<%
'2013/05/20 GV #1505 add start
Call fEAgency_CreateRecommendSearchNotFoundJS()
'2013/05/20 GV #1505 add end
%>
    <h2>�S�J�e�S���[�ꗗ</h2>
    <ul class="otherlist">
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=1"><img src="../Navi/Images/side/gnav_pa.png" alt="PA&amp;���R�[�f�B���O" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=12"><img src="../Navi/Images/side/gnav_guitar.png" alt="�M�^�[" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=13"><img src="../Navi/Images/side/gnav_bassa.png" alt="�x�[�X" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=14"><img src="../Navi/Images/side/gnav_drum.png" alt="�h����&amp;�p�[�J�b�V����" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=15"><img src="../Navi/Images/side/gnav_keyboard.png" alt="�V���Z�T�C�U�[�E�L�[�{�[�h" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=16"><img src="../Navi/Images/side/gnav_otherinstrumentsa.png" alt="���̑� �y��" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=7"><img src="../Navi/Images/side/gnav_djvja.png" alt="DJ &amp; VJ" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=8"><img src="../Navi/Images/side/gnav_dtmdawa.png" alt="DTM / DAW" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=3"><img src="../Navi/Images/side/gnav_recorder.png?20140618" alt="�f���@��E���R�[�_�[" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=4"><img src="../Navi/Images/side/gnav_lighting.png" alt="�Ɩ��E�X�e�[�W�V�X�e��" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=9"><img src="../Navi/Images/side/gnav_stand.png" alt="�X�^���h�e��" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=5"><img src="../Navi/Images/side/gnav_rack.png" alt="���b�N�E�P�[�X" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=10"><img src="../Navi/Images/side/gnav_cable.png" alt="�P�[�u���e��" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=6"><img src="../Navi/Images/side/gnav_headphone.png" alt="�w�b�h�z���E�C���z��" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=51"><img src="../Navi/Images/side/gnav_furniture.png" alt="�X�^�W�I�Ƌ�" class="imgover"></a></li>
    </ul>

    <h2>�l�C�̃R���e���c</h2>
    <ul class="linklist">
      <li><a href="http://hotplaza.soundhouse.co.jp/otoya_movie/"><img src="../top/ura/otoyamovie_bana.png" alt="�����Check ! OTOYA MOVIE" class="opover">�����Check�I<br>OTOYA MOVIE</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/material/serviceman_diary/serviceman_diary.asp"><img src="../top/ura/service.jpg" alt="�T�[�r�X�}�������I�C���S���̋Ɩ�����" class="opover">�T�[�r�X�}�������I<br>�C���S���̋Ɩ�����</a></li>
      <!--<li><a href="http://hotplaza.soundhouse.co.jp/report/index.asp"><img src="../top/ura/sijo.jpg" alt="�C�O�G���Ɍf�ڂ��ꂽ���i���|�[�g���Љ�I" class="opover">�C�O�G���Ɍf�ڂ��ꂽ<br>���i���|�[�g���Љ�I</a></li>-->
      <li><a href="http://www.soundhouse.co.jp/shop/ManualDownload.asp"><img src="../top/ura/mdl.jpg" alt="���i�}�j���A���͂�����Ń_�E�����[�h" class="opover">���i�}�j���A����<br>������Ń_�E�����[�h</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/present/present.asp"><img src="../top/ura/hot_plaza_present.jpg" alt="�����̃v���[���g" class="opover">�����̃v���[���g</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/mi_maintenance/index.asp"><img src="../top/ura/guiter_02.jpg" alt="�M�^�[�̃C���n�̓M�^���X�g�̕�����" class="opover">�M�^�[�̃C���n��<br>�M�^���X�g�̕�����</a></li>
    </ul>
    <ul class="linklist">
      <li><a href="http://hotplaza.soundhouse.co.jp/dj_guide/index.asp"><img src="../top/ura/djmyles_guide_bana.jpg" alt="�ꂩ��o����DJ����" class="opover">�ꂩ��o����DJ����</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/bass_guide/index.asp"><img src="../top/ura/howtobass.jpg" alt="�G���L�x�[�X����u��" class="opover">�G���L�x�[�X����u��</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/drumm_guide/index.asp"><img src="../top/ura/howtodrum.jpg" alt="�h�����u��" class="opover">�h�����u��</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/dtm_guide/index.asp"><img src="../top/ura/dtm.jpg" alt="DTM�EDAW�w���K�C�h" class="opover">DTM�EDAW�w���K�C�h</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/how_to/light/index.asp"><img src="../top/ura/banner_howtolight.jpg" alt="�Ɩ�����u��" class="opover">�Ɩ�����u��</a></li>
    </ul>
    <ul class="linklist">
      <li><a href="http://hotplaza.soundhouse.co.jp/how_to/pa/"><img src="../top/ura/pa_guide_bana.png" alt="PA�V�X�e���u��" class="opover">PA�V�X�e���u��</a></li>
            <li><a href="http://hotplaza.soundhouse.co.jp/how_to/keyboard/"><img src="../top/ura/keyboard_guide_bana.jpg" alt="�L�[�{�[�h�E�s�A�m�u��" class="opover">�L�[�{�[�h�E�s�A�m�u��</a></li>
            <li><a href="http://hotplaza.soundhouse.co.jp/how_to/headphone/"><img src="../top/ura/headphone_guide_bana.jpg" alt="�w�b�h�z���E�C���z���u��" class="opover">�w�b�h�z���E�C���z���u��</a></li>
    </ul>

</div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>