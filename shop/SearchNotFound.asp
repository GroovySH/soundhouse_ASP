<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<!--#include file="../3rdParty/EAgency.inc"-->
<%
'========================================================================
'
'	「商品がみつかりません」ページ
'更新履歴
'2013/05/20 GV #1505 さぶみっと！レコメンド機能
'
'========================================================================
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<meta name="Description" content="楽器、PA音響機器、DJ・DTM、照明機器、カラオケ機材をどこよりも【激安特価】でご提供する全国通販サイトです。楽器、音響機器のことならサウンドハウスにお任せください！">
<meta name="keywords" content="楽器,音響,機材,DJ,DTM,サウンドハウス">
<title>商品が見つかりません｜サウンドハウス</title>
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/otherlist.css" type="text/css">
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>
  <!-- コンテンツstart -->
  <div id="globalContents">
    <h1 class="title">商品が見つかりません</h1>
    <p>どのカテゴリーの商品をお探しですか？<br>下記カテゴリー一覧からお探しの商品が見つからない場合は、お気軽に<a href="<%=g_HTTPS%>shop/Inquiry.asp"><strong>お問い合わせ</strong></a>ください。</p>
<%
'2013/05/20 GV #1505 add start
Call fEAgency_CreateRecommendSearchNotFoundJS()
'2013/05/20 GV #1505 add end
%>
    <h2>全カテゴリー一覧</h2>
    <ul class="otherlist">
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=1"><img src="../Navi/Images/side/gnav_pa.png" alt="PA&amp;レコーディング" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=12"><img src="../Navi/Images/side/gnav_guitar.png" alt="ギター" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=13"><img src="../Navi/Images/side/gnav_bassa.png" alt="ベース" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=14"><img src="../Navi/Images/side/gnav_drum.png" alt="ドラム&amp;パーカッション" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=15"><img src="../Navi/Images/side/gnav_keyboard.png" alt="シンセサイザー・キーボード" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=16"><img src="../Navi/Images/side/gnav_otherinstrumentsa.png" alt="その他 楽器" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=7"><img src="../Navi/Images/side/gnav_djvja.png" alt="DJ &amp; VJ" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=8"><img src="../Navi/Images/side/gnav_dtmdawa.png" alt="DTM / DAW" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=3"><img src="../Navi/Images/side/gnav_recorder.png?20140618" alt="映像機器・レコーダー" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=4"><img src="../Navi/Images/side/gnav_lighting.png" alt="照明・ステージシステム" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=9"><img src="../Navi/Images/side/gnav_stand.png" alt="スタンド各種" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=5"><img src="../Navi/Images/side/gnav_rack.png" alt="ラック・ケース" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=10"><img src="../Navi/Images/side/gnav_cable.png" alt="ケーブル各種" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=6"><img src="../Navi/Images/side/gnav_headphone.png" alt="ヘッドホン・イヤホン" class="imgover"></a></li>
      <li><a href="LargeCategoryList.asp?LargeCategoryCd=51"><img src="../Navi/Images/side/gnav_furniture.png" alt="スタジオ家具" class="imgover"></a></li>
    </ul>

    <h2>人気のコンテンツ</h2>
    <ul class="linklist">
      <li><a href="http://hotplaza.soundhouse.co.jp/otoya_movie/"><img src="../top/ura/otoyamovie_bana.png" alt="動画でCheck ! OTOYA MOVIE" class="opover">動画でCheck！<br>OTOYA MOVIE</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/material/serviceman_diary/serviceman_diary.asp"><img src="../top/ura/service.jpg" alt="サービスマンが語る！修理担当の業務日誌" class="opover">サービスマンが語る！<br>修理担当の業務日誌</a></li>
      <!--<li><a href="http://hotplaza.soundhouse.co.jp/report/index.asp"><img src="../top/ura/sijo.jpg" alt="海外雑誌に掲載された製品レポートを紹介！" class="opover">海外雑誌に掲載された<br>製品レポートを紹介！</a></li>-->
      <li><a href="http://www.soundhouse.co.jp/shop/ManualDownload.asp"><img src="../top/ura/mdl.jpg" alt="製品マニュアルはこちらでダウンロード" class="opover">製品マニュアルは<br>こちらでダウンロード</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/present/present.asp"><img src="../top/ura/hot_plaza_present.jpg" alt="今月のプレゼント" class="opover">今月のプレゼント</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/mi_maintenance/index.asp"><img src="../top/ura/guiter_02.jpg" alt="ギターのイロハはギタリストの部屋で" class="opover">ギターのイロハは<br>ギタリストの部屋で</a></li>
    </ul>
    <ul class="linklist">
      <li><a href="http://hotplaza.soundhouse.co.jp/dj_guide/index.asp"><img src="../top/ura/djmyles_guide_bana.jpg" alt="一から覚えるDJ入門" class="opover">一から覚えるDJ入門</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/bass_guide/index.asp"><img src="../top/ura/howtobass.jpg" alt="エレキベース入門講座" class="opover">エレキベース入門講座</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/drumm_guide/index.asp"><img src="../top/ura/howtodrum.jpg" alt="ドラム講座" class="opover">ドラム講座</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/dtm_guide/index.asp"><img src="../top/ura/dtm.jpg" alt="DTM・DAW購入ガイド" class="opover">DTM・DAW購入ガイド</a></li>
      <li><a href="http://hotplaza.soundhouse.co.jp/how_to/light/index.asp"><img src="../top/ura/banner_howtolight.jpg" alt="照明入門講座" class="opover">照明入門講座</a></li>
    </ul>
    <ul class="linklist">
      <li><a href="http://hotplaza.soundhouse.co.jp/how_to/pa/"><img src="../top/ura/pa_guide_bana.png" alt="PAシステム講座" class="opover">PAシステム講座</a></li>
            <li><a href="http://hotplaza.soundhouse.co.jp/how_to/keyboard/"><img src="../top/ura/keyboard_guide_bana.jpg" alt="キーボード・ピアノ講座" class="opover">キーボード・ピアノ講座</a></li>
            <li><a href="http://hotplaza.soundhouse.co.jp/how_to/headphone/"><img src="../top/ura/headphone_guide_bana.jpg" alt="ヘッドホン・イヤホン講座" class="opover">ヘッドホン・イヤホン講座</a></li>
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