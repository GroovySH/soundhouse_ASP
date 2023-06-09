<%@ LANGUAGE="VBScript" %>
<%
 Option Explicit
Response.Status="301 Moved Permanently" 
Response.AddHeader "Location", "http://www.soundhouse.co.jp/search/index?i_type=c&s_category_cd=513"
%>
<!--#include file="../common/ADOVBS.inc"-->
<!--#include file="../common/system_common.inc"-->
<!--#include file="../common/shop_common_functions.inc"-->
<!--#include file="../common/bfunctions1.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta charset="Shift_JIS">
<title>プレミアムギター｜サウンドハウス</title>
<meta name="description" content="世界を代表するギターメーカーが威信をかけて送り出すプレミアムギターの数々。歴史的名品、この世に1本しか存在しないカスタムメイドギター、超希少な木材を惜しげもなく用いた限定モデルなど、担当スタッフが厳選した逸品を一挙にご紹介。">
<meta name="keywords" content="プレミアムギター,ベース,カスタムメイド,限定,premium,guitar">
<!--#include file="../Navi/NaviStyle.inc"-->
<link rel="stylesheet" href="style/PremiumGuitars.css" type="text/css">
<script src="JSlib/FlashPlayerVersionDetection.js" type="text/javascript"></script>
<style type="text/css">
#globalContents ul.sns {
	overflow: hidden;
	padding: 5px;
}

#globalContents ul.sns li {
	float: right;
	width: 100px;
	height: 20px;
}
</style>
</head>
<body>
<!--#include file="../Navi/Navitop.inc"-->

<div id="globalMain">
  <span class="guidance"><a name="contents_start" id="contents_start"><img src="../images/spacer.gif" alt="ここから本文です"></a></span>

<!-- コンテンツstart -->
<div id="globalContents">
    <div id='path_box'><div id='path_box_inner01'><div id='path_box_inner02'>
    <p class='home'><a href="<%=g_HTTP%>"><img src="<%=g_RelLink%>images/icon_home.gif" alt="HOME"></a></p>
    <ul id='path'>
      <li><a href="<%=g_HTTP%>material/">SPECIAL SELECTION一覧</a></li>
      <li class="now">プレミアムギター</li>
    </ul>
  </div></div></div>
    <ul class="sns">
          <li><a href="https://twitter.com/share" class="twitter-share-button" data-lang="ja">ツイート</a><script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script></li>
          <li><iframe src="//www.facebook.com/plugins/like.php?href=http%3A%2F%2Fwww.soundhouse.co.jp%2Fshop%2FPremiumGuitars.asp&amp;send=false&amp;layout=button_count&amp;width=100&amp;show_faces=false&amp;action=like&amp;colorscheme=light&amp;font&amp;height=21&amp;appId=191447484218062" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:100px; height:21px;" allowTransparency="true"></iframe></li>
        </ul>
<!--
  <h1 class="title">プレミアムギター</h1>
-->
  <div id="pgTopContainer">
    <div id="pgHead">
      <img src="images/PremiumGuitars/header_index.jpg" align="Premium Guitars">
    </div>
    <p class="pgCommentBox"> 世界を代表するギターメーカーが威信をかけて送り出すプレミアムギターの数々。歴史的名品、この世に1本しか存在しないカスタムメイドギター、超希少な木材を惜しげもなく用いた限定モデルなど、担当スタッフが厳選した逸品を一挙にご紹介するサイトです。最高品質を誇る、贅沢なギター＆ベースセレクションをじっくりとご覧ください。</p>
    <div id="pgEnterBtn">
      <a href="PremiumGuitarsList.asp"><img src="images/PremiumGuitars/enter.jpg"></a>
    </div>
    <div id="pg_top_main_fla">
      <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="713" height="733">
        <param name="movie" value="flash/intro.swf">
        <param name="quality" value="high">
        <param name="wmode" value="transparent">
        <embed src="flash/intro.swf" wmode="transparent" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="713" height="733"></embed>
      </object>
	</div>
  </div>

  </div>
<div id="globalSide">
<!--#include file="../Navi/NaviSide.inc"-->
</div>
</div>
<!--#include file="../Navi/NaviBottom.inc"-->
<!--#include file="../Navi/NaviScript.inc"-->
</body>
</html>