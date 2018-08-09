<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:output method="html" doctype-system="http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" doctype-public="-//W3C//DTD XHTML 1.0 Transitional//EN" />
<xsl:variable name="title" select="/rss/channel/title"/>
<xsl:variable name="feedUrl" select="/rss/channel/atom:Currentlink[@rel='self']/@href" xmlns:atom="http://purl.org/atom/ns#"/>

<xsl:template match="/">
<xsl:element name="html">
<head>
<title><xsl:value-of select="$title"/></title>
<style>
body {
	margin: 0;
	padding:0;
	background-color: #FFFFFF;
	background-image: url(/images/rssbg.gif);
	background-repeat: repeat-y;
	background-position: top center;
	font-family: Lucida Sans, Trebuchet MS, Helvetica, sans-serif;
	text-align:center;
}

img {
	border: none;
}

img#feedimage {
	display:block;
	float:left;
     padding:0 15px 15px 0;
     margin:15px 0 0 7px;
}

div#bodyfence {
	width:684px;
	padding:0;
	margin:0px auto;
	position:relative;
	text-align:center;
}

h1 {
	color:#900;
	font-weight: normal;
	padding:0;
	margin:0;
	line-height:100%;
	text-align:left;
	letter-spacing: -.06em;
}

h2 {
	font-weight: normal;
	color:#aaa;
	padding: 0;
	margin:0 0 10px 0;
	font-size:16px;
	text-align:left;
	letter-spacing: -0.03em;
}

h3 {
	float:left;
	font-size:14px;
	width:136px;
	text-align:right;
	padding-right:19px;
	margin:0;
}

h4 {
	padding: 8px 0 0 30px;
	margin: 0;
	font-size: 16px;
}

div#header {
	padding-top:15px;
	margin:0 0 0 176px;
}

div#subscribe {
	width:650px;
	border: 1px solid #ccc;
	margin:0 auto;
	padding:10px;
	text-align:left;
	clear:both;
}

div#webreaders {
	margin:0 0 9px 208px;
}

div#webreaders img {
	vertical-align:middle;
	margin-right:15px;
}

div#feeddemon {
	padding:2px;
	position:absolute;
	top:0;
	left:0;
}

div#subscribe p {
	margin:0 0 9px 160px;
	font-size:11px;
}

div#subscribe p.with {
	font-size:14px;
	color:#444;
	font-weight:bold;
}

p.with span {
	font-size:11px;
	font-weight:normal;
}

blockquote {
	margin:0 0 0 28px;
	padding:0;
	position:relative;
}

ul {
	padding:4px 0 4px 0;
	margin:0 0 9px 198px;
}

li {
	margin:0;
	padding:0;
	font-size:11px;
	line-height:110%;
}

p#ownerblurb {
	background-color:#ffffcc;
	border:1px solid #ddd;
	padding: 2px;
}

div#content {
	padding: 0;
	margin-left:146px;
	text-align:left;
}

div#content dt a, a:link, a:visited, a:active {
text-decoration: none;
}

div#content dd a, a:link, a:visited, a:active {
text-decoration: none;
}

dl {
	background-image: url(/images/item.gif);
	background-position: top left;
	background-repeat: no-repeat;
	padding-left: 24px;
}


dt {
	font-size:13px;
	font-weight: bold;
	margin-left: 6px;
	padding-bottom:4px;
}

dd {
	font-size:13px;
	font-weight:   normal;
	margin: 0 15px 0 6px;
	overflow: hidden;
	text-align:left;
}

div#footer {
	border-top: 1px solid #ccc;
	margin-left: auto;
	margin-right: auto;
	padding: 0 12px;
	font-size:11px;
	text-align: center;
	width:95%;	
}

div#bodyfence h1 a, a:link, a:visited, a:active {
	color: #990000;
	text-decoration: none;
}

div#bodyfence h1 a:hover {
	color: #990000;
	text-decoration: none;
}

a.btn, a.btn:link, a.btn:visited, a.btn:active {
	text-decoration:none;
	border: 1px outset;
	background: #eee;
	padding:2px 4px 2px 4px;
	color:black;
}

a.btn:hover {
	color:black;
}

a, a:link, a:visited, a:active  {
text-decoration: underline;
	color: #000099;
}

a:hover {
	color: red;
}
</style>
<link rel="alternate" type="application/rss+xml" title="RSS" href="{$feedUrl}" /> 
</head>
<xsl:apply-templates select="rss/channel"/>
</xsl:element>
</xsl:template>
<xsl:template match="channel">
	<body>
		<div id="bodyfence">
			<xsl:apply-templates select="image"/>
			<div id="header">
				<h1><a href="{link}"><xsl:value-of select="$title"/></a></h1>				<h2>an RSS feed powered </h2>
			</div>
		  <div id="subscribe">
				<h3>ҳ</h3>
			  <p>ڿҳ<strong><xsl:value-of select="$title"/></strong>ṩģңӣӾۺϷҳװˣңӣĶԶı£ʱƿرվݡ</p>
				<h3>ңӣĶ</h3>
			  <p><a href="http://www.potu.com/" target="_blank"><img src="http://www.potu.com/index/images/potu_logo.gif" alt="POTUܲͨ" border="0" longdesc="http://www.potu.com/" /></a><br />
		      ңӣĶصַ<a href="http://www.potu.com/" target="_blank"></a><a href="http://www.potu.com/index/potu_down.php" target="_blank">ܲͨRSSĶ</a><a href="http://www.potu.com/" target="_blank"></a></p>
				</div>
			<div id="content">
				<xsl:apply-templates select="item"/>
			</div>
			<div id="footer">
				<br /><p>This syndication service powered by </p><br />
			</div>
		</div>
	</body>
</xsl:template>

<xsl:template match="item">
	<xsl:if test="position() = 1">
		<h4 xmlns="http://www.w3.org/1999/xhtml"></h4>
	</xsl:if>
	<dl xmlns="http://www.w3.org/1999/xhtml">
		<dt>
			<a href="{link}"><xsl:value-of select="title"/></a>
		</dt>
		<dd>
			<xsl:value-of select="pubDate"/>
		</dd>
		<dd name="decodeable">
			<xsl:call-template name="outputContent"/>
		</dd>
	</dl>
</xsl:template>

<xsl:template match="image">
	<xsl:element name="img" namespace="http://www.w3.org/1999/xhtml">
		<xsl:attribute name="src"><xsl:value-of select="url"/></xsl:attribute>
		<xsl:attribute name="alt">Link to <xsl:value-of select="title"/></xsl:attribute>
		<xsl:attribute name="id">feedimage</xsl:attribute>
	</xsl:element>
	<xsl:text/>
</xsl:template>

<xsl:template name="outputContent">
	<xsl:choose>
		<xsl:when test="xhtml:body" xmlns:xhtml="http://www.w3.org/1999/xhtml">
			<xsl:copy-of select="xhtml:body/*"/>
		</xsl:when>
		<xsl:when test="xhtml:div" xmlns:xhtml="http://www.w3.org/1999/xhtml">
			<xsl:copy-of select="xhtml:div"/>
		</xsl:when>
		<xsl:when test="content:encoded" xmlns:content="http://purl.org/rss/1.0/modules/content/">
			<xsl:value-of select="content:encoded" disable-output-escaping="yes"/>
		</xsl:when>
		<xsl:when test="description">
			<xsl:value-of select="description" disable-output-escaping="yes"/>
		</xsl:when>
	</xsl:choose>
</xsl:template>
</xsl:stylesheet>