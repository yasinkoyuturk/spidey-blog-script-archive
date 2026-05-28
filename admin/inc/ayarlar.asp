<%
dim ana_ayarlar
Set ana_ayarlar = Server.Createobject("ADODB.Recordset")
ana_ayarlar.Open "SELECT * FROM ana_ayarlar" ,BAG,1,3

dim kull_ayar
Set kull_ayar = Server.Createobject("ADODB.Recordset")
kull_ayar.Open "SELECT * FROM uyeler" ,BAG,1,3
%>
<head>
<meta name="keywords" content="<%=ana_ayarlar("site_anahtar_kelimeleri")%>">
<meta name="description" content="<%=ana_ayarlar("site_aciklama")%>">
<meta name="generator" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="author" content="Yasin Koyutürk"> 
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Page-Enter" content="blendTrans(Duration=0.5)" /> 
<meta http-equiv="Page-Exit" content="blendTrans(Duration=0.5)" />
<link rel="Shortcut Icon" href="favicon.ico">

<link rel="stylesheet" type="text/css" href="../<%=ana_ayarlar("css_yolu")%>">

</head>
<link rel="stylesheet" type="text/css" href="../css/soultip.css">
<script language="javascript" src="soultip.js" type="text/javascript"></script>
