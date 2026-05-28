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

<title><%=ana_ayarlar("site_adi")%></title>
<link rel="stylesheet" type="text/css" href="<%=ana_ayarlar("css_yolu")%>">

<%if ana_ayarlar("yazi_secme_yasagi")=1 then%>
<script language="Javascript">
function disableselect(e){
return false
}
function reEnable(){
return true
}
//if IE4+
document.onselectstart=new Function ("return false")
//if NS6
if (window.sidebar){
document.onmousedown=disableselect
document.onclick=reEnable
}
</script>
<%end if%>
<%if ana_ayarlar("sag_tus_kilidi")=1 then%>
<SCRIPT language=JavaScript1.2>
if (window.Event) // Only Netscape will have the CAPITAL E.
document.captureEvents(Event.MOUSEUP); // catch the mouse up event
function nocontextmenu() // this function only applies to IE4, ignored otherwise.
{
event.cancelBubble = true
event.returnValue = false;
return false;
}
function norightclick(e) // This function is used by all others
{
if (window.Event) // again, IE or NAV?
{
if (e.which == 2 || e.which == 3)
return false;
}
else
if (event.button == 2 || event.button == 3)
{
event.cancelBubble = true
event.returnValue = false;
return false;
}
}
document.oncontextmenu = nocontextmenu; // for IE5+
document.onmousedown = norightclick; // for all others
</SCRIPT>
<%end if%>
</head>
<link rel="stylesheet" type="text/css" href="css/soultip.css">
<script language="javascript" src="soultip.js" type="text/javascript"></script>
