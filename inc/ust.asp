<!--#include file="baglan.asp" -->
<!--#include file="ayarlar.asp" -->

<head>
<meta http-equiv="Content-Language" content="tr">
</head>

<body topmargin="<%=ana_ayarlar("ust_bosluk")%>" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" text="#000000" onMouseOver="status='<%=trim(ana_ayarlar("durum_yazisi"))%>'; return true">
<div align="center">


<table border="0" cellpadding="0" cellspacing="0" width="700" id="table1" height="100%">
	<tr>
		<td width="5" bgcolor="#C0C0C0">&nbsp;</td>
		<td width="690" height="100" colspan="5" background="<%=ana_ayarlar("site_logosu")%>">&nbsp;</td>
		<td width="5" bgcolor="#C0C0C0">&nbsp;</td>
	</tr>
	<tr>
		<td width="5" height="13" bgcolor="#C0C0C0">&nbsp;</td>
		<td width="550" height="13" colspan="5" bgcolor="#C0C0C0">
		<span id="digitalclock" class="silik_yazi"><%
bugun= Day(date) 
ay=Month(date) 
yil=Year(date) 
aylar = Array("yok","Ocak","Þubat","Mart","Nisan","Mayýs", "Haziran","Temmuz","Aðustos","Eylül","Ekim","Kasým","Aralýk") 
ayname = aylar(ay)

gun = Weekday(Weekday(date))
gunler = Array("yok","Pazar","Pazartesi","Salý","Çarþamba", "Perþembe","Cuma","Cumartesi") 
%>

<script>
<!--
var alternate=0
var standardbrowser=!document.all&&!document.getElementById

function show(){
if (!standardbrowser)
var clockobj=document.getElementById? document.getElementById("digitalclock") : document.all.digitalclock
var Digital=new Date()
var hours=Digital.getHours()
var minutes=Digital.getMinutes()

if (hours==0) hours=12
if (hours.toString().length==1)
hours="0"+hours
if (minutes<=9)
minutes="0"+minutes

if (standardbrowser){
if (alternate==0)
document.tick.tock.value=hours+":"+minutes
else
document.tick.tock.value=hours+" "+minutes
}
else{
if (alternate==0)
clockobj.innerHTML="<span class=normal_yazi2><%=bugun%>&nbsp;<%=ayname%>&nbsp;<%=yil%>,&nbsp;<%=gunler(gun)%>&nbsp;&nbsp;</span><span class=normal_yazi2>"+hours+"</span><span class=normal_ryazi>:</span><span class=normal_yazi2>"+minutes+"</span>"
else
clockobj.innerHTML="<span class=normal_yazi2><%=bugun%>&nbsp;<%=ayname%>&nbsp;<%=yil%>,&nbsp;<%=gunler(gun)%>&nbsp;&nbsp;</span><span class=normal_yazi2>"+hours+"</span><span class=normal_yazi>:</span><span class=normal_yazi2>"+minutes+"</span>"
}
alternate=(alternate==0)? 1 : 0
setTimeout("show()",1000)
}
window.onload=show

//-->
</script></span>
		</td>
		<td width="5" height="13" bgcolor="#C0C0C0">&nbsp;</td>
	</tr>
	<tr>
		<td width="5" bgcolor="#C0C0C0">&nbsp;</td>

		<td width="550" height="100%" valign="top">