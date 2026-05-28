<!--#include file="inc/baglan.asp" -->
<!--#include file="inc/ayarlar.asp" -->

<%if Session("blogadmin")=True Then '---admin giriþ yapmýþsa---


yazi=request.querystring("yazi")
kategori=request.querystring("kategori")

set alt = Server.CreateObject("ADODB.Recordset")
sor = "SELECT * FROM makale_alt_kategoriler WHERE id="& yazi
alt.Open sor,bag,1,3

set kat = Server.CreateObject("ADODB.Recordset")
sor2 = "SELECT * FROM makale_kategoriler WHERE kate_id="& kategori
kat.Open sor2,bag,1,3

%>
<center><br>
<a href="default.asp"><span class="normal_r2yazi">Yönetici Bölümü Anasayfa</span></a><span class="normal_kyazi"> > </span>
<a href="blog_admin.asp"><span class="normal_r2yazi">Blog Admin Anasayfa</span></a><br>
<br></center>
<table border="0" cellpadding="0" cellspacing="0" width="621" id="table1">
	<tr>
		<td bgcolor="#CCCCCC" height="20">
<form action="blog_admin_goster.asp?gorev=blog_duzenle&yazi=<%=yazi%>&kategori=<%=kategori%>" method="POST">
<table>
<tr>
	<td width="722" colspan="2" height="1">
	<hr size="1"></td>
</tr>
<tr>
	<td width="722" colspan="2">
	<font color="#FF0000">
	<span style="font-family: verdana; font-weight: 700; font-size: 8pt">-Makale 
	Düzenle-</span></font></td>
</tr>
<tr>
	<td width="82">
	<font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">Baþlýk:</font>
	</td>
	<td width="640">
	<input type="text" name="ad" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="<%=alt("ad")%>"></td>
</tr>
<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Açýklama</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="aciklama" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="<%=alt("aciklama")%>"></td>
</tr>
<tr>
	<td width="82">
	<font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">Makale</font>
	</td>
	<td width="640">
<!-- EDÝTÖR -->
<script language="Javascript1.2"><!-- // load htmlarea
_editor_url = "../";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
  document.write('<scr' + 'ipt src="' +_editor_url+ 'editor.js"');
  document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }
// --></script>
<textarea name="makale_body" style="width:670; height:300" rows="1" cols="20"><%=alt("makale_body")%></textarea>
</font></b>
<script language="javascript1.2">
editor_generate('makale_body');
</script> 
<!--  -->
	</td>
	<input type="hidden" name="ekleyen" value="<%=alt("ekleyen")%>">
</tr>
<tr>
	<td width="82">
	<font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">
	Þu anki Kategori</font>
	</td>
	<td width="640">
	<select size="1" name="kate_id" style="font-family: Verdana; font-size: 7pt; border: 1px solid #000000">
<%set kat1 = Server.CreateObject("ADODB.Recordset")
sor2 = "SELECT * FROM makale_kategoriler"
kat1.Open sor2,bag,1,3

for i=1 to kat1.recordcount%>
	<option value="<%=kat1("kate_id")%>" <%if alt("kate_id")=kat1("kate_id") then%>selected<%else%><%end if%>><%=kat1("kate_ad")%></option>
<%
kat1.movenext
next
%>
	</select></td>
</tr>
<tr>
	<td width="82">
	&nbsp;</td>
	<td width="640">
	<input type="submit" value="Güncelle" style="cursor: hand; font-size: 7pt; font-family: Verdana; border-width: 1;">
	</td>
	</tr>
	</table></form></td></tr></table>
<%
kat.close
set kat=nothing
alt.close
set alt=nothing

%>
<%else
Response.Redirect "default.asp"
end if%>