<!--#include file="inc/baglan.asp" -->
<!--#include file="inc/ayarlar.asp" -->

<%if Session("blogadmin")=True Then '---admin giriþ yapmýþsa---

	islem=request.querystring("islem")
	proje_id=request.querystring("id")


if islem="proje_duzenle2" then

	proje_id=request.querystring("id")

	Set proje = Server.Createobject("ADODB.Recordset")
	sor="SELECT * FROM projelerim WHERE id="&proje_id
	proje.Open sor,BAG,1,3
	
	proje("ad") = Trim(Request.Form("ad"))
	proje("seviye") = Trim(Request.Form("seviye"))
	proje("amac") = Trim(Request.Form("amac"))
	proje("durum") = Trim(Request.Form("durum"))
	proje("aciklama") = Trim(Request.Form("aciklama"))
	proje("kisa_aciklama") = Trim(Request.Form("kisa_aciklama"))
	
	proje.update
	Set proje= Nothing
	Response.Redirect "default.asp?goster=proje_ayarlari&islem=proje_duzenlendi"
end if

Set proje = Server.Createobject("ADODB.Recordset")
proje.Open "SELECT * FROM projelerim WHERE id="&proje_id&"" ,BAG,1,3
%>
<form action="proje_duzenle.asp?islem=proje_duzenle2&id=<%=proje("id")%>" method="POST" id="asdsad">
<table border="0" cellpadding="0" cellspacing="0" id="table2">
<tr>
	<td width="722" colspan="2" height="1"><a href="default.asp" style="text-decoration: none"><center><br>
<a href="default.asp"><span class="normal_r2yazi">Yönetici Bölümü Anasayfa</span></a>
<br></center>

</td>
</tr>
<tr>
	<td width="722" colspan="2">
	<font color="#FF0000">
	<span style="font-family: verdana; font-weight: 700; font-size: 8pt">-Proje düzenle-</span></font></td>
</tr>
<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Proje Adý</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="ad" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="<%=proje("ad")%>"></td>
</tr>
<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Seviye</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="seviye" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="<%=proje("seviye")%>"></td>
</tr>
		<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Amacý</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="amac" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="<%=proje("amac")%>"></td>
		</tr>
		<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Durum</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="durum" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="<%=proje("durum")%>"></td>
		</tr>
<tr>
	<td width="82">
	<font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">
	Kýsa Açýklama:</font>
	</td>
	<td width="640">
	<input type="text" name="kisa_aciklama" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="<%=proje("kisa_aciklama")%>"></td>
		</tr>
<tr>
	<td width="82">
	<font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">
	Proje Açýklamasý:</font>
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
<textarea name="aciklama" style="width:670; height:300" rows="1" cols="20">
<%=proje("aciklama")%></textarea>
</font></b>
<script language="javascript1.2">
editor_generate('aciklama');
</script> 
<!--  -->
	</td>
	<input type="hidden" name="ekleyen" value="<%=kull_ayar("kullanici_adi")%>">
</tr>
<tr>
	<td width="82">
	&nbsp;</td>
	<td width="640">
	<input type="submit" value="Kaydet" style="cursor: hand; font-size: 7pt; font-family: Verdana; border-width: 1;"></form>
</td>
</table>
<%else
Response.Redirect "default.asp"
end if
%>