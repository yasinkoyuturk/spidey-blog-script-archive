<!-- #include file="inc/baglan.asp" -->
<!--#include file="inc/ayarlar.asp" -->
<%
Response.Buffer=True

Session.Codepage=1254
Session.LCID=1055
%>
<head>
<title><%=ana_ayarlar("site_adi")%> - Yönetici Bölümü</title>

<meta http-equiv="Content-Language" content="tr">
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<%
islem=request("islem")
goster=request("goster")
hata=request("hata")


if Session("blogadmin")=True Then

Set ana_ayarlar = Server.Createobject("ADODB.Recordset")
ana_ayarlar.Open "SELECT * FROM ana_ayarlar" ,BAG,1,3


if islem="cikis" then
	Session("blogadmin")="" 
	Response.Cookies("TekTech")("kullanici_adi")=""
	Response.Cookies("TekTech")("sifre")=""
	Session("id")="" 
	
	Response.Redirect "default.asp?islem=cikis_yapildi"
end if


if goster="ayarlar" then
call ayarlar

elseif goster="kullanici_ayarlari" then
call kullanici_ayarlari

elseif goster="proje_ayarlari" then
call proje_ayarlari

elseif goster="blog" then
response.redirect "blog_admin.asp"

elseif goster="baglantilar" then
call baglantilar_bolumu

elseif goster="resimler" then
call resimler

else
%>
<table border="0" cellpadding="0" cellspacing="0" bgcolor="#E1E1E1" id="table1">
	<tr>
		<td>

<table border="0" cellpadding="0" cellspacing="1" width="523" id="table2">
	<tr>
		<td height="2" width="521" align="right" bgcolor="#808080" colspan="5">
		<p align="left"><b>
		<font face="Tahoma" style="font-size: 8pt" color="#FFFFFF">&nbsp; Spidey 
		Blog Script v.1.5 - Yönetici Bölümü Anasayfasý</font></b></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<font color="#FF0000" face="Tahoma" style="font-size: 8pt">
		<a style="text-decoration: none" href="?goster=ayarlar">
		<font color="#FF0000">ayarlar</font></a></font></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
		<font face="Tahoma" style="font-size: 8pt">Spidey Blog Script v.10'ýn 
		temel ayarlarýný buradan deðiþtirebilirsiniz.</font></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<a style="text-decoration: none" href="?goster=kullanici_ayarlari">
		<font color="#FF0000" face="Tahoma" style="font-size: 8pt">
		kullanýcý
		<font color="#FF0000">ayarlar</font>ý</font></a></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
		<font face="Tahoma" style="font-size: 8pt">Kullanýcý adýnýzý, þifrenizi 
		vb. kiþisel bilgilerinizi deðiþtirebileceðiniz bölüm</font></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<a style="text-decoration: none" href="?goster=proje_ayarlari">
		<font face="Tahoma" style="font-size: 8pt" color="#FF0000">projelerim 
		bölümü</font></a></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
		<font face="Tahoma" style="font-size: 8pt">projelerinizi 
		yönetebileceðiniz bölüm</font></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<a style="text-decoration: none" href="?goster=blog">
		<font face="Tahoma" style="font-size: 8pt" color="#FF0000">blog bölümü
		</font></a></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
		<font face="Tahoma" style="font-size: 8pt">blogunuza yeni yazýlar 
		yazabilir, silebilir veyahut güncelleyebilirsiniz</font></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<a style="text-decoration: none" href="?goster=baglantilar">
		<font face="Tahoma" style="font-size: 8pt" color="#FF0000">baglantilar 
		bölümü</font></a></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
		<font face="Tahoma" style="font-size: 8pt">baðlantýlar bölümünü 
		düzenleyebilir, silebilir veya yeni baðlantýlar ekleyebilirsiniz</font></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<a style="text-decoration: none" href="?goster=resimler">
		<font face="Tahoma" style="font-size: 8pt" color="#FF0000">yeni resim 
		yükle</font></a></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
		<font face="Tahoma" style="font-size: 8pt">upload/resimler klasörüne 
		yeni bir resim yükleyebilirsiniz</font></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<font face="Tahoma" style="font-size: 8pt">
		<a style="text-decoration: none" href="?islem=cikis">sistemden çýkýþ</a></font></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
		<font face="Tahoma" style="font-size: 8pt">yönetici 
		bölümünden güvenli çýkýþ</font></td>
	</tr>
	</table>
		</td>
	</tr>
</table>

<%end if%>
<%
'---------------------AYARLAR-----------------------
sub ayarlar
%>
<%
site_ayar=request("site_ayar")

site_ayar=1
Set ana_ayarlar = Server.Createobject("ADODB.Recordset")
ana_ayarlar.Open "SELECT * FROM ana_ayarlar" ,BAG,1,3
%>
<center><br>
<a href="default.asp"><span class="normal_r2yazi">Yönetici Bölümü Anasayfa</span></a>
<br></center>
<form method="POST" action="default.asp?goster=ayarlar&ayarlar_kayit_onay=1" name="ayarlar">
<table border="0" cellpadding="0" style="border-collapse: collapse" width="296">
	<tr>
		<td colspan="2" align="center">Site Ayarlarý</td>
	</tr>
	<tr>
		<td width="148">Site ismi</td>
		<td width="148">
		<input type="text" name="site_adi" size="20" value="<%=ana_ayarlar("site_adi")%>"></td>
	</tr>
	<tr>
		<td width="148">Site Logosu (691x100)</td>
		<td width="148">
		<input type="text" name="site_logosu" size="20" value="<%=ana_ayarlar("site_logosu")%>"></td>
	</tr>
	<tr>
		<td width="148">Durum çubuðu yazýsý</td>
		<td width="148"><input type="text" name="durum_yazisi" size="20" value="<%=ana_ayarlar("durum_yazisi")%>"></td>
	</tr>
	<tr>
		<td width="148">Site açýklamasý</td>
		<td width="148">
		<input type="text" name="site_aciklama" size="20" value="<%=ana_ayarlar("site_aciklama")%>"></td>
	</tr>
	<tr>
		<td width="148">Site anahtar kelimeleri</td>
		<td width="148">
		<input type="text" name="site_anahtar_kelimeleri" size="20" value="<%=ana_ayarlar("site_anahtar_kelimeleri")%>"></td>
	</tr>
	<tr>
		<td width="148">Stil sayfasý yolu</td>
		<td width="148">
		<input type="text" name="css_yolu" size="20" value="<%=ana_ayarlar("css_yolu")%>"></td>
	</tr>
	<tr>
		<td width="148">Resimler klasörü Yolu</td>
		<td width="148">
		<input type="text" name="resim_yolu" size="20" value="<%=ana_ayarlar("resim_yolu")%>"></td>
	</tr>
	<tr>
		<td width="148">Site üst boþluðu</td>
		<td width="148">
		<input type="text" name="ust_bosluk" size="2" value="<%=ana_ayarlar("ust_bosluk")%>"></td>
	</tr>
	<tr>
		<td width="148">Sað týklama yasaðý</td>
		<td width="148"><input type="radio" name="sag_tiklama_yasagi" value="1" <%if ana_ayarlar("sag_tus_kilidi")=1 then%> checked<%else%><%end if%>>Açýk&nbsp;&nbsp;&nbsp;
		<input type="radio" name="sag_tiklama_yasagi" value="0" <%if ana_ayarlar("sag_tus_kilidi")=1 then%><%else%> checked<%end if%>>Kapalý</td>
	</tr>
	<tr>
		<td width="148">Yazý seçme yasaðý</td>
		<td width="148">
		<input type="radio" name="yazi_secme_yasagi" value="1"<%if ana_ayarlar("sag_tus_kilidi")=1 then%> checked<%else%><%end if%>>Açýk&nbsp;&nbsp;&nbsp;
		<input type="radio" name="yazi_secme_yasagi" value="0" <%if ana_ayarlar("sag_tus_kilidi")=1 then%><%else%> checked<%end if%>>Kapalý</td>
	</tr>
	<tr>
		<td width="148">
		<p align="center">
		<input type="submit" value="Güncelle" name="submit" style="font-family: Verdana; font-size: 7pt; font-weight: 700; color: #000000; border: 1px solid #000000; float:right"></td>
</form>
		<td width="148"><% 
ayarlar_kayit_onay=request("ayarlar_kayit_onay")
if ayarlar_kayit_onay=1 and 1 = 1 then

Set ana_ayarlar = Server.Createobject("ADODB.Recordset")
ana_ayarlar.Open "SELECT * FROM ana_ayarlar" ,BAG,1,3

'formdan gelen bilgileri yerlerine eþitleyelim
ana_ayarlar("site_adi") 				= request.form("site_adi")
ana_ayarlar("durum_yazisi") 			= request.form("durum_yazisi")
ana_ayarlar("durum_yazisi")				= request.form("durum_yazisi")
ana_ayarlar("site_aciklama") 			= request.form("site_aciklama")
ana_ayarlar("site_anahtar_kelimeleri") 	= request.form("site_anahtar_kelimeleri")
ana_ayarlar("css_yolu") 				= request.form("css_yolu")
ana_ayarlar("resim_yolu")				= request.form("resim_yolu")
ana_ayarlar("ust_bosluk") 				= request.form("ust_bosluk")
ana_ayarlar("sag_tus_kilidi") 			= request.form("sag_tiklama_yasagi")
ana_ayarlar("yazi_secme_yasagi") 		= request.form("yazi_secme_yasagi")
ana_ayarlar("site_logosu") 				= request.form("site_logosu")

'veritabanýmýzý upload edelim
ana_ayarlar.update

'sonrada serverýn hafýzasýndan yer kaplamamasý için silelim
ana_ayarlar.Close 
Set ana_ayarlar = Nothing

'son olarakta ayarlar sayfamýza geri dönelim
Response.write("<font color='red'>&nbsp;Ayarlar Güncellendi</font>")

else

Response.write("<font color='red'>&nbsp;Ýþlemi tamamlamak için<br>&nbsp;güncelle butonuna basýn</font>")

end if
%></td>
	</tr>
</table>
<%end sub%>
<%
'-------------KULLANICI AYARLARI-----------------------
sub kullanici_ayarlari

Set kull_ayar = Server.Createobject("ADODB.Recordset")
kull_ayar.Open "SELECT * FROM uyeler" ,BAG,1,3
%>
<center><br>
<a href="default.asp"><span class="normal_r2yazi">Yönetici Bölümü Anasayfa</span></a>
<br></center>
<form method="POST" action="default.asp?goster=kullanici_ayarlari&kullanici_ayarlari_kayit_onay=1" name="kullayarlar">
<table border="0" cellpadding="0" style="border-collapse: collapse" width="100%">
	<tr>
		<td colspan="2" align="center"><span class="blog_baslik">Kullanýcý Ayarlarý</span></td>
	</tr>
	<tr>
		<td colspan="2"><font color="#FF0000">*</font> Yazýlmasý zorunlu bölgeleri gösterir</td>
	</tr>
	<tr>
		<td width="148" style="border: 1px dotted #C0C0C0">
		<font color="#FF0000">*</font>Kullanýcý Adým</td>
		<td width="100%" style="border: 1px dotted #C0C0C0">
		<input type="text" name="kullanici_adi" size="20" value="<%=kull_ayar("kullanici_adi")%>" maxlength="20"> 
		</td>
	</tr>
	<tr>
		<td width="148" style="border: 1px dotted #C0C0C0">
		<font color="#FF0000">*</font>Þifrem</td>
		<td width="100%" style="border: 1px dotted #C0C0C0">
		<input type="text" name="sifre" size="20" value="<%=kull_ayar("sifre")%>" maxlength="30"> 
		</td>
	</tr>
	<tr>
		<td width="148" style="border: 1px dotted #C0C0C0">Gerçek Kimliðim</td>
		<td width="100%" style="border: 1px dotted #C0C0C0">
		<input type="text" name="ad" size="20" value="<%=kull_ayar("ad")%>" maxlength="30"> 
		</td>
	</tr>
	<tr>
		<td width="148" style="border: 1px dotted #C0C0C0">Mail adresim</td>
		<td width="100%" style="border: 1px dotted #C0C0C0">
		<input type="text" name="mail" size="20" value="<%=kull_ayar("mail")%>" maxlength="50">
		<font color="#FF0000">Gösterilsinmi:</font><input type="radio" name="mail_goster" value="1"<%if kull_ayar("mail_goster")=true then%> checked<%else%><%end if%>>Evet&nbsp;&nbsp;&nbsp;
		<input type="radio" name="mail_goster" value="0" <%if kull_ayar("mail_goster")=true then%><%else%> checked<%end if%>>Hayýr </td>
	</tr>
	<tr>
		<td width="148" style="border: 1px dotted #C0C0C0">Icq adresim</td>
		<td width="100%" style="border: 1px dotted #C0C0C0">
		<input type="text" name="icq" size="20" value="<%=kull_ayar("icq")%>" maxlength="50">
		<font color="#FF0000">Gösterilsinmi:</font><input type="radio" name="icq_goster" value="1"<%if kull_ayar("icq_goster")=true then%> checked<%else%><%end if%>>Evet&nbsp;&nbsp;&nbsp;
		<input type="radio" name="icq_goster" value="0" <%if kull_ayar("icq_goster")=true then%><%else%> checked<%end if%>>Hayýr 
		</td>
	</tr>
	<tr>
		<td width="148" style="border: 1px dotted #C0C0C0">Gmail/Google Talk adresim</td>
		<td width="100%" style="border: 1px dotted #C0C0C0">
		<input type="text" name="gtalk" size="20" value="<%=kull_ayar("gtalk")%>" maxlength="50">
		<font color="#FF0000">Gösterilsinmi:</font><input type="radio" name="gtalk_goster" value="1"<%if kull_ayar("gtalk_goster")=true then%> checked<%else%><%end if%>>Evet&nbsp;&nbsp;&nbsp;
		<input type="radio" name="gtalk_goster" value="0" <%if kull_ayar("gtalk_goster")=true then%><%else%> checked<%end if%>>Hayýr 
		</td>
	</tr>
	<tr>
		<td width="148" style="border: 1px dotted #C0C0C0">Kiþisel Resmim</td>
		<td width="100%" style="border: 1px dotted #C0C0C0">
		<input type="text" name="resim" size="20" value="<%=kull_ayar("resim")%>" maxlength="50">
		<font color="#FF0000">Resim yolunu yazýn (64x64) (Resim baþka adresteyse 
		baþýna http:// koyun)</font></td>
	</tr>
	<tr>
		<td width="148" style="border: 1px dotted #C0C0C0">Hakkýmda Kýsa Bilgi</td>
		<td width="100%" style="border: 1px dotted #C0C0C0">
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
<textarea rows="10" name="kisa_bilgi" cols="52"><%=kull_ayar("kisa_bilgi")%></textarea>
<script language="javascript1.2">
editor_generate('kisa_bilgi');
</script> 
<!--  -->

		</td>
	</tr>
	<tr>
		<td width="148">
		<p align="center">
		<input type="submit" value="Güncelle" name="submit" style="font-family: Verdana; font-size: 7pt; font-weight: 700; color: #000000; border: 1px solid #000000; float:right"></td>
</form>
		<td width="148"><% 
kullanici_ayarlari_kayit_onay=request("kullanici_ayarlari_kayit_onay")
if kullanici_ayarlari_kayit_onay=1 and 1 = 1 then

Set kull_ayar = Server.Createobject("ADODB.Recordset")
kull_ayar.Open "SELECT * FROM uyeler" ,BAG,1,3

'formdan gelen bilgileri yerlerine eþitleyelim
kull_ayar("kullanici_adi") 	= request.form("kullanici_adi")
kull_ayar("sifre") 			= request.form("sifre")
kull_ayar("ad")				= request.form("ad")
kull_ayar("mail")			= request.form("mail")
kull_ayar("mail_goster") 	= request.form("mail_goster")
kull_ayar("icq")			= request.form("icq")
kull_ayar("icq_goster") 	= request.form("icq_goster")
kull_ayar("gtalk")			= request.form("gtalk")
kull_ayar("gtalk_goster") 	= request.form("gtalk_goster")
kull_ayar("kisa_bilgi") 	= request.form("kisa_bilgi")
kull_ayar("resim") 			= request.form("resim")

'veritabanýmýzý upload edelim
kull_ayar.update

'sonrada serverýn hafýzasýndan yer kaplamamasý için silelim
kull_ayar.Close 
Set kull_ayar = Nothing
response.redirect "?goster=kullanici_ayarlari&islem=guncellendi"

'son olarakta ayarlar sayfamýza geri dönelim
if islem="guncellendi" then
Response.write("<font color='red'>&nbsp;Ayarlar Güncellendi</font>")

else

Response.write("<font color='red'>&nbsp;Ýþlemi tamamlamak için<br>&nbsp;güncelle butonuna basýn</font>")

end if
end if

%>
<%
end sub
%>
<%
'---------------------PROJELERÝM BÖLÜMÜ-----------------------
sub proje_ayarlari
%>
<center><br>
<a href="default.asp"><span class="normal_r2yazi">Yönetici Bölümü Anasayfa</span></a>
<br></center>

<%
Set proje = Server.Createobject("ADODB.Recordset")
proje.Open "SELECT * FROM projelerim" ,BAG,1,3
for i=1 to proje.recordcount
if not proje.eof then
%>	
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3" height="76">
		<tr>
			<td width="101" height="15"><span class="normal_ryazi">Proje Adý</span></td>
			<td width="12" align="left" height="15"><span class="normal_ryazi">:</span></td>
			<td width="868" height="15"><span class="normal_yazi"><%=proje("ad")%></span></td>
		</tr>
		<tr>
			<td width="101" height="15"><span class="normal_ryazi">Seviye</span></td>
			<td width="12" align="left" height="15"><span class="normal_ryazi">:</span></td>
			<td width="868" height="15"><span class="normal_yazi"><%=proje("seviye")%></span></td>
		</tr>
		<tr>
			<td width="101" height="15"><span class="normal_ryazi">Amacý</span></td>
			<td width="12" align="left" height="15"><span class="normal_ryazi">:</span></td>
			<td width="868" height="15"><span class="normal_yazi"><%=proje("amac")%></span></td>
		</tr>
		<tr>
			<td width="101" height="19"><span class="normal_ryazi">Proje Durumu</span></td>
			<td width="12" align="left" height="19"><span class="normal_ryazi">:</span></td>
			<td width="868" height="19"><span class="normal_yazi"><%=proje("durum")%></span></td>
		</tr>
		<tr>
			<td width="101" valign="top" height="19"><span class="normal_ryazi">KýsaAçýklama</span></td>
			<td width="12" align="left" valign="top" height="19"><span class="normal_ryazi">:</span></td>
			<td width="868" valign="top" height="19"><span class="normal_yazi"><%=proje("kisa_aciklama")%></span></td>
		</tr>
		<tr>
			<td width="101" valign="top" height="19"><span class="normal_ryazi">Açýklama</span></td>
			<td width="12" align="left" valign="top" height="19"><span class="normal_ryazi">:</span></td>
			<td width="868" valign="top" height="19"><span class="normal_yazi"><%=proje("aciklama")%></span></td>
		</tr>
		<tr>
			<td width="101" valign="top" height="19"><span class="normal_ryazi">Yönetici iþlevi</span></td>
			<td width="12" align="left" valign="top" height="19"><span class="normal_ryazi">:</span></td>
			<td width="868" valign="top" height="19"><a href="?goster=proje_ayarlari&islem=proje_sil&id=<%=proje("id")%>"><span class="normal_ryazi">Sil</span></a><span class="normal_yazi"> | </span><a href="proje_duzenle.asp?id=<%=proje("id")%>"><span class="normal_ryazi">Düzenle</span></a></td>
		</tr>
		<tr>
			<td width="981" valign="top" height="19" colspan="3"><hr style="border: 1px dotted #C0C0C0" width="97%" size="1"></td>
		</tr>
	</table>
<%else%>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3" height="76">
		<tr>
		<td><span class="normal_ryazi">Kayýtlý Proje bulunmuyor</span></td>
		</tr>
	</table>
<%
end if

proje.movenext
next

if islem="proje_eklendi" then%>
<script language="JavaScript">alert('Uyarý!\n\nProje sisteme eklendi\nTeþekkür ederiz.');</script>
<%end if
if islem="proje_duzenlendi" then%>
<script language="JavaScript">alert('Uyarý!\n\nProje düzenlendi\nTeþekkür ederiz.');</script>
<%end if%>

<form action="default.asp?goster=proje_ayarlari&islem=proje_ekle" method="POST">
<table border="0" cellpadding="0" cellspacing="0" id="table2">
<tr>
	<td width="722" colspan="2" height="1">
	<hr size="1"></td>
</tr>
<tr>
	<td width="722" colspan="2">
	<font color="#FF0000">
	<span style="font-family: verdana; font-weight: 700; font-size: 8pt">-Yeni bir proje ekle-</span></font></td>
</tr>
<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Proje Adý</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="ad" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="Kýsa bir ad girin"></td>
</tr>
<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Seviye</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="seviye" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42"></td>
</tr>
		<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Amacý</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="amac" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42"></td>
		</tr>
		<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Durum</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="durum" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42"></td>
		</tr>
<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Kýsa Açýklama</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="kisa_aciklama" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42"></td>
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
</textarea>
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
	<input type="submit" value="Ekle" style="cursor: hand; font-size: 7pt; font-family: Verdana; border-width: 1;">
	</form>
	<font color="#FF0000"><span style="cursor: hand; font-size: 7pt; font-family: Verdana; border-width: 1; border: 1;"><a href="javascript:window.close();">Kapat</a></span></font></td>
</tr>
</table>
<%if islem="proje_ekle" then
'-----------proje ekle----------------------
Set proje = Server.Createobject("ADODB.Recordset")
proje.Open "SELECT * FROM projelerim" ,BAG,1,3

proje.AddNew

proje("ad")		 = Trim(Request.Form("ad"))
proje("seviye")	 = Trim(Request.Form("seviye"))
proje("amac")	 = Trim(Request.Form("amac"))
proje("durum")	 = Trim(Request.Form("durum"))
proje("aciklama")= Trim(Request.Form("aciklama"))
proje("kisa_aciklama")= Trim(Request.Form("kisa_aciklama"))

proje.update
Set proje= Nothing
Response.Redirect "default.asp?goster=proje_ayarlari&islem=proje_eklendi"

elseif islem="proje_sil" then
'-------------proje sil--------------------
proje_id=request.querystring("id")

Set proje = Server.Createobject("ADODB.Recordset")
proje.Open "DELETE * FROM projelerim WHERE id="&proje_id&"" ,BAG,1,3

Response.Redirect "default.asp?goster=proje_ayarlari&islem=proje_silindi"

proje .close
set proje =nothing
'proje sil bitiþ--------------------

end if
end sub
'---------------------BAGLANTILAR BÖLÜMÜ-----------------------
sub baglantilar_bolumu

hata=request("hata")

if hata="3" then%>
<script language="JavaScript">alert('HATA!\n\nSistem devam edemiyor.\nLütfen girilmesi zorunlu bölümleri eksiksiz doldurunuz.');</script>
<%end if

if request("gorew")="ekle" then

set baglantilar = Server.CreateObject("ADODB.Recordset")
baglantilar_sor = "SELECT * FROM baglantilarim ORDER BY baglantilar_id ASC"
baglantilar.Open baglantilar_sor,bag,1,3

If request.form("ad")="" Then
Response.redirect("default.asp?goster=baglantilar&hata=3")
end if
if request.form("aciklama")="" Then
Response.redirect("default.asp?goster=baglantilar&hata=3")
end if
if request.form("link")="" Then
Response.redirect("default.asp?goster=baglantilar&hata=3")
end if
if request.form("resim")="" Then
Response.redirect("default.asp?goster=baglantilar&hata=3")
end if


baglantilar.addnew

baglantilar("baglantilar_ad") 				= request.form("ad")
baglantilar("baglantilar_aciklama") 		= request.form("aciklama")
baglantilar("baglantilar_link")				= request.form("link")
baglantilar("baglantilar_resim") 			= request.form("resim")

baglantilar.update

Response.redirect("default.asp?goster=baglantilar")


elseif request("gorew")="sil" then

set baglantilar = Server.CreateObject("ADODB.Recordset")
baglantilar_sor = "DELETE * FROM baglantilarim WHERE baglantilar_id="&request("id")
baglantilar.Open baglantilar_sor,bag,1,3

Response.redirect("default.asp?goster=baglantilar")

end if
%>
<center><br>
<a href="default.asp"><span class="normal_r2yazi">Yönetici Bölümü Anasayfa</span></a>
<br></center>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
			<tr>
			<td width="100%" height="15" colspan="2">
<table border="0" cellpadding="0" cellspacing="0" width="866" id="table2">
<tr>
	<td align="left" width="192">
	<p align="left"><b><font color="#FF0000">Resim</font></b></td>
	<td align="left" width="193"><b><font color="#FF0000">Ad</font></b></td>
	<td align="left" width="217"><b><font color="#FF0000">Link</font></b></td>
	<td align="left" width="210"><b><font color="#FF0000">Açýklama</font></b></td>
	<td align="left" width="54"><font color="#FF0000"><b>Seçim</b></font></td>
</tr>
</table>

<%
set baglantilar = Server.CreateObject("ADODB.Recordset")
baglantilar_sor = "SELECT * FROM baglantilarim ORDER BY baglantilar_id ASC"
baglantilar.Open baglantilar_sor,bag,1,3

if baglantilar.eof then
%>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
			<tr>
				<td width="10" height="15"></td>
				<td width="16" height="15"></td>
				<td width="2" height="15"></td>
				<td width="248" height="15">Kayýtlý baðlantý yok</span></td>
			</tr>
		</table>
<%
else

for b=1 to baglantilar.recordcount
if baglantilar.eof then exit for
%>
<table border="0" cellpadding="0" cellspacing="0" width="865" id="table2">
<tr>
	<td width="194"><%=baglantilar("baglantilar_resim")%></td><td width="191"><%=baglantilar("baglantilar_ad")%></td>
	<td width="217"><%=baglantilar("baglantilar_link")%></td><td width="212"><%=baglantilar("baglantilar_aciklama")%></td>
	<td width="51"><a href="default.asp?goster=baglantilar&gorew=sil&id=<%=baglantilar("baglantilar_id")%>">Sil</a></td>
</tr>
</table>
<%
baglantilar.movenext
next

end if
%>
			</td>
			</tr>
		</table>
<form method="POST" action="default.asp?goster=baglantilar&gorew=ekle" name="baglanti_ekle">
<table border="0" cellpadding="0" style="border-collapse: collapse" width="296">
	<tr>
		<td colspan="2" align="center">Yeni bir baðlantý ekle</td>
	</tr>
	<tr>
		<td width="148">Baðlantý adý</td>
		<td width="148">
		<input type="text" name="ad" size="20"></td>
	</tr>
	<tr>
		<td width="148">Açýklama</td>
		<td width="148">
		<input type="text" name="aciklama" size="20"></td>
	</tr>
	<tr>
		<td width="148">Baðlantý linki</td>
		<td width="148"><input type="text" name="link" size="20"></td>
	</tr>
	<tr>
		<td width="148">Baðlantý Küçük Resmi</td>
		<td width="148"><input type="text" name="resim" size="20"></td>
	</tr>
	<tr>
		<td width="148">
		<p align="center">
		<input type="submit" value="ekle" name="submit" style="font-family: Verdana; font-size: 7pt; font-weight: 700; color: #000000; border: 1px solid #000000; float:right"></td>
</form>
<%end sub'---------------------RESÝMLER BÖLÜMÜ-----------------------
sub resimler
%>
<%

Kok=Server.Mappath("\") & "\blog\upload\resimler"

Set Fs=Server.CreateObject("Scripting.FileSystemObject")
Set DizinO=Fs.GetFolder(Kok)
Set Dosyalar=DizinO.Files
%><center><br>
<a href="default.asp"><span class="normal_r2yazi">Yönetici Bölümü Anasayfa</span></a>
<br></center>

<table cellspacing="0" cellpadding="0">
<tr>
	<td width="32"><span class="normal_ryazi">Ýþlem</span></td>
	<td width="221" align="left"><span class="normal_ryazi">Dosya</span></td>
	<td width="12%"><span class="normal_ryazi">Boyut</span></td>
</tr>
<%
For Each BirDosya In Dosyalar
%>
	<tr>
		<td width="32">
		<p align="center">
		<a href="default.asp?goster=resimler&islem=resim_sil&dosya=<%=Server.UrlEncode(BirDosya.Name)%>"><img src="../<%=ana_ayarlar("resim_yolu")%>sil.gif" border="0" width="13" height="14"></a></a></td>
		<td align="left"><a href="../upload/resimler/<%=Server.UrlEncode(BirDosya.Name)%>" target="_blank"><span class="normal_ryazi"><%=BirDosya.Name%></span></a></td>
		<td align="right"><span class="normal_yazi"><%=BoyutFormati(BirDosya.Size)%></span></td>
	</tr>
<%
Next

Set Dosyalar=Nothing
Set DizinO=Nothing
Set Fs=Nothing

%>
<tr>
	<td colspan="3" height="1"><hr style="border: 1px dotted #C0C0C0" width="97%" size="1"></td>
</tr>
<tr>
	<td colspan="3"><span class="normal_ryazi">Yeni Resim Yükle</span></td>
</tr>
<form method="post" enctype="multipart/form-data" action="resimyukle.asp">
	<tr><td colspan="3">
		<input class="normal_yazi" type="file" name="dosya" size="30" class="normal_yazi">
		<input class="normal_yazi" type="submit" value="Yükle">
	</td></tr>
</form>

</table>


<%
end sub

if islem="resim_sil"then '-----------resim sil---------------

	Dosya=Request.QueryString("dosya")
	
	Kok=Server.Mappath("\") & "\blog\upload\resimler"
	SilinecekDosya=Kok & "\" & Dosya
	
	On Error Resume Next
	
	Set Fs=Server.CreateObject("Scripting.FileSystemObject")
	Fs.DeleteFile (SilinecekDosya)
	
	Set Fs=Nothing
	Response.Redirect "default.asp?goster=resimler"
	
	If Err.Number <> 0 Then
		Response.Redirect "default.asp?goster=resimler&hata=3"
	end if

elseif islem="resim_yukle"then '-----------resim yükle---------------
%>

<%
end if'----REsim yükle bitiþþþþþ--------

'-Boyut fonksiyonu :))
Function BoyutFormati(Boyut)
	If Boyut < 1024 Then
		BoyutFormati=Boyut & " B"
	ElseIf Boyut < 1024*1024 Then
		BoyutFormati=FormatNumber(Boyut/1024,2) & " KB"
	ElseIf Boyut < 1024*1024*1024 Then
		BoyutFormati=FormatNumber(Boyut/(1024*1024),2) & " MB"
	Else
		BoyutFormati=FormatNumber(Boyut/(1024*1024*1024),2) & " GB"
	End If
End Function

else%>
<%if islem="cikis_yapildi" then%><script language="JavaScript">alert('Uyarý!\n\nYönetici bölümünden güvenli çýkýþýnýz yapýldý\nTeþekkür ederiz');</script><%end if%>
<%if hata="1" then%><script language="JavaScript">alert('Hata!!\n\nKullanýcý adýný girmediniz\nLüften tekrar deneyin');</script>
<%elseif hata="2" then%><script language="JavaScript">alert('Hata!!\n\nÞifreyi girmediniz\nLüften tekrar deneyin');</script>
<%elseif hata="3" then%><script language="JavaScript">alert('Hata!!\n\nGiriþ onaylanmadý!\nBilgileri yanlýþ girmiþ olabilirsiniz, Lüften tekrar deneyin');</script><%end if%>
<table border="0" cellpadding="0" cellspacing="0" bgcolor="#E1E1E1" id="table1">
	<tr>
		<td>
<FORM name=loginform action="giris_kontrol.asp" method="post">
<table border="0" cellpadding="0" cellspacing="1" width="523" id="table2">
	<tr>
		<td height="2" width="521" align="right" bgcolor="#808080" colspan="5">
		<p align="left"><b>
		<font face="Tahoma" style="font-size: 8pt" color="#FFFFFF">&nbsp; Spidey 
		Blog Script v.1 - Yönetici Giriþi</font></b></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<font face="Tahoma" style="font-size: 8pt" color="#FF0000">kullanýcý adý</font></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
		<input class="normal_yazi" name="kullanici_adi" size="20" style="border: 1px solid #808080"></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<font face="Tahoma" style="font-size: 8pt" color="#FF0000">þifre
		</font></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
	<input class="normal_yazi" name="sifre" size="20" style="border: 1px solid #808080" type="password"></td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		</td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
<input type="submit" value="Giriþ yap!" name="submit0" style="font-family: Verdana; font-size: 7pt; font-weight: 700; color: #000000; border: 1px solid #000000;">		</td>
	</tr>
	<tr>
		<td height="13" width="121" align="right">
		<font face="Tahoma" style="font-size: 8pt">
		<a style="text-decoration: none" href="javascript:close();">pencereyi kapat</a></font></td>
		<td height="13" width="4" align="right"></td>
		<td height="13" width="2" bgcolor="#808080">&nbsp;</td>
		<td height="13" width="5" align="left"></td>
		<td height="13" width="378" align="left">
		<font face="Tahoma" style="font-size: 8pt">pencereyi kapatmak için 
		týklayýn</font></td>
	</tr>
	</table>
		</td>
	</tr>
</table>
</form>
<%end if%>