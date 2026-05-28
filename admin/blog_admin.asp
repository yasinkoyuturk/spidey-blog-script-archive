<!--#include file="inc/baglan.asp" -->
<!--#include file="inc/ayarlar.asp" -->

<center><br>
<a href="default.asp"><span class="normal_r2yazi">Yönetici Bölümü Anasayfa</span></a>
<br></center>

<%if Session("blogadmin")=True Then '---admin giriþ yapmýþsa---
call admin_1

else
Response.Redirect "default.asp"
end if

function guvenlik(strForm)

'smiley için
strForm = replace(strForm, ":)", "<img src=smileys/smiley1.gif>")
strForm = replace(strForm, ";)", "<img src=smileys/smiley2.gif>")
strForm = replace(strForm, ":O", "<img src=smileys/smiley3.gif>")
strForm = replace(strForm, ":D", "<img src=smileys/smiley4.gif>")
strForm = replace(strForm, ":S", "<img src=smileys/smiley5.gif>")
strForm = replace(strForm, ":'(", "<img src=smileys/smiley6.gif>")
strForm = replace(strForm, ":@", "<img src=smileys/smiley7.gif>")
strForm = replace(strForm, ":o)", "<img src=smileys/smiley8.gif>")
strForm = replace(strForm, ";))", "<img src=smileys/smiley9.gif>")
strForm = replace(strForm, ":!", "<img src=smileys/smiley10.gif>")
strForm = replace(strForm, ":oo", "<img src=smileys/smiley11.gif>")
strForm = replace(strForm, ":((", "<img src=smileys/smiley13.gif>")
strForm = replace(strForm, ":P", "<img src=smileys/smiley17.gif>")
strForm = replace(strForm, "x(", "<img src=smileys/smiley18.gif>")

'küfür korumasý :)) kusura bakmayýn ama napim yazmalýyým aklýma gelenleri :)
strForm = replace(strForm, "siktir", "*")
strForm = replace(strForm, "salak", "*")
strForm = replace(strForm, "aptal", "*")
strForm = replace(strForm, "gerizekalý", "*")
strForm = replace(strForm, "öküz", "*")
strForm = replace(strForm, "o.c", "*")
strForm = replace(strForm, "a.q", "*")
strForm = replace(strForm, "amk", "*")
strForm = replace(strForm, "orospu", "*")
strForm = replace(strForm, "orospb", "*")
strForm = replace(strForm, "çocuklarý", "*")
strForm = replace(strForm, "göt", "*")
strForm = replace(strForm, "yarak", "*")
strForm = replace(strForm, "ipne", "*")
strForm = replace(strForm, "ibne", "*")
strForm = replace(strForm, "amcýk", "*")
strForm = replace(strForm, "aq", "*")
strForm = replace(strForm, "amýna", "*")
strForm = replace(strForm, "kodumun", "*")
strForm = replace(strForm, "siktimin", "*")
strForm = replace(strForm, "sikerim", "*")
strForm = replace(strForm, "sikim", "*")
strForm = replace(strForm, "sikerim", "*")
strForm = replace(strForm, "ananýzý", "*")

guvenlik = strForm
end function


%>


<%
'----------------------admin---------------------------
sub admin_1
%>
<title>Yönetici Bölümü - Yeni yazý yaz</title>
<%
kate_ekle = Request.QueryString("kate_ekle")
ekle = Request.QueryString("Ekle")
gorev = Request.QueryString("gorev")
guncellendi = Request.QueryString("guncellendi")
islem=request.querystring("islem")
%>
<%
set kat = Server.CreateObject("ADODB.Recordset")
sor2 = "SELECT * FROM makale_kategoriler"
kat.Open sor2,bag,1,3

if kat.eof then 
	Response.write("<span class=""normal_ryazi""> Kayýtlý kategori bulunmuyor..</span><br><br><span class=""normal_yazi"">Yeni bir yazý yazmak için, öncelikle bir kategori eklemelisiniz</span>")
	%>
	<form action="blog_admin.asp?gorev=kate_ekle" method="POST">
<span style="font-family: verdana; font-weight: 700; font-size: 8pt">-Yeni blog kategorisi Ekle-</span><br>
<input name="ad" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="30" value="Kategori Adý"><br>
<input name="aciklama" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="30" value="Açýklamasý"><br>
<input type="submit" value="Kategoriyi Ekle" style="cursor: hand; font-size: 7pt; font-family: Verdana; border-width: 1;">
</form>

	<%
else

	for i=1 to kat.recordcount
	
	kategori=kat("kate_id")
	
	set alt = Server.CreateObject("ADODB.Recordset")
	sor = "SELECT * FROM makale_alt_kategoriler WHERE kate_id="& kategori
	alt.Open sor,bag,1,3
	
	
	%><table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1">
		<tr height="20">
			<td width="17" bgcolor="#C0C0C0"></td>
			<td width="100%" bgcolor="#C0C0C0"><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;" title="<%=kat("kate_aciklama")%>"><%=kat("kate_ad")%></font> - <a href="?gorev=kate_sil&id=<%=kat("kate_id")%>">Sil</a></td>
		</tr>
		<tr height="20">
			<td height="16" colspan="3">
	<%
	if not alt.eof then 
		
		for ii=1 to alt.recordcount
			
			yorum_id=alt("id")
			
			set my = Server.CreateObject("ADODB.Recordset")
			mysor = "SELECT * FROM makale_yorumlar where my_onay=False AND my_id='"&yorum_id&"'"
			my.Open mysor,bag,1,3
			
			if my.recordcount=0 then
				onaysiz_yorumlar="<font color=blue>bulunmuyor</font>"
			else
				onaysiz_yorumlar="(<font color=red>"&my.recordcount&"</font> <a href=blog_admin_goster.asp?hitti=0&yazi="&alt("id")&"&kategori="&kat("kate_id")&"><font color=blue>Bak</font></a>)"
			end if
	%>
			<table width="758" border="0" cellpadding="0" cellspacing="0" id="table2">
				<tr>
					<td height="16" width="20"></td>
					<td width="276"><a href="blog_admin_goster.asp?hitti=0&yazi=<%=alt("id")%>&kategori=<%=kat("kate_id")%>"><font color="black" style="text-decoration: none;  font-family: verdana; font-size: 8pt;" title="Okumak için týklayýn"><%=trim(alt("ad"))%></font></a></td>
					<td width="54"><a href="blog_admin_duzenle.asp?yazi=<%=alt("id")%>&kategori=<%=kat("kate_id")%>"><font color="blue" style="font-family: verdana; font-size: 8pt;">Düzenle</font></a></td>
					<td width="38"><a href="?gorev=makale_sil&id=<%=alt("id")%>"><font color="#00FF00">Sil</font></a></td>
					<td width="400"><%if alt("onayli")=False then%><a href="?gorev=onay_ver&id=<%=alt("id")%>"><font color="#FF0000">Onay ver</font></a><%else%><font color="#0000FF">Onaylý... <a href="?gorev=onay_sil&id=<%=alt("id")%>">Onayý geri al</a></font><%end if%> 
					<font size="2">Onaysýz Yorumlar</font> <%=onaysiz_yorumlar%></td>
				</tr>
			</table>
	<%
		alt.movenext
		next
	
	else
		Response.write(" Kategori boþ..")
	end if
	%></td>
		</tr>
		</table><%kat.movenext
			next	
			kat.close
			set kat=nothing

%>
	<table>
	
	
		<tr> 
	    <td width="616" colspan="2">
<%
if ekle = 1 then
	Response.write("<center><br><br>Öðe Eklendi<br><a href='javascript:window.close();'>Kapat</a> / <a href='blog_admin.asp'>Geri</a></center>")
else
%>
<form action="blog_admin.asp?gorev=makale_ekle&ekle=1" method="POST" name="minieditor">
<tr>
	<td width="722" colspan="2" height="1">
	<hr size="1"></td>
</tr>
<tr>
	<td width="722" colspan="2">
	<font color="#FF0000">
	<span style="font-family: verdana; font-weight: 700; font-size: 8pt">-Yeni 
	bir yazý yaz-</span></font></td>
</tr>
<tr>
	<td width="82">
	<font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">Baþlýk:</font>
	</td>
	<td width="640">
	<input type="text" name="ad" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="Kýsa bir ad girin"></td>
</tr>
<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Açýklama</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
	</td>
	<td width="640">
	<input type="text" name="aciklama" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="42" value="Açýklama.."></td>
</tr>
<tr>
	<td width="82">
	<font color="#FF0000"><span style="font-family: verdana; font-size: 8pt">
	Yazýnýzýn içeriði</span></font><font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">:</font>
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
<textarea name="makale_body" style="width:670; height:300" rows="1" cols="20">smileylerinizi eklerken  direkt olarak <font color=red>:)          :(            x)           :'(   </font>þeklinde girmelisiniz :)</textarea><script language="javascript1.2">
editor_generate('makale_body');
</script> 
<!--  -->
	</td>
	<input type="hidden" name="ekleyen" value="<%=kull_ayar("kullanici_adi")%>">
</tr>
<tr>
	<td width="82">
	<font color="red" style="text-decoration: none;  font-family: verdana; font-size: 8pt;">Kategori</font>
	</td>
	<td width="640">
	<select size="1" name="kate_id" style="font-family: Verdana; font-size: 7pt; border: 1px solid #000000">
<%set kat1 = Server.CreateObject("ADODB.Recordset")
sor2 = "SELECT * FROM makale_kategoriler"
kat1.Open sor2,bag,1,3

for i=1 to kat1.recordcount%>
	<option value="<%=kat1("kate_id")%>"><%=kat1("kate_ad")%></option>
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
	<input type="submit" value="Ekle" style="cursor: hand; font-size: 7pt; font-family: Verdana; border-width: 1;"></form><font color="#FF0000"><span style="cursor: hand; font-size: 7pt; font-family: Verdana; border-width: 1; border: 1;"><a href="javascript:window.close();">Kapat</a></span></font></td>
</tr>
</font>
<tr><td></td><td>
<center>
<%end if%>
<form action="blog_admin.asp?gorev=kate_ekle" method="POST">
<span style="font-family: verdana; font-weight: 700; font-size: 8pt">-Yeni blog kategorisi Ekle-</span><br>
<input name="ad" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="30" value="Kategori Adý"><br>
<input name="aciklama" style="border:1px solid #000000; font-size: 7pt;  font-family:Verdana; color:#000000" size="30" value="Açýklamasý"><br>
<input type="submit" value="Kategoriyi Ekle" style="cursor: hand; font-size: 7pt; font-family: Verdana; border-width: 1;">
</form>
<center>
</td></tr>

<%
end if
%></table>





<%
'-----------makale ekle----------------------
IF gorev = "makale_ekle" THEN

ad = trim(Request.Form("ad"))
aciklama = trim(Request.Form("aciklama"))
makale_body = trim(guvenlik(Request.Form("makale_body")))
ekleyen = Trim(Request.Form("ekleyen"))
kate_id = Trim(Request.Form("kate_id"))

makale_body = Replace(makale_body,"[+bol]","<!--bol-->")

Set makale = Server.Createobject("ADODB.Recordset")
makale.open "SELECT * FROM makale_alt_kategoriler",BAG,1,3

makale.AddNew

makale("ad") = ad
makale("makale_body") = makale_body
makale("aciklama") = aciklama
makale("ekleyen") = ekleyen
makale("tarih") = Trim(date())
makale("saat") = Trim(time())
makale("okunma") = 0
makale("onayli") = False
makale("kate_id") = kate_id

makale.update
Set makale = Nothing
Response.Redirect "blog_admin.asp"
END IF
%><%
'------------kategori ekle-------------
IF gorev = "kate_ekle" THEN

ad = Trim(Request.Form("ad"))
aciklama = Trim(Request.Form("aciklama"))

Set makale = Server.Createobject("ADODB.Recordset")
makale.open "SELECT * FROM makale_kategoriler",BAG,1,3

makale.AddNew

makale("kate_ad") = ad
makale("kate_aciklama") = aciklama

makale.update
Set makale = Nothing
Response.Redirect "blog_admin.asp"
END IF
%>
		
<%
'-----------------------makale sil--------------------
if gorev = "makale_sil" then

id=request.querystring("id")

set alt = Server.CreateObject("ADODB.Recordset")
sor = "DELETE * FROM makale_alt_kategoriler WHERE id="& id
alt.Open sor,bag,1,3

Response.redirect("blog_admin.asp")

alt.close
set alt=nothing
end if
'-----------------------makale sil bitiþ--------------------
%>

<%
'-----------------------kategori sil--------------------
IF gorev = "kate_sil" THEN

id = Request("id")

set kat1 = Server.CreateObject("ADODB.Recordset")
sor2 = "DELETE * FROM makale_kategoriler where kate_id="& id
kat1.Open sor2,bag,1,3


set alt = Server.CreateObject("ADODB.Recordset")
sor = "DELETE * FROM makale_alt_kategoriler where kate_id="& id
alt.Open sor,bag,1,3

Response.Redirect "blog_admin.asp"

alt.close
set alt=nothing
kat.close
set kat=nothing
END IF
'-----------------------kategori sil bitiþ--------------------
%>
<%
'-----------------------ONAY VER--------------------
IF gorev = "onay_ver" THEN

id = Request.querystring("id")

set alt = Server.CreateObject("ADODB.Recordset")
sor = "SELECT onayli FROM makale_alt_kategoriler where id="& id
alt.Open sor,bag,1,3

alt("onayli") = True 
alt.update


Response.Redirect "blog_admin.asp"

alt.close
set alt=nothing

END IF
'-----------------------ONAY VER bitiþ--------------------

'-----------------------ONAY SiiiL--------------------
IF gorev = "onay_sil" THEN

id = Request.querystring("id")

set alt = Server.CreateObject("ADODB.Recordset")
sor = "SELECT onayli FROM makale_alt_kategoriler where id="& id
alt.Open sor,bag,1,3

alt("onayli") = False 
alt.update


Response.Redirect "blog_admin.asp"

alt.close
set alt=nothing

END IF
'-----------------------ONAY SiL bitiþ--------------------
%>

<%
end sub
%>