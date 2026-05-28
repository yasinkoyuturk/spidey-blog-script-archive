<!--#include file="inc/baglan.asp" -->
<!--#include file="inc/ayarlar.asp" -->

<%if Session("blogadmin")=True Then '---admin giriþ yapmýþsa---%>
<center><br>
<a href="default.asp"><span class="normal_r2yazi">Yönetici Bölümü Anasayfa</span></a><span class="normal_kyazi"> > </span>
<a href="blog_admin.asp"><span class="normal_r2yazi">Blog Admin Anasayfa</span></a><br>
<br></center>


<%
function guvenlik(strForm)

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

'küfür korumasý
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


yazi=request.querystring("yazi")
kategori=request.querystring("kategori")

set alt = Server.CreateObject("ADODB.Recordset")
sor = "SELECT * FROM makale_alt_kategoriler WHERE onayli=true AND id="& yazi
alt.Open sor,bag,1,3

if alt.eof then Response.redirect "default.asp"

yorum_id=alt("id")

set my = Server.CreateObject("ADODB.Recordset")
mysor = "SELECT * FROM makale_yorumlar where my_id='"&yorum_id&"'"
my.Open mysor,bag,1,3




set kat = Server.CreateObject("ADODB.Recordset")
sor2 = "SELECT * FROM makale_kategoriler WHERE kate_id="& kategori
kat.Open sor2,bag,1,3
%>

<%
onay_bekliyor=request("onay_bekliyor")
if onay_bekliyor=1 then%>
<script language="JavaScript">alert('Yorumunuz eklendi!\n\nYetkililer tarafýndan onaylandýktan sonra listelenmeye baþlanacaktýr\nTeþekkür ederiz!');</script>
<%end if%>
<%
'--------------makale düzenleme----------
gorev=request.querystring("gorev")
yazi=request.querystring("yazi")
kategori=request.querystring("kategori")

IF gorev = "blog_duzenle" THEN

ad = trim(Request.Form("ad"))
aciklama = trim(Request.Form("aciklama"))
makale_body = Trim(guvenlik(Request.Form("makale_body")))
kate_id = Trim(Request.Form("kate_id"))

Set makale = Server.Createobject("ADODB.Recordset")
msor = "SELECT * FROM makale_alt_kategoriler WHERE id="&yazi
makale.open msor,BAG,1,3

makale("ad") = ad
makale("makale_body") = makale_body
makale("aciklama") = aciklama
makale("kate_id") = kate_id

makale.update
Set makale = Nothing
response.Redirect "blog_admin_goster.asp?yazi="&yazi&"&kategori="&kategori
END IF

'-----------------------yorum sil--------------------
yorum_sil=request.querystring("yorum_sil")

if yorum_sil=1 then

yorum=request.querystring("yorum")

set sil = Server.CreateObject("ADODB.Recordset")
silsor = "DELETE * FROM makale_yorumlar WHERE my_id2="& yorum
sil.Open silsor,bag,1,3

Response.Redirect "blog_admin_goster.asp?yazi="&yazi&"&kategori="&kategori

sil.close
set sil=nothing
end if
'-----------------------yorum sil bitiþ--------------------
'-----------------------yorum onayla--------------------
yorum_onayla=request.querystring("yorum_onayla")

if yorum_onayla=1 then

yorum=request.querystring("yorum")

set onayla = Server.CreateObject("ADODB.Recordset")
onaylasor = "SELECT * FROM makale_yorumlar WHERE my_id2="& yorum
onayla.Open onaylasor,bag,1,3

onayla("my_onay")=True
onayla.update

Response.Redirect "blog_admin_goster.asp?yazi="&yazi&"&kategori="&kategori

onayla.close
set onayla=nothing
end if
'-----------------------yorum onayla bitiþ--------------------
IF gorev = "blog_duzenle" THEN

yazi = Request.querystring("yazi")

ad = Trim(Request.Form("ad"))
aciklama = Trim(Request.Form("aciklama"))
makale_body = Trim(Request.Form("makale_body"))
kate_id = Trim(Request.Form("kate_id"))

Set makale = Server.Createobject("ADODB.Recordset")
msor = "SELECT * FROM makale_alt_kategoriler WHERE id="&yazi
makale.open msor,BAG,1,3

makale("ad") = ad
makale("makale_body") = makale_body
makale("aciklama") = aciklama
makale("kate_id") = kate_id

makale.update
Set makale = Nothing
response.Redirect "blog_admin_goster.asp?yazi="&yazi&"&kategori="&kategori
END IF

%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1">

	<tr>
		<td><span class="blog_baslik"><%=alt("ad")%></span></td>
	</tr>
	<tr>
		<td height="19"><span class="normal_yazi"><i><%=alt("aciklama")%></i><br><br><%=alt("makale_body")%></span><br>
		<span class="silik_yazi">Ekleyen: <font color=red><%=alt("ekleyen")%></font> | <%=alt("tarih")%>&nbsp;<%=alt("saat")%></span></td>
	</tr>
	<tr>
		<td height="15">
		<hr style="border: 1px dotted #C0C0C0" width="97%" size="1"></td>
	</tr>
<%
kat.close
set kat=nothing
alt.close
set alt=nothing%>
</td></tr><tr>
	<td valign="middle" width="100%" colspan="2">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table7">
		<tr>
			<td height="15"></td>
		</tr>
		<tr>
			<td>
			<table border="0" width="100%" height="16" cellspacing="0" id="table8"><tr>
							<td weight="16" height="16" width="16"><img src="../<%=ana_ayarlar("resim_yolu")%>yorumlar.gif"></td>
							<td weight="140" valign="middle" height="16"><p align="left">
							<span class="makale_baslik">Bu makaleye yapýlan yorumlar</span></p></td>
							</tr>
							<tr>
							<td height="5"></td>
							</tr>
							</table></td>
		</tr>
		<tr>
			<td>
			
<%
if my.eof then Response.write("&nbsp;&nbsp;<span class=normal_yazi>Üzgünüz, bu makaleye hiç yorum yazýlmamýþ.Yorum yazmak istiyorsanýz '<font color=red>Sende Yaz</font>' butonuna týklayýn</span>")

do while not my.eof


ad=my("my_yazan")


%>

						<table width="100%" cellspacing="0" cellpadding="0" id="table3">
							<tr>
								<td colspan="3">
								<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table6" height="38">
									<tr>
										<td>
										<span class="normal_ryazi"><%=my("my_yazan")%></span>
										<span class="normal_yazi">@ Tarih: <%=my("my_tarih")%></span>
												<%if 1=1 then%>
												<a href="blog_admin_goster.asp?yorum_sil=1&yorum=<%=my("my_id2")%>&yazi=<%=yazi%>&kategori=<%=kategori%>"><span class=normal_ryazi>Yorumu sil</span></a> | <a href="blog_admin_goster.asp?yorum_onayla=1&yorum=<%=my("my_id2")%>&yazi=<%=yazi%>&kategori=<%=kategori%>"><span class=normal_ryazi>Yorumu Onayla</span><span class=normal_ryazi></a>   <%if my("my_onay")=False then%>Onay Bekliyor<%end if%></span>
												<%end if%>
										<br>
										<span class="normal_ryazi">+</span><span class="normal_yazi"><%=my("my_yorum")%></span>
										</td>
									</tr>
									<tr>
										<td height="5"><hr style="border: 1px dotted #C0C0C0" width="97%" size="1"></td>
									</tr>
								</table>
								</td>
							</tr>
						</table>

<%
my.movenext
loop
%>			
			
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>

	</table>
	<span class="normal_ryazi" style="cursor: pointer;" onClick="javascript:if (document.getElementById('YorumYaz').style.display=='none' ) { document.getElementById('YorumYaz').style.display='block' } else { document.getElementById('YorumYaz').style.display='none' }"><center><table><tr><td><img src="../<%=ana_ayarlar("resim_yolu")%>sende_yaz.gif" border="0"></td><td><span class=normal_ryazi> Sende yaz</span></td></tr></table></center></span>

<span ID="YorumYaz" style="border:1px solid #C0C0C0; display:none">
<script LANGUAGE="JavaScript">
<!-- Begin
function textCounter2(field, countfield, maxlimit) {
if (field.value.length > maxlimit) // if too long...trim it!
field.value = field.value.substring(0, maxlimit);
// otherwise, update 'characters left' counter
else
countfield.value = maxlimit - field.value.length;
}
// End -->
</script>




  <SCRIPT language=JavaScript>
		function AddForm2(form)
			{
				document.minieditor2.makale_yorum.value = document.minieditor2.makale_yorum.value + form
				document.minieditor2.makale_yorum.focus();
			}

  </SCRIPT>

<form name="YorumYaz" action="../blog_yorumyaz.asp" method="post" name="minieditor2" id="minieditor2">
	<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="72%" id="table5">
		<tr>
			<td>
			<p align="left"><span class="normal_yazi">* Yazýlmasý zorunlu alanlaýr gösterir!</td>
		</tr>
		<tr>
			<td>
			<div align="center">
			<table border="0" cellpadding="0" cellspacing="0" width="73%" id="table6">
				<tr>
					<td width="351" valign="bottom">
					<p align="right">

<span ID="YorumYaz1" style="display:none">
</span><input type="text" name="yazan" size="41" value="Adýnýz" style="border:1px solid #808080; font-family:Tahoma; font-size:8pt; color:#000000; background-color:#FFFFFF"><br><textarea name="makale_yorum" onKeyDown="textCounter(this.form.makale_yorum,this.form.remLen,300);" maxlength="300" onKeyUp="textCounter(this.form.makale_yorum,this.form.remLen,300);" style="border:1px solid #808080; width:223; height:106; font-family:Tahoma; font-size:8pt; color:#000000; background-color:#FFFFFF" rows="1" cols="20">Yorumuzda; ahlaksýzca davranýþlardan lütfen kaçýnýn
Teþekkür ederim</textarea><br>
<input type="text" name="site" size="18" value="Anasayfanýz" style="border:1px solid #808080; font-family:Tahoma; font-size:8pt; color:#000000; background-color:#FFFFFF"><input type="text" name="mail" size="19" value="Mail Adresiniz" style="border:1px solid #808080; font-family:Tahoma; font-size:8pt; color:#000000; background-color:#FFFFFF"><span ID="YorumYaz0" style="display:none"></td>
					<td width="270">
					&nbsp;
					<fieldset style="border:1px solid #C0C0C0; padding:1px; ; vertical-align:middle; width:144px; height:133px">
						<legend align="center"><span class="normal_yazi">Mini 
						Editör</span></legend>
						<p align="center"><A 
                        href="javascript:AddForm2('[:)]')"><IMG 
                        src="../smileys/smiley1.gif" 
                        border=0 ></A>&nbsp;<A 
                        href="javascript:AddForm2('[:D]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley4.gif" 
                        border=0 ></A>&nbsp;<A 
                        href="javascript:AddForm2('[:S]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley5.gif" 
                        border=0 ></A>&nbsp;<A 
                        href="javascript:AddForm2('[(:(]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley6.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[;)]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley2.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:o]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley3.gif" 
                        border=0 ></A>&nbsp;<br>
					<A 
                        href="javascript:AddForm2('[D)]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley16.gif" 
                        border=0></A>&nbsp;<A 
						href="javascript:AddForm2('[x(]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley18.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:@]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley7.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:o)]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley8.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[;))]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley9.gif" 
                        border=0 ></A> 
<A 
                        href="javascript:AddForm2('[:oo]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley11.gif" 
                        border=0 ></A><br>
						&nbsp;<A 
                        href="javascript:AddForm2('[))]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley12.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:((]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley13.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:]]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley14.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[):)]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley15.gif" 
                        border=0 ></A> <A
						href="javascript:AddForm2('[:P]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="../smileys/smiley17.gif" 
                        border=0 ></A><br>
</fieldset></td>
				</tr>
			</table>
			</div>
			</td>
		</tr>
		<tr>
			<td><center>
			&nbsp;<input type="hidden" name="yazi" value="<%=yazi%>"><input type="hidden" name="kategori" value="<%=kategori%>"><input type="hidden" name="tarih" value="<%=now()%>"><input class="textfield" value="300" size="2" name="remLen" onkeypress="return handleEnter(this, event)" disabled style="color: #000000; font-family: Verdana; font-size: 8pt; font-weight: bold; border: 1px solid #000000"> <input type="submit" value="Gönder" name="yolla" style="font-family: Tahoma; font-size: 8pt; color: #000000; border: 1px solid #000000; background-color: #FFFFFF"></center></td>
		</tr>
	</table>
	</div>
</form>
	</td>	
	</tr>
</table>
<%else
Response.Redirect "default.asp"
end if%>