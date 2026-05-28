<!--#include file="inc/ust.asp" -->
<%
yazi=request.querystring("yazi")
kategori=request.querystring("kategori")

set alt = Server.CreateObject("ADODB.Recordset")
sor = "SELECT * FROM makale_alt_kategoriler WHERE onayli=true AND id="& yazi
alt.Open sor,bag,1,3

if alt.eof then Response.redirect "default.asp"

yorum_id=alt("id")

set my = Server.CreateObject("ADODB.Recordset")
mysor = "SELECT * FROM makale_yorumlar where my_onay=True AND my_id='"&yorum_id&"'"
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
'-----------------------yorum sil--------------------
yorum_sil=request.querystring("yorum_sil")

if yorum_sil=1 then

yorum=request.querystring("yorum")

set sil = Server.CreateObject("ADODB.Recordset")
silsor = "DELETE * FROM makale_yorumlar WHERE my_id2="& yorum
sil.Open silsor,bag,1,3

Response.Redirect "blog_yorum.asp?yazi="&yazi&"&kategori="&kategori

sil.close
set sil=nothing
end if
'-----------------------yorum sil bitiþ--------------------
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1">
	<tr>
		<td width="1"></td>
		<td><table><tr><td><img src="<%=ana_ayarlar("resim_yolu")%>blog_baslik.gif"></td><td><span class="blog_baslik"><%=alt("ad")%></span></td></tr></table></td>
	</tr>
	<tr>
		<td height="19" width="1"></td>
		<td height="19"><span class="normal_yazi"><i><%=alt("aciklama")%></i></span>
		<br><br>
		<span class="normal_yazi"><%=alt("makale_body")%></span><br><br>
		<span class="silik_yazi">Ekleyen: <font color=red><%=alt("ekleyen")%></font> | <%=alt("tarih")%>&nbsp;<%=alt("saat")%> | Kategori: </span><a href="blog_kategori.asp?kategori=<%=kat("kate_id")%>"><span class="normal_ryazi" help="<%=kat("kate_aciklama")%>"><%=kat("kate_ad")%></span></a><span class="silik_yazi"> | </span><a href="blog_yorum.asp?yazi=<%=alt("id")%>&kategori=<%=kat("kate_id")%>"><span class="silik_yazi">Yorum(lar) </span><span class="normal_ryazi"><%=toplam_yorum%></span><span class="silik_yazi"> / </span><span class="normal_ryazi">yorum yaz</span></a></td>
	</tr>
	<tr>
		<td height="1" width="1"></td>
		<td height="1"><hr style="border: 1px dotted #C0C0C0" width="97%" size="1"></td>
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
							<td weight="16" height="16" width="16"><img src="<%=ana_ayarlar("resim_yolu")%>yorumlar.gif"></td>
							<td weight="140" valign="middle" height="16"><p align="left">
							<span class="blog_baslik">Bu makaleye yapýlan yorumlar</span></p></td>
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

if my("my_onay")=True then

ad=my("my_yazan")


%>

						<table width="100%" cellspacing="0" cellpadding="0" id="table3">
							<tr>
								<td colspan="3">
								<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table6" height="38">
									<tr>
										<td>
										<span class="normal_ryazi"><%=my("my_yazan")%></span>
										<span class="normal_yazi">@ Tarih: <%=my("my_tarih")%></span><br>
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
end if
my.movenext
loop
%>			
			
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>

	</table>
	<span class="normal_ryazi" style="cursor: pointer;" onClick="javascript:if (document.getElementById('YorumYaz').style.display=='none' ) { document.getElementById('YorumYaz').style.display='block' } else { document.getElementById('YorumYaz').style.display='none' }"><center><table><tr><td><img src="<%=ana_ayarlar("resim_yolu")%>sende_yaz.gif" border="0"></td><td><span class=normal_ryazi> Sende yaz</span></td></tr></table></center></span>

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

<form name="YorumYaz" action="blog_yorumyaz.asp" method="post" name="minieditor2" id="minieditor2">
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
                        src="smileys/smiley1.gif" 
                        border=0 ></A>&nbsp;<A 
                        href="javascript:AddForm2('[:D]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley4.gif" 
                        border=0 ></A>&nbsp;<A 
                        href="javascript:AddForm2('[:S]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley5.gif" 
                        border=0 ></A>&nbsp;<A 
                        href="javascript:AddForm2('[(:(]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley6.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[;)]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley2.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:o]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley3.gif" 
                        border=0 ></A>&nbsp;<br>
					<A 
                        href="javascript:AddForm2('[D)]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley16.gif" 
                        border=0></A>&nbsp;<A 
						href="javascript:AddForm2('[x(]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley18.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:@]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley7.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:o)]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley8.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[;))]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley9.gif" 
                        border=0 ></A> 
<A 
                        href="javascript:AddForm2('[:oo]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley11.gif" 
                        border=0 ></A><br>
						&nbsp;<A 
                        href="javascript:AddForm2('[))]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley12.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:((]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley13.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[:]]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley14.gif" 
                        border=0 ></A> <A 
                        href="javascript:AddForm2('[):)]')"><IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley15.gif" 
                        border=0 ></A> <A
						href="javascript:AddForm2('[:P]')"> <IMG 
                        alt="Form 'a gülücük ekle" 
                        src="smileys/smiley17.gif" 
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
<!--#include file="inc/alt.asp" -->