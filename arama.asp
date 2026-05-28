<!--#include file="inc/ust.asp" -->
<%
kelime=request.querystring("kelime")

Set alt = Server.CreateObject("Adodb.Recordset")
sql = "select * from makale_alt_kategoriler where ad like '%" & kelime & "%' OR aciklama like '%" & kelime & "%' order by okunma desc"
alt.open sql,bag,1,3

if alt.eof or kelime="" or kelime=" " then
Response.redirect "default.asp?hata=kelime_bulunamadi"
end if
%>
<%
kategori=alt("kate_id")

set kat = Server.CreateObject("ADODB.Recordset")
sor2 = "SELECT * FROM makale_kategoriler WHERE kate_id="& kategori
kat.Open sor2,bag,1,3
%>

<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1">
<!-- MSTableType="nolayout" -->
<tr>
	<td>
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5" height="5">
	<tr>
		<td width="18" valign="bottom"><img src="<%=ana_ayarlar("resim_yolu")%>ara.gif"></td>
		<td><span class="blog_baslik">Blogda Arama</span><span class="normal_yazi"> | Anahtar kelime: <span class="normal_ryazi"><%=kelime%><br>
		<%=alt.recordcount%></span> Sonuç bulundu - Bloglarýn <span class="normal_ryazi">baþlýklarýna</span> ve <span class="normal_ryazi">açýklamalarýna</span> göre aranýyor</span></td>
	</tr>
	<tr>
		<td colspan="2" height="5"></td>
	</tr>
</table>
	</td>
</tr>
<%do while not alt.eof
yorum_id=alt("id")

set my = Server.CreateObject("ADODB.Recordset")
mysor = "SELECT * FROM makale_yorumlar where my_onay=True AND my_id='"&yorum_id&"'"
my.Open mysor,bag,1,3

toplam_yorum=my.recordcount
%>
<%if alt("onayli")=True then%>
	<tr>
		<td>
		<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
		<td width="1"></td>
		<td><table><tr><td><img src="<%=ana_ayarlar("resim_yolu")%>blog_baslik.gif"></td><td><a href="blog_yorum.asp?yazi=<%=alt("id")%>&kategori=<%=kat("kate_id")%>"><span class="blog_baslik" help="Okumak için týkla!"><%=alt("ad")%></span></a></td></tr></table></td>
	</tr>
	<tr>
		<td height="19" width="1"></td>
		<td height="19"><span class="normal_yazi"><i><%=alt("aciklama")%></i></span><br><br>
		<span class="silik_yazi">Ekleyen: <span class="normal_ryazi"><%=alt("ekleyen")%></span> | <%=alt("tarih")%>&nbsp;<%=alt("saat")%> | Kategori: </span><a href="blog_kategori.asp?kategori=<%=kat("kate_id")%>"><span class="normal_ryazi" help="<%=kat("kate_aciklama")%>"><%=kat("kate_ad")%></span></a><span class="silik_yazi"> | </span><a href="blog_yorum.asp?yazi=<%=alt("id")%>&kategori=<%=kat("kate_id")%>"><span class="silik_yazi">Yorum(lar) </span><span class="normal_ryazi"><%=toplam_yorum%></span><span class="silik_yazi"> / </span><span class="normal_ryazi">yorum yaz</span></a></td>
	</tr>
	<tr>
		<td height="1" width="1"></td>
		<td height="1"><hr style="border: 1px dotted #C0C0C0" width="97%" size="1"></td>
	</tr>

		</table>
		</td>
	</tr>
<%
else
end if
%>
<%
alt.movenext
loop
%>
<tr><td><center><a href="javascript:history.back();"><span class="normal_ryazi">Geri</span></a></center></td></tr>
<tr><td height="5"></td></tr>
</table>
<%
alt.close
set alt=nothing
%><!--#include file="inc/alt.asp" -->