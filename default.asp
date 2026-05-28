<!--#include file="inc/ust.asp" -->

<%
sf = Request.querystring("sf")
if sf="" then 
sf=1
end if

hata = Request.querystring("hata")
if hata="kelime_bulunamadi" then
%>
<script language="JavaScript">alert('Üzgünüm!\n\nArattýðýnýz kelime bulunamadý\nLütfen Tekrar Deneyin.');</script>
<%elseif hata="sayfa_yok" then%>
<script language="JavaScript">alert('Üzgünüm!\n\nMakaleleri bölümü yeni versiyonda (v2.0da) çýkacaktýr\nÝlginiz için teþekkür ederim\n\nSpidey Blog Script v1.5\n\nGüncellemeleri http://oyuncum.net/blog veya http://blog.turkbilgiforum.com adresinden takip edebilirsiniz');</script>
<%end if%>

<%
'------alt kategori bað--------
set alt = Server.CreateObject("ADODB.Recordset")
sor = "SELECT * FROM makale_alt_kategoriler where onayli=True order by id desc"
alt.Open sor,bag,1,3

if alt.eof then
%>
<center><br><br><br><br><br><br><br><br><br><br><br><hr style="border: 1px dotted #C0C0C0" width="97%" size="1"><span class="normal_ryazi">Blogda kayýtlý yazý bulunmuyor</span><br><hr style="border: 1px dotted #C0C0C0" width="97%" size="1"><br><br></center>
<%
else
yok=alt.recordcount


kategori=alt("kate_id")

'------kategori bað----
set kat = Server.CreateObject("ADODB.Recordset")
sor2 = "SELECT * FROM makale_kategoriler WHERE kate_id="& kategori
kat.Open sor2,bag,1,3

tablo_baslik=kat("kate_ad")

alt.pagesize = 10
alt.absolutepage = sf
sayfa = alt.pagecount
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1">

<%
for sayfala=1 to alt.pagesize
if alt.eof then exit for

yorum_id=alt("id")

set my = Server.CreateObject("ADODB.Recordset")
mysor = "SELECT * FROM makale_yorumlar where my_onay=True AND my_id='"&yorum_id&"'"
my.Open mysor,bag,1,3

toplam_yorum=my.recordcount
%>
<%if alt("onayli")=True then%>
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
else
end if
%>
<%
alt.movenext
next
%></table>
<table border="0" cellspacing="0" cellpadding="0" width="100%"><tr>
	<td valign="middle" width="214"><span class="normal_yazi">Toplam <font color=red><%=alt.pagecount%></font> sayfada, <font color=red><%=alt.recordcount%></font> yazý mevcut.<br>Her sayfada <font color=red>10</font> yazý gösteriliyor.</span></td>
	<td width="407" valign="middle">
	<p align="right"><span class="normal_yazi">.:Sayfa</span>
<%
for y=1 to sayfa 
if CINT(TRIM(sf))=CINT(TRIM(y)) then%>
<span class="normal_yazi"><%=y%></span></b>
<%else%>
<a href="default.asp?sf=<%=y%>&kategori=<%=kat("kate_id")%>"><span class="normal_ryazi"><%=y%></span></a>
<%end if%><%next
kat.close
set kat=nothing
alt.close
set alt=nothing

%>
</td></tr></table>
<%end if%><!--#include file="inc/alt.asp" -->