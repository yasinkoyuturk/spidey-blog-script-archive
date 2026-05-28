		<table border="0" cellpadding="0" cellspacing="0" width="150" id="table2">
			<tr>
				<td width="2" height="20"></td> 
				<td width="122">
				<table border="0" cellpadding="0" cellspacing="0" width="150" id="table3">
					<tr>
						<td width="17">
						<img src="<%=ana_ayarlar("resim_yolu")%>menu.gif" width="16" height="16"></td>
						<td><span class="normal_ryazi">Menü</span></td>
					</tr>
				</table>
				</td>
				<td width="2"></td>
			</tr>
			<tr>
			<td width="195" height="15" colspan="2">
			<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
			<tr>
				<td width="10" height="15"></td>
				<td width="16" height="15"><img src="<%=ana_ayarlar("resim_yolu")%>/anasayfa.gif" border=0></td>
				<td width="2" height="15"></td>
				<td width="96" height="15"><a href="default.asp" target="_self"><span class="normal_yazi" help="Anasayfa!">Eve dönüþ!</span></a></td>
			</tr>
			<tr>
				<td width="10" height="15"></td>
				<td width="16" height="15"><img src="<%=ana_ayarlar("resim_yolu")%>/proje.gif" border=0></td>
				<td width="2" height="15"></td>
				<td width="96" height="15"><a href="projelerim.asp" target="_self"><span class="normal_yazi" help="Projelerim!">Projelerim</span></a>
				<span class="normal_ryazi">
				<%
				Set proje = Server.Createobject("ADODB.Recordset")
				proje.Open "SELECT * FROM projelerim" ,BAG,1,3
				Response.write "("& proje.recordcount &")"
				proje.close
				set proje=nothing
				%>
				</span></td>
			</tr>
			<tr>
				<td width="10" height="15"></td>
				<td width="16" height="15"><img src="<%=ana_ayarlar("resim_yolu")%>/makale.gif" border=0></td>
				<td width="2" height="15"></td>
				<td width="96" height="15"><a href="makalelerim.asp" target="_self"><span class="normal_yazi" help="Birkaç püf noktasý!">Makalelerim</span></a></td>
			</tr>
			<tr>
				<td width="10" height="15"></td>
				<td width="16" height="15"><img src="<%=ana_ayarlar("resim_yolu")%>/ben.gif" border=0></td>
				<td width="2" height="15"></td>
				<td width="96" height="15"><a href="ben.asp" target="_self"><span class="normal_yazi" help="Kýsaca ben!">Ben</span></a></td>
			</tr>
			
			</table>
			</tr>
			<tr>
				<td height="5"></td>
			</tr>
			<tr>
				<td width="2" height="20"></td> 
				<td width="122">
				<table border="0" cellpadding="0" cellspacing="0" width="150" id="table4">
					<tr>
						<td width="23">
						<img src="<%=ana_ayarlar("resim_yolu")%>kategoriler.gif" width="22" height="20"></td>
						<td><span class="normal_ryazi">Kategoriler</span></td>
					</tr>
				</table>
				</td>
				<td width="2"></td>
			</tr>
			<tr>
			<td width="195" height="15" colspan="2">
<%
set kat = Server.CreateObject("ADODB.Recordset")
kat.Open "SELECT * FROM makale_kategoriler ORDER BY kate_id ASC" ,bag,1,3
kategori=kat.recordcount

if kat.eof then
%>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
			<tr>
				<td width="10" height="15"></td>
				<td width="16" height="15"><img src="<%=ana_ayarlar("resim_yolu")%>/resim_kate2.gif" border=0></td>
				<td width="2" height="15"></td>
				<td width="96" height="15"><span class="normal_yazi">Kategori yok</span></a></td>
			</tr>
		</table>
<%
else
	
	set alt = Server.CreateObject("ADODB.Recordset")
	sor = "SELECT * FROM makale_alt_kategoriler  WHERE Onayli=True"
	alt.Open sor,bag,1,3
	
	if alt.eof then
	yok="0"
	else
	yok=alt.recordcount
	end if
	%>		
<%
for i=1 to kat.recordcount
if kat.eof then exit for

kategori=kat("kate_id")

set alt_sayi = Server.CreateObject("ADODB.Recordset")
sor_sayi = "SELECT * FROM makale_alt_kategoriler WHERE onayli=true AND kate_id="& kategori
alt_sayi.Open sor_sayi,bag,1,3

%>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
			<tr>
				<td width="10" height="15"></td>
				<td width="16" height="15"><img src="<%=ana_ayarlar("resim_yolu")%>/resim_kate2.gif" border=0></td>
				<td width="2" height="15"></td>
				<td width="96" height="15"><a href="blog_kategori.asp?kategori=<%=kat("kate_id")%>" target="_self"><span class="normal_yazi" help="<%=kat("kate_aciklama")%>"><%=kat("kate_ad")%></span><span class="normal_ryazi"> (<%=alt_sayi.recordcount%>)</span></a></td>
			</tr>
		</table>
<%
kat.movenext
next
kat_sayi=kat.recordcount

end if%>
			</td>
			</tr>
			<tr>
				<td height="5"></td>
			</tr>
			<tr>
				<td width="2" height="20"></td> 
				<td width="122">
				<table border="0" cellpadding="0" cellspacing="0" width="150" id="table5">
					<tr>
						<td width="22">
						<img src="<%=ana_ayarlar("resim_yolu")%>baglantilar.gif" width="22" height="20"></td>
						<td><span class="normal_ryazi">Baðlantýlar</span></td>
					</tr>
				</table>
				</td>
				<td width="2"></td>
			</tr>
			<tr>
			<td width="195" height="15" colspan="2">
		
<%
set baglantilar = Server.CreateObject("ADODB.Recordset")
baglantilar_sor = "SELECT * FROM baglantilarim ORDER BY baglantilar_id ASC"
baglantilar.Open baglantilar_sor,bag,1,3

if baglantilar.eof then
%>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
			<tr>
				<td width="10" height="15"></td>
				<td width="16" height="15"><img src="<%=ana_ayarlar("resim_yolu")%>/resim_kate2.gif" border=0></td>
				<td width="2" height="15"></td>
				<td width="96" height="15"><span class="normal_yazi">Baðlantý yok</span></a></td>
			</tr>
		</table>

<%
else

for b=1 to baglantilar.recordcount
if baglantilar.eof then exit for
%>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
			<tr>
				<td width="10" height="15"></td>
				<td width="16" height="15"><img src="<%=baglantilar("baglantilar_resim")%>"></td>
				<td width="2" height="15"></td>
				<td width="96" height="15"><a href="<%=baglantilar("baglantilar_link")%>" target="_self"><span class="normal_yazi" help="<%=baglantilar("baglantilar_aciklama")%>"><%=baglantilar("baglantilar_ad")%></span></a></td>
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