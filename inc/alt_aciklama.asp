<table border="0" cellpadding="0" cellspacing="0" width="690" id="table1">
	<tr>
		<td width="100" valign="bottom">
<%
if ana_ayarlar("g_sayac")=1 then
		
	Set sayac = Server.Createobject("ADODB.Recordset")
	sayac.Open "SELECT * FROM sayac" ,BAG,1,3

	if sayac.eof then
	%>
		<span class="normal_yazi" help="çoðul hit">0</span>
	<%
		sayac("cogul")=sayac("cogul")+1
		sayac.update
	else
		%>
		<span class="normal_yazi" help="çoðul hit"><%=sayac("cogul")%></span>
		<%
		sayac("cogul")=sayac("cogul")+1
		sayac.update

	end if			
		sayac.close
		set sayac=nothing
end if
		%>
		</td>
		<td width="490" valign="middle">
		<p align="center"><span class="site_alt_aciklamalar">Spidey Blog Script © v1.5 | Tüm haklarý 
		<a href="mailto:info[a]oyuncum.net?subject=Spidey blog Scrit hakkýnda!&body=Mesajýnýzý buraya yazýn :)" style="text-decoration: none">Yasin Koyutürk</a>'e aittir</span></td>
		<td width="100" valign="bottom">
		<p align="right"><%if kull_ayar("mail_goster")=true AND not kull_ayar("mail")="" then%><a href="mailto:<%=kull_ayar("mail")%>"><img src="<%=ana_ayarlar("resim_yolu")%>mail.gif" border="0" help="Mail/Msn<br><font color=red><%=kull_ayar("mail")%></font><br>Mail Göndermek için týkla!"></a><span class="normal_yazi"> </span><%end if%>
					<%if kull_ayar("icq_goster")=true AND not kull_ayar("icq")="" then%><a href="www.icq.com"><img src="<%=ana_ayarlar("resim_yolu")%>icq.gif" border="0" help="Icq adresim:<br><font color=red><%=kull_ayar("icq")%></font>"></a><span class="normal_yazi"> </span><%end if%>
					<%if kull_ayar("gtalk_goster")=true AND not kull_ayar("gtalk")="" then%><a href="www.google.com/talk"><img src="<%=ana_ayarlar("resim_yolu")%>gtalk.gif" border="0" help="Google Talk Adresim:<br><font color=red><%=kull_ayar("gtalk")%></font>"></a><span class="normal_yazi"> </span><%end if%> <a href="#top"><img src="<%=ana_ayarlar("resim_yolu")%>yukari.gif" border="0" help="Yukarý Dön!"></a></td>
	</tr>
</table>
