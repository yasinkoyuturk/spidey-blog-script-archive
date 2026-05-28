<!--#include file="inc/ust.asp" -->
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" height="370">
<tr>
	<td valign="top" height="1" width="1"></td>
	<td valign="top" height="1" colspan="2"></td>
</tr>
<tr>
	<td valign="top" width="1" height="15"></td>
	<td valign="top" height="15" width="4">
	</td>
	<td valign="top" height="15" width="100%">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
		<tr>
			<td width="18"><img src="<%=ana_ayarlar("resim_yolu")%>ben.gif"></td>
			<td><span class="blog_baslik">Kýsaca Ben</span></td>
		</tr>
		<tr>
			<td height="5" colspan="2"></td>
		</tr>
	</table>
	</td>
</tr>
<tr>
	<td valign="top" width="1" height="96"></td>
	<td valign="top" height="64" width="4">
	&nbsp;</td>
	<td valign="top" height="64" width="100%">
	<table border="0" cellpadding="0" cellspacing="0" id="table5" height="64" width="100%">
		<tr>
			<td width="64"><%if not kull_ayar("resim")="" then%><img src="<%=kull_ayar("resim")%>" border="0" height="64" width="64"><%else%><img src="<%=ana_ayarlar("resim_yolu")%>resim_yok.gif" border="0" height="64" width="64"><%end if%></td>
			<td width="2"></td>
			<td valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%" height="75" id="table6">
				<tr>
					<td height="20" valign="bottom">
			<span class="blog_baslik"><%=kull_ayar("kullanici_adi")%></span><br>
			<%if not kull_ayar("ad")="" then%><span class="normal_yazi"><%=kull_ayar("ad")%></span><%else%><span class="normal_yazi">Ýsim girilmemiþ</span><%end if%>
					</td>
				</tr>
				<tr>
					<td height="5"></td>
				</tr>
				<tr>
					<td height="21">
					<%if kull_ayar("mail_goster")=true AND not kull_ayar("mail")="" then%><a href="mailto:<%=kull_ayar("mail")%>"><img src="<%=ana_ayarlar("resim_yolu")%>mail.gif" border="0" help="Mail/Msn<br><font color=red><%=kull_ayar("mail")%></font><br>Mail Göndermek için týkla!"></a><span class="normal_yazi"> </span><%end if%>
					<%if kull_ayar("icq_goster")=true AND not kull_ayar("icq")="" then%><a href="http://www.icq.com" target="_blank"><img src="<%=ana_ayarlar("resim_yolu")%>icq.gif" border="0" help="Icq adresim:<br><font color=red><%=kull_ayar("icq")%></font>"></a><span class="normal_yazi"> </span><%end if%>
					<%if kull_ayar("gtalk_goster")=true AND not kull_ayar("gtalk")="" then%><a href="http://www.google.com/talk" target="_blank"><img src="<%=ana_ayarlar("resim_yolu")%>gtalk.gif" border="0" help="Google Talk Adresim:<br><font color=red><%=kull_ayar("gtalk")%></font>"></a><span class="normal_yazi"> </span><%end if%>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td colspan="3">
			<hr style="border: 1px dotted #C0C0C0" width="97%" size="1"></td>
		</tr>
	</table>
	</td>
</tr>
<tr>
	<td valign="top" width="1"></td>
	<td valign="top" width="4"></td>
	<td valign="top" width="100%"><span class="normal_ryazi">Hakkýmda Kýsa Bilgi:</span><br><%if not kull_ayar("kisa_bilgi")="" then%><span class="normal_yazi"><%=kull_ayar("kisa_bilgi")%></span><%else%><span class="normal_yazi">Girilmemiþ</span><%end if%></td>
</tr>
</table>
<!--#include file="inc/alt.asp" -->