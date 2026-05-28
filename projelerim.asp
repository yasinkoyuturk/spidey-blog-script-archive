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
			<td width="18"><img src="<%=ana_ayarlar("resim_yolu")%>proje.gif"></td>
			<td><span class="blog_baslik">Projelerim</span></td>
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
	<td valign="top" width="100%" rowspan="2">
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
			<td width="868" height="15"><a href="proje_goster.asp?pid=<%=proje("id")%>"><span class="normal_ryazi" help="Projeyi incelemek için týkla!"><%=proje("ad")%></span></a></td>
		</tr>
		<tr>
			<td width="101" height="15"><span class="normal_ryazi">Seviye</span></td>
			<td width="12" align="left" height="15"><span class="normal_ryazi">:</span></td>
			<td width="868" height="15"><span class="normal_yazi"><%=proje("seviye")%></span></td>
		</tr>
		<tr>
			<td width="101" height="15" valign="top"><span class="normal_ryazi">Amacý</span></td>
			<td width="12" align="left" height="15" valign="top"><span class="normal_ryazi">:</span></td>
			<td width="868" height="15"><span class="normal_yazi"><%=proje("amac")%></span></td>
		</tr>
		<tr>
			<td width="101" height="19"><span class="normal_ryazi">Proje Durumu</span></td>
			<td width="12" align="left" height="19"><span class="normal_ryazi">:</span></td>
			<td width="868" height="19"><span class="normal_yazi"><%=proje("durum")%></span></td>
		</tr>
		<tr>
			<td width="101" valign="top" height="19"><span class="normal_ryazi">Açýklama</span></td>
			<td width="12" align="left" valign="top" height="19"><span class="normal_ryazi">:</span></td>
			<td width="868" valign="top" height="19"><span class="normal_yazi"><%=proje("kisa_aciklama")%></span> <a href="proje_goster.asp?pid=<%=proje("id")%>"><span class="normal_ryazi" help="Projeyi incelemek için týkla!">[Projeyi incele!]</span></a></td>
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
%>
	</td>
</tr>
<tr>
	<td valign="top" width="1"></td>
	<td valign="top" width="4"></td>
</tr>
</table>
<!--#include file="inc/alt.asp" -->