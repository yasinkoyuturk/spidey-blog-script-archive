<!-- #include file="inc/baglan.asp" -->
<!--#include file="inc/ayarlar.asp" -->
<%
Set ana_ayarlar = Server.Createobject("ADODB.Recordset")
ana_ayarlar.Open "SELECT * FROM ana_ayarlar" ,BAG,1,3
%>

<% 
'Sql Injection 'ý Önlemek ve HTML Kodlarýný Pasif Hale Getirmek Ýçin Kullanýlan Karakter Temizleme Metodu 
function guvenlik(data) 
data = Replace (data ,"`","",1,-1,1) 
data = Replace (data ,"=","",1,-1,1) 
data = Replace (data ,"&","",1,-1,1) 
data = Replace (data ,"%","",1,-1,1) 
data = Replace (data ,"!","",1,-1,1) 
data = Replace (data ,"#","",1,-1,1) 
data = Replace (data ,"<","",1,-1,1) 
data = Replace (data ,">","",1,-1,1) 
data = Replace (data ,"*","",1,-1,1) 
data = Replace (data ,"And","",1,-1,1) 
data = Replace (data ,"'","",1,-1,1) 
data = Replace (data ,"Chr(34)","",1,-1,1) 
data = Replace (data ,"Chr(39)","",1,-1,1) 
data = Replace (data ,"select","",1,-1,1) 
data = Replace (data ,"join","",1,-1,1) 
data = Replace (data ,"union","",1,-1,1) 
data = Replace (data ,"where","",1,-1,1) 
data = Replace (data ,"insert","",1,-1,1) 
data = Replace (data ,"delete","",1,-1,1) 
data = Replace (data ,"update","",1,-1,1) 
data = Replace (data ,"like","",1,-1,1) 
data = Replace (data ,"drop","",1,-1,1) 
data = Replace (data ,"create","",1,-1,1) 
data = Replace (data ,"modify","",1,-1,1) 
data = Replace (data ,"rename","",1,-1,1) 
data = Replace (data ,"alter","",1,-1,1) 
data = Replace (data ,"cast","",1,-1,1) 
guvenlik=data 
end function 
%> 

<% 
If trim(guvenlik(request.form("kullanici_adi")))="" then 
Response.redirect "default.asp?hata=1"

ElseIf trim(guvenlik(request.form("sifre")))="" then 
Response.redirect "default.asp?hata=2"
end if 
%> 
<% 
Set rs = Server.CreateObject("Adodb.Recordset") 
Sorgu = "select id, kullanici_adi, sifre from uyeler where kullanici_adi = '" & guvenlik(request.form("kullanici_adi")) & "' and sifre = '" & guvenlik(Request.form ("sifre")) & "'" 
rs.Open Sorgu, BAG, 1, 3 
If rs.BOF And RS.EOF Then 
Response.redirect "default.asp?hata=3"
Else 
Session("blogadmin")=True 
Session("id") = rs("id") 
Session("kullanici_adi") = rs("kullanici_adi") 
Response.Cookies("TekTech")("kullanici_adi") = Request.Form("kullanici_adi") 
Response.Cookies("TekTech")("sifre") = Request.Form("sifre") 
Response.Cookies("TekTech").Expires = Now()+15 
Response.Redirect "default.asp" 
End If 
%>
