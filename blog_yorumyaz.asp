<!-- #include file="inc/baglan.asp" -->
<%
function editor(strForm)

StrForm = replace(StrForm, "<", "& lt;")
StrForm = replace(StrForm, ">", "& qt;")

StrForm = Replace (StrForm ,"`","",1,-1,1)
StrForm = Replace (StrForm ,"=","",1,-1,1)
StrForm = Replace (StrForm ,"&","",1,-1,1)
StrForm = Replace (StrForm ,"%","",1,-1,1)
StrForm = Replace (StrForm ,"!","",1,-1,1)
StrForm = Replace (StrForm ,"#","",1,-1,1)
StrForm = Replace (StrForm ,"<","",1,-1,1)
StrForm = Replace (StrForm ,">","",1,-1,1)
StrForm = Replace (StrForm ,"*","",1,-1,1)
StrForm = Replace (StrForm ,"/","",1,-1,1)
StrForm = Replace (StrForm ,"\","",1,-1,1)
StrForm = Replace (StrForm ,"And","",1,-1,1)
StrForm = Replace (StrForm ,"'","",1,-1,1)
StrForm = Replace (StrForm ,"Chr(34)","",1,-1,1)
StrForm = Replace (StrForm ,"Chr(39)","",1,-1,1)

strForm = replace(strForm, "[:)]", "<img src=smileys/smiley1.gif>")
strForm = replace(strForm, "[:D]", "<img src=smileys/smiley4.gif>")
strForm = replace(strForm, "[:S]", "<img src=smileys/smiley5.gif>")
strForm = replace(strForm, "[(:(]", "<img src=smileys/smiley6.gif>")
strForm = replace(strForm, "[;)]", "<img src=smileys/smiley2.gif>")
strForm = replace(strForm, "[:o]", "<img src=smileys/smiley3.gif>")
strForm = replace(strForm, "[D)]", "<img src=smileys/smiley16.gif>")
strForm = replace(strForm, "[x(]", "<img src=smileys/smiley18.gif>")
strForm = replace(strForm, "[:@]", "<img src=smileys/smiley7.gif>")
strForm = replace(strForm, "[:o)]", "<img src=smileys/smiley8.gif>")
strForm = replace(strForm, "[;))]", "<img src=smileys/smiley9.gif>")
strForm = replace(strForm, "[:!]", "<img src=smileys/smiley10.gif>")
strForm = replace(strForm, "[:oo]", "<img src=smileys/smiley11.gif>")
strForm = replace(strForm, "[))]", "<img src=smileys/smiley12.gif>")
strForm = replace(strForm, "[:((]", "<img src=smileys/smiley13.gif>")
strForm = replace(strForm, "[:]]", "<img src=smileys/smiley14.gif>")
strForm = replace(strForm, "[):)]", "<img src=smileys/smiley15.gif>")
strForm = replace(strForm, "[:P]", "<img src=smileys/smiley17.gif>")

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

editor = strForm
end function
%>
<%
yazi = Trim(Request.Form("yazi"))
kategori = Trim(Request.Form("kategori"))

yazan = editor(Trim(Request.Form("yazan")))
makale_yorum = editor(Trim(Request.Form("makale_yorum")))
makale_mail = editor(Trim(Request.Form("makale_mail")))
makale_site = editor(Trim(Request.Form("makale_site")))
tarih = Trim(Request.Form("tarih"))
onay = False

Set my_yorum_ekle = Server.Createobject("ADODB.Recordset")
my_yorum_ekle.open "SELECT * FROM makale_yorumlar",BAG,1,3

my_yorum_ekle.AddNew

my_yorum_ekle("my_yazan") = yazan
my_yorum_ekle("my_yorum") = makale_yorum
my_yorum_ekle("my_tarih") = tarih
my_yorum_ekle("my_mail") = mail
my_yorum_ekle("my_site") = site
my_yorum_ekle("my_id") = yazi

my_yorum_ekle.update
Response.Redirect "blog_yorum.asp?onay_bekliyor=1&yazi="&yazi&"&kategori="&kategori
%>