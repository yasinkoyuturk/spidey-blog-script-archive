<%
dim bag
Set BAG = Server.Createobject("ADODB.Connection")
BAG.Open ("DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("admin/db/veritabani.mdb"))
%>