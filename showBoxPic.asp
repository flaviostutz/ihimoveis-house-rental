<%Response.buffer=true%>
<%Response.Expires=60%>
<html>
<!--#include file="myCommon.inc" --> 
<style TYPE="text/css">
<!--
	A:link {color:"#ffffff"; text-decoration: none}
	A:visited {color: "#ffffff"; text-decoration: none}
	A:hover {color:"gold"; text-decoration: underline}
-->
</style>
<head>
<TITLE>Imobiliária Higienópolis - www.ihimoveis.com.br</TITLE>
<meta http-equiv="Content-Type" content="text/html; chaRequest.QueryStringet=iso-8859-1">
</head>

<%
'comment & number - comment for the pic
pic=Request.QueryString("pic")
cod=Request.QueryString("cod")

'choose the next and previous pic
if fs.FileExists(realDir & "\imoveis\" & cod & "_" & (pic+1) & ".jpg") then
	nextPic=pic+1
else
	nextPic=1
end if

if fs.FileExists(realDir & "\imoveis\" & cod & "_" & (pic-1) & ".jpg") then
	prevPic=pic-1
else
	prevPic=0
	for i=0 to 7
		if (fs.FileExists(realDir & "\imoveis\" & cod & "_" & (8-i) & ".jpg")) and (prevPic=0) then
 			prevPic=(8-i)
		end if
	next
end if
%>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="imagens/bg_imagens.gif" link="#ffffff" vlink="#ffffff" alink="#ffffff">
<table border="0" cellspacing="0" cellpadding="0" height="100%" width="100%">
  <tr> 
    <td align="center" height="<%=Application("imgHeight")%>"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000066"><b><a href=# onclick="window.close()"><img src="imoveis\<%=cod%>_<%=pic%>.jpg" width="<%=Application("imgWidth")%>" height="<%=Application("imgHeight")%>" border="0" alt="Fechar"></a></b></font></div>
    </td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#48A1B8" height="20"> 
      <table border="0" cellspacing="0" cellpadding="4" width="100%" height="20">
        <tr> 
          <td align="center" valign="middle" width="12"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">&gt;</font></b></td>
          <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><%=Request.QueryString("comment"&pic)%></b></font></td>
          <td align="center" width="13"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="javascript:window.print()"><img src="imagens/bot_box_imprimir.gif" alt="clique para imprimir a foto" border="0"></a></font></b></td>
          <%'if nextPic<>prevPic then%>
          <td width="100" align="center"> <img src="imagens/bot_box_fotos.gif" border="0" usemap="#Map"> 
          </td>
          <%'end if%>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000066"><b><br>
</b></font> 
<map name="Map"> 
  <area shape="rect" coords="1,0,23,14" href="showBoxPic.asp?cod=<%=cod%>&pic=<%=prevPic%>&comment1=<%=Request.QueryString("comment1")%>&comment2=<%=Request.QueryString("comment2")%>&comment3=<%=Request.QueryString("comment3")%>&comment4=<%=Request.QueryString("comment4")%>&comment5=<%=Request.QueryString("comment5")%>&comment6=<%=Request.QueryString("comment6")%>&comment7=<%=Request.QueryString("comment7")%>&comment8=<%=Request.QueryString("comment8")%>" target="_self" alt="foto anterior" title="foto anterior">
  <area shape="rect" coords="32,1,89,15" href="showBoxPic.asp?cod=<%=cod%>&pic=<%=nextPic%>&comment1=<%=Request.QueryString("comment1")%>&comment2=<%=Request.QueryString("comment2")%>&comment3=<%=Request.QueryString("comment3")%>&comment4=<%=Request.QueryString("comment4")%>&comment5=<%=Request.QueryString("comment5")%>&comment6=<%=Request.QueryString("comment6")%>&comment7=<%=Request.QueryString("comment7")%>&comment8=<%=Request.QueryString("comment8")%>" alt="pr&oacute;xima foto" title="pr&oacute;xima foto">
</map>
</body>
</html>
<%response.flush%>