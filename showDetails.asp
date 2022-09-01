<%Response.buffer=true%>
<%Response.Expires=60%>
<html>
<!--#include file="common.inc" --> 
<!--#include file="myCommon.inc" --> 
<style TYPE="text/css">
<!--
	A:link {color:"#006666"; text-decoration: none}
	A:visited {color: "#006666"; text-decoration: none}
	A:hover {color:"#009966"; text-decoration: underline}
-->
</style>
<%
On error resume next
cod = Request.QueryString("cod")

set myConn = getMdb(Application("mdbFile"))
sqlStr = "select * from imoveis where cod='" & cod & "';"
set RS = myConn.Execute(sqlStr)

%>
<head>

</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#006666" vlink="#006666" alink="#009966">
<%if cod <> "" then%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" height="100%">
  <tr valign="middle" align="center"> 
    <td> 
      <table width="98%" border="0" cellspacing="0" cellpadding="2" bgcolor="#f4f4f4" align="right">
        <tr> 
            <td width="75%" valign="middle" align="right"> 
              <table border="0" width="98%" cellpadding="0" cellspacing="2" height="100%">
                <tr> 
                  <td colspan="3" height="17"><font face="Arial, Helvetica, sans-serif" size="1" color="#000000"><%=iTypeStr(RS("itype"))%> 
                    - Código <%=RS("cod")%></font></td>
                </tr>
                <tr> 
                  <td colspan="3"><font size="2" color="#000000" face="Verdana, Arial, Helvetica, sans-serif"><b><%=RS("building")%> 
                    <%if RS("building")<>"" then%>
                    - </b> 
                    <%end if%>
                    <%=RS("quarter")%> 
                    <%if RS("building")<>"" then write "</b>"%>
                    </font></td>
                </tr>
                <tr> 
                  <td colspan="3" height="17"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">Endereço:</font></b> 
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("address")%></font><font size="2"> 
                    - </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("city")%></font></font></td>
                </tr>
                <tr> 
                  <%if RS("rooms") <> "" then%>
                  <td colspan="3"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">N° 
                    de quartos:</font></b> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=RS("rooms")%></font></font> 
                  </td>
                  <%end if%>
                </tr>
                <%if RS("lote") <> "" or RS("area") <> "" then%>
                <tr> 
                  <%if RS("area") <> "" then%>
                  <td width="238"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">&Aacute;rea 
                    total:</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                    </font></b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"><%=RS("area")%></font></font> 
                    m&sup2;</td>
                  <%end if%>
                  <%if RS("lote") <> "" then%>
                  <td width="521"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">&Aacute;rea 
                    lote:</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                    </font></b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"><%=RS("lote")%></font></font> 
                    m&sup2;</td>
                  <%end if%>
                </tr>
                <%end if%>
                <tr> 
                  <td colspan="3" height="19"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">Valor:</font></b> 
                    <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%=FormatCurrency(RS("price"))%></font></font></td>
                </tr>
                <%if RS("cond")<>"" then%>
                <tr> 
                  <td colspan="3"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">Condom&iacute;nio:</font></b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"> 
                    <%=FormatCurrency(RS("cond"))%></font></font></td>
                </tr>
                <%end if%>
                <%if RS("IPTU")<>"" then%>
                <tr> 
                  <td colspan="3"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">IPTU:</font></b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"> 
                    <%
			if RS("IPTU")>=0 then
				write FormatCurrency(RS("IPTU"))
			else
            	write "Isento"
			end if
			%>
                    </font></font></td>
                </tr>
                <%end if
		descStr=RS("description")
		%>
                <%if descStr <> "" then%>
                <tr> 
                  <td colspan="3" valign="top"> 
                    
                  <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Detalhes:</b></font> 
                    <br>
                      <font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"><%=descStr%></font> 
                    </p>
                  </td>
                </tr>
                <%end if%>
              </table>
            </td>
            <td valign="middle" align="center" rowspan="2"> 
              <%if fs.FileExists(realDir & "\imoveis\" & cod & "_1.jpg") then%>
              <table border="0" cellspacing="6" cellpadding="0" height="100%" align="center">
                <%for i=0 to 3%>
                <%if fs.FileExists(realDir & "\imoveis\" & cod & "_" & (i*2)+1 & ".jpg") then%>
                <tr> 
                  <td width="50%" valign="middle" height="60" align="center"><a href="javascript:pic=openBox('showBoxPic.asp?cod=<%=cod%>&pic=<%=(i*2)+1%>&comment1=<%=RS("comment1")%>&comment2=<%=RS("comment2")%>&comment3=<%=RS("comment3")%>&comment4=<%=RS("comment4")%>&comment5=<%=RS("comment5")%>&comment6=<%=RS("comment6")%>&comment7=<%=RS("comment7")%>&comment8=<%=RS("comment8")%>','pic',<%=Application("imgWidth")%>,<%=Application("imgHeight")+24%>,false,1,90);pic.focus();"><img src=<%="imoveis/" & cod & "_" & (i*2)+1 & "p.jpg"%> border="1" width="80" height="60" alt="Clique nas fotos para ampliá-las"></a></td>
                  <td width="50%" valign="middle" height="60" align="center"><%if fs.FileExists(realDir & "\imoveis\" & cod & "_" & (i*2)+2 & ".jpg") then%><a href="javascript:pic=openBox('showBoxPic.asp?cod=<%=cod%>&pic=<%=(i*2)+2%>&comment1=<%=RS("comment1")%>&comment2=<%=RS("comment2")%>&comment3=<%=RS("comment3")%>&comment4=<%=RS("comment4")%>&comment5=<%=RS("comment5")%>&comment6=<%=RS("comment6")%>&comment7=<%=RS("comment7")%>&comment8=<%=RS("comment8")%>','pic',<%=Application("imgWidth")%>,<%=Application("imgHeight")+24%>,false,1,90);pic.focus();"><img src=<%="imoveis/" & cod & "_" & (i*2)+2 & "p.jpg"%> border="1" width="80" height="60" alt="Clique nas fotos para ampliá-las"></a><%else%>&nbsp; <%end if%></td>
                </tr>
                <%end if%>
                <%next%>
              </table>
              <%else%>
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><b> 
              N&atilde;o h&aacute; nenhuma foto disponível.</b></font> 
              <%end if%>
          </tr>
          <tr> 
            
          <td valign="bottom" align="right"> 
            <table width="98%" border="0" cellspacing="0" cellpadding="5">
              <tr valign="middle"> 
                <td width="19%"><font color="#0000FF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="javascript:history.go(-1)"><img src="imagens/bot_voltar.gif" border="0" alt="voltar"></a></b></font></td>
                <td align="center" width="30%"><a href="javascript:window.print()"><img src="imagens/bot_imprimir.gif" width="71" height="16" border="0" alt="imprimir"></a></td>
                <td width="51%"> 
                  <div align="right"><font color="#6666FF"> 
                    <%if fs.FileExists(realDir & "\imoveis\" & cod & "_1.jpg") then%>
                    <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#006666"><b>CLIQUE 
                    PARA AMPLIAR</b></font></font> <font color="#6666FF"> 
                    <%end if%>
                    </font></div>
                </td>
              </tr>
            </table>
            </td>
          </tr>
        </table>
      
    </td>
  </tr>
</table>
<%else%>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><b> 
N&atilde;o h&aacute; nenhum im&oacute;vel para visualizar.</b></font> 
<%end if%>
</html>
<%Response.flush%>