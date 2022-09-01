<html>
<!--#include file="common.inc" --> 
<!--#include file="myCommon.inc" --> 
<%
On error resume next
transaction = Request.QueryString("transaction")
itype = Request.QueryString("itype")
%>
<body bgcolor="#FFFFFF" leftmargin="8" topmargin="6" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<table border="0" cellspacing="0" cellpadding="5" width="100%" height="100%">
  <tr>
<td> 
	<table border="0" cellspacing="0" cellpadding="0" width="1">
	<tr>
		<td width="9" rowspan="2"><img border="0" src="imagens/img_top_barra.gif" width="9" height="49"></td>
		<td><img src="imagens/img_top_<%=left(LCase(transactionStr(transaction)),4)%>.gif" alt="<%=transactionStr(transaction)%>"></td>
	</tr>
    <tr>
	<td width="183">
			<img src="imagens/img_top_<%=left(LCase(iTypeStr(iType)),4)%>.gif" alt="<%=LCase(iTypeStr(iType))%>">
	</td>
	</tr>
	</table>
    </td>
    <td width="352">
	  <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
        <tr align="right"> 
          <td><a href=mailto:info@ihimoveis.com.br><img border="0" src="imagens/img_top_logoimob.gif" align="right" width="222" height="49"></a> 
          </td>
        </tr>
        <tr align="right"> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="97%">
                  <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#000000">(43) 
                    321-7580 - <a href="mailto:info@ihimoveis.com.br">info@<b>ih</b>imoveis.com.br</a></font></div>
                </td>
                <td width="1%">&nbsp;</td>
              </tr>
            </table>
            
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
</body>
</html>