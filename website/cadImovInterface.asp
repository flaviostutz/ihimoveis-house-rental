<%Response.buffer=true%>
<%Response.Expires=0%>
<html>
<!--#include file="common.inc" --> 
<!--#include file="myCommon.inc" --> 
<script>
	function validateFields() {
		if (document.form1.cod.value == '' || document.form1.itype.value == '' || document.form1.transaction.value == '' || document.form1.quarter.value == '' || document.form1.address.value == '' || document.form1.city.value == '' || document.form1.price.value == '') {
			alert("Você deve preencher todos os campos marcados por *");
			return false;
		} else
			return true;
	}
</script>
<HEAD>
<meta http-equiv="pragma" content="no-cache">
<title>Imobiliária Higienópolis - Administração</title>

<style TYPE="text/css">
<!--
	A:link {color:"darkblue"; text-decoration: none}
	A:visited {color: "darkblue"; text-decoration: none}
	A:hover {color:"red"; text-decoration: underline}
-->
</style>

<BODY bgcolor="#FFFFFF">
<form name="form1" method="get" action="cadImov.asp">
  <table width="100%" border="0" cellspacing="0" cellpadding="3">
    <tr> 
      <td> 
        <table border="0" cellspacing="0" cellpadding="3" width="37%">
          <tr bgcolor="#006699"> 
            <td valign="middle" height="18" bgcolor="#006699"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Dados</b></font></td>
          </tr>
          <tr bgcolor="#D5D9EE"> 
            <td valign="top" bgcolor="#D5D9EE"> 
              <table border="0" cellspacing="1" cellpadding="0" width="100%">
                <tr bgcolor="#DAE8FE"> 
                  <td align="right" width="27%"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>*Transa&ccedil;&atilde;o</b></font></td>
                  <td width="73%" bgcolor="#DAE8FE"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <select name="transaction">
                      <option value=""></option>
                      <option value="1"<%=selItem(1,Request.QueryString("transaction"))%>>Loca&ccedil;&atilde;o</option>
                      <option value="2"<%=selItem(2,Request.QueryString("transaction"))%>>Venda</option>
                    </select>
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td align="right" width="27%"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>*Tipo</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <select name="itype">
                      <option value=""></option>
                      <option value="1"<%=selItem(1,Request.QueryString("itype"))%>>Casa</option>
                      <option value="2"<%=selItem(2,Request.QueryString("itype"))%>>Apartamento</option>
                      <option value="3"<%=selItem(3,Request.QueryString("itype"))%>>Sala/Loja</option>
                      <option value="4"<%=selItem(4,Request.QueryString("itype"))%>>Terreno/Chácara</option>
                    </select>
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>*Cod</b></font></td>
                  <td width="73%" bgcolor="#DAE8FE"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <input type="text" name="cod" size="5" maxlength="5" value="<%=Request.QueryString("cod")%>">
                    <input type="hidden" name="oldCod" value="<%=Request.QueryString("cod")%>">
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"> 
                    <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Nome/edif&iacute;cio</b></font></div>
                  </td>
                  <td width="73%"> 
                    <div align="left"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                      <input type="text" name="building" size="26" maxlength="26" value="<%=Request.QueryString("building")%>">
                      </font></b></div>
                  </td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>*Bairro</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <input type="text" name="quarter" size="26" maxlength="26" value="<%=Request.QueryString("quarter")%>">
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"> 
                    <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>*Endere&ccedil;o</b></font></div>
                  </td>
                  <td width="73%"> 
                    <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><b> 
                      <input type="text" name="address" size="35" maxlength="35" value="<%=Request.QueryString("address")%>">
                      </b></font></div>
                  </td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>*Cidade</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <input type="text" name="city" size="25" maxlength="25" value="<%if Request.QueryString("city")<>"" then
			  	write Request.QueryString("city")
			else
			  	write("Londrina - PR")
			end if%>">
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Quartos</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <input type="text" name="rooms" size="5" maxlength="2" value="<%=Request.QueryString("rooms")%>">
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>&Aacute;rea 
                    total</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <input type="text" name="area" size="7" maxlength="7" value="<%=Request.QueryString("area")%>">
                    </font></b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">m&sup2;</font></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>&Aacute;rea 
                    lote</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <input type="text" name="lote" size="7" maxlength="7" value="<%=Request.QueryString("lote")%>">
                    </font></b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">m&sup2;</font></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>*Valor 
                    (R$)</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <input type="text" name="price" value="<%=Request.QueryString("price")%>" maxlength="10" size="10">
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Condom&iacute;nio</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <input type="text" name="cond" size="10" maxlength="10" value="<%=Request.QueryString("cond")%>">
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>IPTU</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <input type="text" name="IPTU" size="10" maxlength="10" value="<%if Request.QueryString("city")<>"" then
			  	write Request.QueryString("IPTU")
			else
			  	write("Isento")
			end if%>">
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Detalhes</b></font></td>
                  <td width="73%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                    <textarea name="description" cols="45" rows="3" onKeyDown="if(form1.description.value.length>180) form1.description.value=form1.description.value.substring(0,180);"><%=Request.QueryString("description")%></textarea>
                    </font></b></td>
                </tr>
                <tr bgcolor="#DAE8FE"> 
                  <td width="27%" align="right">&nbsp; </td>
                  <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">*</font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><i> 
                    preenchimento obrigat&oacute;rio</i></font></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <br>
        <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><a name="images"></a></b></font> 
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td bgcolor="#006699" height="18" valign="middle"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Imagens 
        e coment&aacute;rios</b></font></td>
    </tr>
    <tr> 
      <td bgcolor="#D5D9EE"> 
        <table border="0" cellpadding="5" cellspacing="1" width="100%" align="left">
          <tr> 
            <td bgcolor="#dae8fe" width="29"> 
              <div align="right"><font color="#006699" size="5" face="Verdana, Arial, Helvetica, sans-serif"> 
                1</font></div>
            </td>
            <td bgcolor="#DAE8FE" width="327"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="file" name="file1">
              <br>
              <input type="text" name="comment1" maxlength="35" size="35" value="<%=Request.QueryString("comment1")%>">
              <a href=#preview onClick="document.picPreview.src=document.form1.file1.value">ver</a></font></td>
            <td bgcolor="#DAE8FE" width="30" align="right"> 
              <div align="right"><font color="#006699" size="5" face="Verdana, Arial, Helvetica, sans-serif">5</font></div>
            </td>
            <td bgcolor="#DAE8FE" width="326"><font color="#006699" size="1"><font face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="file" name="file5">
              <br>
              <input type="text" name="comment5" maxlength="35" size="35" value="<%=Request.QueryString("comment5")%>">
              </font><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href=#preview onClick="document.picPreview.src=document.form1.file5.value">ver</a></font></font></td>
          </tr>
          <tr> 
            <td bgcolor="#DAE8FE" width="29"> 
              <div align="right"><font color="#006699" size="5" face="Verdana, Arial, Helvetica, sans-serif">2</font></div>
            </td>
            <td bgcolor="#DAE8FE" width="327"><font color="#006699" size="1"><font face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="file" name="file2">
              <br>
              <input type="text" name="comment2" maxlength="35" size="35" value="<%=Request.QueryString("comment2")%>">
              </font><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href=#preview onClick="document.picPreview.src=document.form1.file2.value">ver</a></font></font></td>
            <td bgcolor="#DAE8FE" width="30" align="right"> 
              <div align="right"><font color="#006699" size="5" face="Verdana, Arial, Helvetica, sans-serif">6</font></div>
            </td>
            <td bgcolor="#DAE8FE" width="326"><font color="#006699" size="1"><font face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="file" name="file6">
              <br>
              <input type="text" name="comment6" maxlength="35" size="35" value="<%=Request.QueryString("comment6")%>">
              </font><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href=#preview onClick="document.picPreview.src=document.form1.file6.value">ver</a></font></font></td>
          </tr>
          <tr> 
            <td bgcolor="#DAE8FE" width="29"> 
              <div align="right"><font color="#006699" size="5" face="Verdana, Arial, Helvetica, sans-serif">3</font></div>
            </td>
            <td bgcolor="#DAE8FE" width="327"><font color="#006699" size="1"><font face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="file" name="file3">
              <br>
              <input type="text" name="comment3" maxlength="35" size="35" value="<%=Request.QueryString("comment3")%>">
              </font><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href=#preview onClick="document.picPreview.src=document.form1.file3.value">ver</a></font></font></td>
            <td bgcolor="#DAE8FE" width="30" align="right"> 
              <div align="right"><font color="#006699" size="5" face="Verdana, Arial, Helvetica, sans-serif">7</font></div>
            </td>
            <td bgcolor="#DAE8FE" width="326"><font color="#006699" size="1"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="file" name="file7">
              </font><font face="Verdana, Arial, Helvetica, sans-serif"><br>
              <input type="text" name="comment7" maxlength="35" size="35" value="<%=Request.QueryString("comment7")%>">
              </font><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href=#preview onClick="document.picPreview.src=document.form1.file7.value">ver</a></font></font></td>
          </tr>
          <tr> 
            <td bgcolor="#DAE8FE" width="29"> 
              <div align="right"><font color="#006699" size="5" face="Verdana, Arial, Helvetica, sans-serif">4</font></div>
            </td>
            <td bgcolor="#DAE8FE" width="327"><font color="#006699" size="1"><font face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="file" name="file4">
              <br>
              <input type="text" name="comment4" maxlength="35" size="35" value="<%=Request.QueryString("comment4")%>">
              </font><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href=#preview onClick="document.picPreview.src=document.form1.file4.value">ver</a><br>
              </font></font></td>
            <td bgcolor="#DAE8FE" width="30" align="right"> 
              <div align="right"><font color="#006699" size="5" face="Verdana, Arial, Helvetica, sans-serif">8</font></div>
            </td>
            <td bgcolor="#DAE8FE" width="326"><font color="#006699" size="1"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="file" name="file8">
              </font><font face="Verdana, Arial, Helvetica, sans-serif"><br>
              <input type="text" name="comment8" maxlength="35" size="35" value="<%=Request.QueryString("comment8")%>">
              </font><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href=#preview onClick="document.picPreview.src=document.form1.file8.value">ver</a></font></font></td>
          </tr>
          <tr> 
            <td bgcolor="#DAE8FE" colspan="4"><font size="1">&nbsp; </font></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td> 
        <div align="center"> 
          <p align="center"> <br>
            <input type="submit" name="submit1" value="Cadastrar" onClick="return validateFields()">
            <input type="reset" name="Submit2" value="Apagar">
          </p>
          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><a name="preview"></a></b></font><br>
            <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#333366"><b>Preview:</b></font><br>
            <br>
            <a href="#images"><img src="" name="picPreview" border="0" alt="clique para voltar"></a> 
          </div>
        </div>
      </td>
    </tr>
  </table>
  <div align="center"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
    </font> <br>
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#006699"><a href="adminImov.asp"><b>Sess&atilde;o 
    de administra&ccedil;&atilde;o</b></a></font><br>
  </div>
  <p>&nbsp;</p>
</form>
</BODY>
</HTML>

<%Response.flush%>