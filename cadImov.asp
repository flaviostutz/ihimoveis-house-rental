<html>
<!--#include file="common.inc" --> 
<!--#include file="myCommon.inc" --> 

<head>
<title>Imobiliária Higienópolis - Administração</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<style TYPE="text/css">
<!--
	A:link {color:"darkblue"; text-decoration: none}
	A:visited {color: "darkblue"; text-decoration: none}
	A:hover {color:"red"; text-decoration: underline}
-->
</style>

<body bgcolor="#FFFFFF" text="#000000">
<center>
<%
changeImages = Request.QueryString("changeImages")
cod = Request.QueryString("cod")
oldCod = Request.QueryString("oldCod")
transaction = Request.QueryString("transaction")
itype = Request.QueryString("itype")
price = ReplaceStr(Request.QueryString("price"),".","")
rooms = Request.QueryString("rooms")
area = ReplaceStr(Request.QueryString("area"),".","")
lote = ReplaceStr(Request.QueryString("lote"),".","")
cond = Request.QueryString("cond")
if Lcase(Request.QueryString("IPTU"))="isento" then
	IPTU = -1
else
	IPTU = replaceStr(Request.QueryString("IPTU"),".","")
end if

description = replaceStr(Request.QueryString("description"),chr(39),chr(34))
quarter = replaceStr(Request.QueryString("quarter"),chr(39),chr(34))
city = replaceStr(Request.QueryString("city"),chr(39),chr(34))
address = replaceStr(Request.QueryString("address"),chr(39),chr(34))
building = replaceStr(Request.QueryString("building"),chr(39),chr(34))


if cod <> "" and itype <> "" and transaction <> "" and quarter <> "" and address <> "" and city <> "" and price <> "" then
%>
<div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#006699"><b>Cadastro 
  de im&oacute;vel</b></font> </div>
<p align="left">
  
  <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#006699">Preparando 
    as fotos e cadastrando o im&oacute;vel. Aguarde alguns instantes...<br>
    <br>
    <a href="adminImov.asp"><b>Sess&atilde;o de administra&ccedil;&atilde;o</b></a><br>
    <a href="cadImovInterface.asp"><b>Cadastrar mais Im&oacute;veis</b></a><b> 
    | <a href="default.htm">Ir para o Site (off-line)</a><br>
    <br>
    </b></font></div>
  <%
	'initialize objects
	set Resizer = Server.CreateObject("ImageSize.AutoImageSize")
	Resizer.SetAspectRatioGuide(0) '0 - none, 1 - auto (preserves)
	set myConn = getMdb(Application("mdbFile"))

'//prepare photos and thumbnails to be sent<br>
on Error resume next 'if there isn't any file, it returns an error. So handle it!
	'delete last preparation
	fs.DeleteFile realDir & "\saida\imoveis\" & cod & "_*.*"

	for i = 1 to 8
		if Request.QueryString("comment"&i) <> "" or Request.QueryString("file"&i) <> "" then 
			if Request.QueryString("file"&i) <> "" then
				thumbImg = Resizer.ChangeImageSize(Request.QueryString("file"&i),Application("thumbWidth"),Application("thumbHeight"))
				thumbImg = Server.MapPath("/") & replaceStr(left(thumbImg, instr(1,thumbImg, ".jpg")+3),"/","\")
				picImg = Resizer.ChangeImageSize(Request.QueryString("file"&i),Application("imgWidth"),Application("imgHeight"))
				picImg = Server.MapPath("/") & replaceStr(left(picImg, instr(1,picImg, ".jpg")+3),"/","\")

				'copy to offLine
				fs.CopyFile thumbImg, realDir & "\imoveis\" & cod & "_" & i & "p.jpg",true
				fs.CopyFile picImg, realDir & "\imoveis\" & cod & "_" & i & ".jpg",true

				'copy to the saida folder
				fs.MoveFile thumbImg, realDir & "\saida\imoveis\" & cod & "_" & i & "p.jpg"
				fs.MoveFile picImg, realDir & "\saida\imoveis\" & cod & "_" & i & ".jpg"
			end if
			separateField "comment"&i,ReplaceStr(Request.QueryString("comment"&i),chr(39),chr(34)), commFields, commValues, true
		end if
	next
	%>
  <table width="670" border="0" cellspacing="5" cellpadding="5" bgcolor="#f0f0f0" align="center">
    <tr> 
      <td width="66%" valign="middle" align="center"> 
        <table border="0" width="98%" cellpadding="0" cellspacing="2">
            <tr> 
              <td colspan="3" height="17"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"> 
                <b><%=transactionStr(transaction)%> - <%=itypeStr(itype)%></b></font></td>
            </tr>
            <tr> 
              <td colspan="3"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Código 
                <%=cod%></font></td>
            </tr>
            <tr> 
              <td colspan="3"><font size="2" color="#000000" face="Verdana, Arial, Helvetica, sans-serif"><b><%=building%> 
                <%if building<>"" then%>
                - </b> 
                <%end if%>
                <%=quarter%> 
                <%if "building"<>"" then write "</b>"%>
                </font></td>
            </tr>
            <tr> 
              <td colspan="3" height="17"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">Endereço:</font></b> 
                <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=address%></font><font size="2"> 
                - </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=city%></font></font></td>
            </tr>
            <%if rooms <> "" then%>
            <tr> 
              <td colspan="2"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">N° 
                de quartos:</font></b> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rooms%></font></font> 
              </td>
            </tr>
            <%end if%>
            <%if area <> "" then%>
            <tr> 
              
            <td width="69%"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">&Aacute;rea 
              total:</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
              </font></b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"><%=area%></font></font> 
              m&sup2;</td>
            </tr>
            <%end if%>
            <%if lote <> "" then%>
            <tr> 
              
            <td width="69%"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">&Aacute;rea 
              lote:</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
              </font></b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"><%=lote%></font></font> 
              m&sup2;</td>
            </tr>
            <%end if%>
            <tr> 
              <td colspan="3" height="19"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">Valor:</font></b> 
                <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> <%=FormatCurrency(price)%></font></font></td>
            </tr>
            <%if cond<>"" then%>
            <tr> 
              <td colspan="3"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">Condom&iacute;nio:</font></b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"> 
                <%=FormatCurrency(cond)%></font></font></td>
            </tr>
            <%end if%>
            <%if IPTU<>"" then%>
            <tr> 
              <td colspan="3"><font face="Arial, Helvetica, sans-serif" color="#000000"><b><font size="2">IPTU:</font></b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"> 
                <%
				if IPTU>=0 then
					write FormatCurrency(IPTU)
				else 
	            	write "Isento"
				end if
				%>
                </font></font></td>
            </tr>
            <%end if%>
            <%if description<>"" then%>
            <tr> 
              <td colspan="3" valign="top"> 
                <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><br>
                  Detalhes:</b></font> <br>
                  <font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="2"><%=description%></font> 
                </p>
              </td>
            </tr>
            <%end if%>
          </table>
      </td>
      <td width="34%" valign="middle" align="center"> 
      <%if fs.FileExists(realDir & "\imoveis\" & cod & "_1.jpg") then%>
        
        <table border="0" cellspacing="9" cellpadding="0" align="right">
          <%for i=0 to 4%>
          <%if fs.FileExists(realDir & "\imoveis\" & cod & "_" & (i*2)+1 & ".jpg") then%>
          <tr align="center" valign="middle"> 
            <td width="50%"><img src=<%="file:///" & realDir & "\imoveis\" & cod & "_" & (i*2)+1 & "p.jpg"%> border="0" width="80" height="60"></td>
            <td width="50%"><%if fs.FileExists(realDir & "\imoveis\" & cod & "_" & (i*2)+2 & ".jpg") then%><img src=<%="file:///" & realDir & "\imoveis\" & cod & "_" & (i*2)+2 & "p.jpg"%> border="0" width="80" height="60"><%else%>&nbsp; <%end if%></td>
          </tr>
          <%end if%>
          <%next%>
        </table>
        <%else%>
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>N&atilde;o 
          h&aacute; fotos dispon&iacute;veis</b></font>
          <%end if%>
        </div>
      </td>
    </tr>
  </table>

  
<%
	if oldCod<>"" then
		'delete the record with the cod to be used if it exists
		sqlStr = "delete * from imoveis where cod='" & oldCod & "';"
		myConn.Execute sqlStr
	end if

	separateField "cod",cod,imovFields,imovValues,true
	separateField "transaction",transaction,imovFields,imovValues,false
	separateField "itype",itype,imovFields,imovValues,false
	separateField "description",description,imovFields,imovValues,true
	separateField "price",price,imovFields,imovValues,false
	separateField "quarter",quarter,imovFields,imovValues,true
	separateField "city",city,imovFields,imovValues,true
	separateField "address",address,imovFields,imovValues,true
	separateField "building",building,imovFields,imovValues,true
	separateField "cond",cond,imovFields,imovValues,false
	separateField "rooms",rooms,imovFields,imovValues,false
	separateField "area",area,imovFields,imovValues,false
	separateField "lote",lote,imovFields,imovValues,false
	separateField "IPTU",IPTU,imovFields,imovValues,true

	if commFields="" then
		sqlStr = "insert into imoveis (" & imovFields & ") values (" & imovValues & ");"
	else
		sqlStr = "insert into imoveis (" & imovFields & "," & commFields & ") values (" & imovValues & "," & commValues & ");"
	end if

on Error resume next
	myConn.Execute sqlStr
	myConn.close
	set myConn = Nothing
	
	fs.CopyFile realDir & "\" & Application("mdbFile"), realDir & "\saida\" & Application("mdbFile"), true
	if Err.description <> "" then
	%>
	<%=Err.description%><br>
	SQL: <%=sqlStr%> 
	<%else%>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#006699"><b><br>
  <br>
  <font size="3">Imóvel cadastrado com sucesso!</font></b></font><br>
  <%
	end if
else
%>
<p><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#006699"><b>Erro:</b></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#006699"><br>
  <br>
  Voc&ecirc; deve digitar todos os dados do im&oacute;vel para cadastr&aacute;-lo.</font></p>
<p><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#006699"> 
  <br>
  <input type="button" name="Button" value="Voltar" onClick=history.go(-1)>
  <br>
  </font><br>
  <%
end if
set fs = Nothing
set Resizer = Nothing
%>
</p>
</center>
</body>
</html>
