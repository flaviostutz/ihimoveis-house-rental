<%
function iTypeStr(tempNbr)
	select case tempNbr
	case 0
		iTypeStr = "Im?vei"
	case 1
		iTypeStr = "Casa"
	case 2
		iTypeStr = "Apartamento"
	case 3
		iTypeStr = "Sala/Loja"
	case 4
		iTypeStr = "Terreno"
	end select
end function

function transactionStr(tempNbr)
	select case tempNbr
	case 1
		transactionStr = "Loca??o"
	case 2
		transactionStr = "Venda"
	end select
end function

Set fs = CreateObject("Scripting.FileSystemObject")
realDir = Server.MapPath(".")

%>
