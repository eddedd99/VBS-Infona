'Conectarse a CREDERE

Set conn = CreateObject("ADODB.Connection")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strConnect = "Provider=OraOLEDB.Oracle;Data Source=PRODUCCION;User Id=FNT_FCRUCES;Password=fonacot1709"
'strConnect = "Provider=OraOLEDB.Oracle;Data Source=DWH;User Id=FNT_ECRUCES;Password=fonacot03"
conn.Open strConnect

StrSQL = "SELECT A.TIPO_CLIENTE "
StrSQL = StrSQL & ",B.NOMBRE_SUCURSAL "
StrSQL = StrSQL & ",COUNT(*) AS TOTAL "
StrSQL = StrSQL & "FROM TOPAZ.CL_CLIENTES A "
StrSQL = StrSQL & "LEFT JOIN TOPAZ.TC_SUCURSALES B ON A.SUCURSAL_ID=B.SUCURSAL_ID "
StrSQL = StrSQL & "WHERE A.TZ_LOCK=0 "
StrSQL = StrSQL & "AND B.TZ_LOCK=0 "
StrSQL = StrSQL & "AND A.TIPO_CLIENTE IN (1,2) "
StrSQL = StrSQL & "GROUP BY A.TIPO_CLIENTE,B.NOMBRE_SUCURSAL "
StrSQL = StrSQL & "ORDER BY TIPO_CLIENTE ASC,COUNT(*) DESC "

Set rs = conn.Execute(StrSQL)

iConta=1
Do While not rs.EOF 
	  Campo0=RS(0) 
	  If IsNull(Campo0) = True Then Campo0="" End If	  	
	  
	  Campo0=Trim(Campo0)
	  Campo0=replace(Campo0,Chr(13),"")
 	  Campo0=replace(Campo0,Chr(10),"")
	  Campo0=replace(Campo0,"|","")
	  Campo0=replace(Campo0,"*","")
	  Campo0=replace(Campo0,Chr(34),"")
	  
	  Campo1=RS(1) 
	  If IsNull(Campo1) = True Then Campo1="" End If	  	
	  
	  Campo1=Trim(Campo1)
	  Campo1=replace(Campo1,Chr(13),"")
 	  Campo1=replace(Campo1,Chr(10),"")	  
	  Campo1=replace(Campo1,"|","")
	  Campo1=replace(Campo1,"*","")
	  Campo1=replace(Campo1,Chr(34),"")

	  Campo2=RS(2) 
	  If IsNull(Campo2) = True Then Campo2="" End If	  	
	  
	  Campo2=Trim(Campo2)
	  Campo2=replace(Campo2,Chr(13),"")
 	  Campo2=replace(Campo2,Chr(10),"")	  
	  Campo2=replace(Campo2,"|","")
	  Campo2=replace(Campo2,"*","")
	  Campo2=replace(Campo2,Chr(34),"")

		'Inserta en Access

		Dim connStr, objConn, getNames
		connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=d:\Documents\INFONACOT\Biblioteca\Rutinas codigo\VBS\CREDERE_REG.accdb"

    'Define object type
    Set objConn = CreateObject("ADODB.Connection")
 
    'Open Connection
    objConn.open connStr

		strConsulta = "Clientes por Sucursal"
		
		strSQL2 = "INSERT INTO CL_CLIENTES_SUCURSAL (CONSULTA,TIPO_CLIENTE,SUCURSAL,TOTAL) VALUES ('" & strConsulta & "','" & Campo0 & "','" & Campo1 & "','" & Campo2 & "')"	
    'MsgBox StrSQL2
    		
		Set rsBeta = objConn.execute(strSQL2)
		
		'Close connection and release objects
		objConn.Close
		Set rsBeta = Nothing
		Set objConn = Nothing

    iConta=iConta+1
    rs.MoveNext
Loop

