'Conectarse a CREDERE

Set conn = CreateObject("ADODB.Connection")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strConnect = "Provider=OraOLEDB.Oracle;Data Source=PRODUCCION;User Id=FNT_FCRUCES;Password=fonacot1709"
'strConnect = "Provider=OraOLEDB.Oracle;Data Source=DWH;User Id=FNT_ECRUCES;Password=fonacot03"
conn.Open strConnect

StrSQL = "SELECT CASE "
StrSQL = StrSQL & "WHEN TIPO_CLIENTE=1 THEN " & chr(39) & "01-Trabajador" & chr(39) & " "
StrSQL = StrSQL & "WHEN TIPO_CLIENTE=2 THEN " & chr(39) & "02-Centro Tra" & chr(39) & " "
StrSQL = StrSQL & "WHEN TIPO_CLIENTE=3 THEN " & chr(39) & "03-Establ Com" & chr(39) & " "
StrSQL = StrSQL & "WHEN TIPO_CLIENTE=4 THEN " & chr(39) & "04-Despacho" & chr(39) & " "
StrSQL = StrSQL & "END AS TIPO_CLIENTE "
StrSQL = StrSQL & ",COUNT(*) AS TOTAL "
StrSQL = StrSQL & "FROM TOPAZ.CL_CLIENTES "
StrSQL = StrSQL & "GROUP BY TIPO_CLIENTE "
StrSQL = StrSQL & "ORDER BY TIPO_CLIENTE "

'MsgBox Len(strSQL)
'StrSQL = "SELECT CLIENTE_ID,TIPO_PERSONA FROM TOPAZ.CL_CLIENTES WHERE CLIENTE_ID=" & chr(39) & 118724078 & chr(39) & " "

'MsgBox StrSQL

Set rs = conn.Execute(StrSQL)

iConta=1
Do While not rs.EOF 
	  Campo0=RS(0) 
	  If IsNull(Campo0) = True Then Campo0="" End If	  	
	  
	  Campo0=Trim(Campo0)
	  Campo0=replace(Campo0,Chr(13),"")
 	  Campo0=replace(Campo0,Chr(10),"")
	  Campo0=replace(Campo0,"|","")
	  Campo0=replace(Campo0,Chr(34),"")
	  
	  Campo1=RS(1) 
	  If IsNull(Campo1) = True Then Campo1="" End If	  	
	  
	  Campo1=Trim(Campo1)
	  Campo1=replace(Campo1,Chr(13),"")
 	  Campo1=replace(Campo1,Chr(10),"")	  
	  Campo1=replace(Campo1,"|","")	  
	  Campo1=replace(Campo1,Chr(34),"")
	  	  	
		'Inserta en Access

		Dim connStr, objConn, getNames
		connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=d:\Documents\INFONACOT\Biblioteca\Rutinas codigo\VBS\CREDERE_REG.accdb"

    'Define object type
    Set objConn = CreateObject("ADODB.Connection")
 
    'Open Connection
    objConn.open connStr

		'Define recordset and SQL query
		'strSQL = "INSERT INTO CREDERE (FECHA,REGISTRO) VALUES (#" & Date() & "#,'" & Campo0 & "')"

		'StrSQL = replace(strSQL,"'","")
		
		strConsulta = "01-Trabajador 02-Centro Tra 03-Establ Com 04-Despacho"
		
		strSQL2 = "INSERT INTO CL_CLIENTES (CONSULTA,TIPO_CLIENTE,TOTAL) VALUES ('" & strConsulta & "','" & Campo0 & "','" & Campo1 & "')"	
    'MsgBox StrSQL2
    		
		Set rsBeta = objConn.execute(strSQL2)
		
		'Close connection and release objects
		objConn.Close
		Set rsBeta = Nothing
		Set objConn = Nothing

    iConta=iConta+1
    rs.MoveNext
Loop
