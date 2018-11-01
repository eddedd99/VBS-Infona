'Conectarse a CREDERE

Set conn = CreateObject("ADODB.Connection")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strConnect = "Provider=OraOLEDB.Oracle;Data Source=PRODUCCION;User Id=FNT_FCRUCES;Password=fonacot1709"
'strConnect = "Provider=OraOLEDB.Oracle;Data Source=DWH;User Id=FNT_ECRUCES;Password=fonacot03"
conn.Open strConnect

StrSQL = "SELECT CASE "
StrSQL = StrSQL & "WHEN STATUS=0 THEN " & chr(39) & "0 XAUTORIZAR" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=1 THEN " & chr(39) & "1 ACTIVO" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=6 THEN " & chr(39) & "6 JURIDICO" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=7 THEN " & chr(39) & "7 FALTAPAGO" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=9 THEN " & chr(39) & "9 BAJAPERMAN" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=10 THEN " & chr(39) & "10 SOLO RECUPERAC" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=11 THEN " & chr(39) & "11 REACTIVACION" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=17 THEN " & chr(39) & "17 PARAAUTORIZAR" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=18 THEN " & chr(39) & "18 PARA_AUTORIZ_DIRECT" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=19 THEN " & chr(39) & "19 RECHAZADO" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=2 THEN " & chr(39) & "2 SUSPENDIDO" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=3 THEN " & chr(39) & "3 EXTRAJUDICIAL" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=4 THEN " & chr(39) & "4 QUIEBRA" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=5 THEN " & chr(39) & "5 HUELGA" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=8 THEN " & chr(39) & "8 ILOCALIZABLE" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=12 THEN " & chr(39) & "12 BAJA OPERATIVA" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=13 THEN " & chr(39) & "13 BAJA PETIC DIRECTOR" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=14 THEN " & chr(39) & "14 BAJA DEPURACION COBRANZ" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=15 THEN " & chr(39) & "15 BAJA NO EJERCER CREDITO" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=16 THEN " & chr(39) & "16 RENUENTE" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=20 THEN " & chr(39) & "20 CEDULA SIN PAGO 31 DIAS" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=21 THEN " & chr(39) & "21 CEDULA PAGO MENOR 90%" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=22 THEN " & chr(39) & "22 NO CUMPLE ANTIGUEDAD" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=23 THEN " & chr(39) & "23 CONDICIONES N BURO" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=24 THEN " & chr(39) & "24 CAMBIO CONDICION BURO" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=25 THEN " & chr(39) & "25 SIN DOCS SOLO RECUPERAC" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=30 THEN " & chr(39) & "30 AFILIACION MICROSITIO" & chr(39) & " "
StrSQL = StrSQL & "WHEN STATUS=88 THEN " & chr(39) & "88 PERIODO DE GRACIA" & chr(39) & " "
StrSQL = StrSQL & "ELSE TO_CHAR(STATUS) "
StrSQL = StrSQL & "END STATUS_DESC "
StrSQL = StrSQL & ",COUNT(*) AS TOTAL "
StrSQL = StrSQL & "FROM TOPAZ.CL_CLIENTES "
StrSQL = StrSQL & "WHERE TZ_LOCK=0 "
StrSQL = StrSQL & "AND TIPO_CLIENTE=2 "
StrSQL = StrSQL & "GROUP BY STATUS "
StrSQL = StrSQL & "ORDER BY STATUS "

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
	  	  	
		'Inserta en Access

		Dim connStr, objConn, getNames
		connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=d:\Documents\INFONACOT\Biblioteca\Rutinas codigo\VBS\CREDERE_REG.accdb"

    'Define object type
    Set objConn = CreateObject("ADODB.Connection")
 
    'Open Connection
    objConn.open connStr

		strConsulta = "02-Centro Trab"
		
		strSQL2 = "INSERT INTO CL_CLIENTES_ESTATUS (CONSULTA,VALOR,TOTAL) VALUES ('" & strConsulta & "','" & Campo0 & "','" & Campo1 & "')"	
    'MsgBox StrSQL2
    		
		Set rsBeta = objConn.execute(strSQL2)
		
		'Close connection and release objects
		objConn.Close
		Set rsBeta = Nothing
		Set objConn = Nothing

    iConta=iConta+1
    rs.MoveNext
Loop

