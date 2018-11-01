'Conectarse a CREDERE

Set conn = CreateObject("ADODB.Connection")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strConnect = "Provider=OraOLEDB.Oracle;Data Source=PRODUCCION;User Id=FNT_FCRUCES;Password=fonacot1709"
'strConnect = "Provider=OraOLEDB.Oracle;Data Source=DWH;User Id=FNT_ECRUCES;Password=fonacot0503"
conn.Open strConnect

StrSQL = "SELECT owner, table_name, num_rows, sample_size, last_analyzed FROM all_tables WHERE OWNER = 'TOPAZ' ORDER BY TABLE_NAME ASC "

Set rs = conn.Execute(StrSQL)

iConta=1
Do While not rs.EOF 
	  Campo0=RS(0) 'owner
	  Campo1=RS(1) 'table_name
	  Campo2=RS(2) 'num_rows
	  Campo3=RS(3) 'sample_size
	  Campo4=RS(4) 'last_analyzed
  
	  If IsNull(Campo0) = True Then Campo0="" End If	  	
	  If IsNull(Campo1) = True Then Campo1="" End If
	  If IsNull(Campo2) = True Then Campo2="" End If
	  If IsNull(Campo3) = True Then Campo3="" End If
	  If IsNull(Campo4) = True Then Campo4="" End If	  
	  
	  Campo0=Trim(Campo0)
	  Campo1=Trim(Campo1)
	  Campo2=Trim(Campo2)
	  Campo3=Trim(Campo3)
	  Campo4=Trim(Campo4) 

	  Campo0=replace(Campo0,Chr(13),"")
 	  Campo0=replace(Campo0,Chr(10),"")	  
	  Campo0=replace(Campo0,"|","")	  
	  Campo0=replace(Campo0,Chr(34),"")	  

	  Campo1=replace(Campo1,Chr(13),"")
 	  Campo1=replace(Campo1,Chr(10),"")	  
	  Campo1=replace(Campo1,"|","")	  
	  Campo1=replace(Campo1,Chr(34),"")
	  
	  Campo2=replace(Campo2,Chr(13),"")
 	  Campo2=replace(Campo2,Chr(10),"")	  
	  Campo2=replace(Campo2,"|","")	  
	  Campo2=replace(Campo2,Chr(34),"")

	  Campo3=replace(Campo3,Chr(13),"")
 	  Campo3=replace(Campo3,Chr(10),"")	  
	  Campo3=replace(Campo3,"|","")	  
	  Campo3=replace(Campo3,Chr(34),"")
	  
 	  Campo4=replace(Campo4,Chr(13),"")
 	  Campo4=replace(Campo4,Chr(10),"")	  
	  Campo4=replace(Campo4,"|","")	  
	  Campo4=replace(Campo4,Chr(34),"")

		'Inserta en Access

		Dim connStr, objConn, getNames
		'connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=D:\Documents\INFONACOT\Biblioteca\Rutinas codigo\VBS\CREDERE_REG.accdb"
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=D:\Documents\INFONACOT\Biblioteca\Rutinas codigo\VBS\CREDERE_REG.accdb"

    'Define object type
    Set objConn = CreateObject("ADODB.Connection")
 
    'Open Connection
    objConn.open connStr
    
    StrSQL=Replace(StrSQL,"'","")

		strSQL2 = "INSERT INTO CREDERE_ALL (OWNER,TABLE_NAME,NUM_ROWS,SAMPLE_SIZE,LAST_ANALYZED) VALUES ('" & Campo0 & "','" & Campo1 & "','" & Campo2 & "','" & Campo3 & "','" & Campo4 & "')"
    'MsgBox strSQL2

		Set rsBeta = objConn.execute(strSQL2)
		
		'Close connection and release objects
		objConn.Close
		Set rsBeta = Nothing
		Set objConn = Nothing

    iConta=iConta+1
    rs.MoveNext
Loop

'MsgBox "Fin " & Now()

