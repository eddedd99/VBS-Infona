'Conectarse a CREDERE

Set conn = CreateObject("ADODB.Connection")
'Set objFSO = CreateObject("Scripting.FileSystemObject")

'strConnect = "Provider=OraOLEDB.Oracle;Data Source=DWH;User Id=FNT_ECRUCES;Password=fonacot03"
strConnect = "Provider=OraOLEDB.Oracle;Data Source=PRODUCCION;User Id=FNT_FCRUCES;Password=fonacot1709"
conn.Open strConnect

StrSQL = "Select "
StrSQL = StrSQL & "fecha_ultimo_movimiento,count(*) as Total,to_char(SYSDATE,'DD/MM/YYYY') "
StrSQL = StrSQL & "From Topaz.Tc_Imss_Trabajador_Diario "
StrSQL = StrSQL & "Where to_char(fecha_ultimo_movimiento,'YYYYMMDD') >= '20180101' "
StrSQL = StrSQL & "Group By fecha_ultimo_movimiento "
StrSQL = StrSQL & "Order By fecha_ultimo_movimiento desc "

Set rs = conn.Execute(StrSQL)

iConta=1
Do While not rs.EOF And iConta <= 1300
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

    'MsgBox Campo0 & " " & Campo1

    'Inserta en archivo plano
     
     'Create
     'filePath = "d:\Documents\INFONACOT\Biblioteca\Rutinas codigo\VBS\CL_CTES_" & Year(Now()) & Month(Now()) & Day(Now()) & ".txt"
     filePath = "d:\Documents\INFONACOT\Biblioteca\Rutinas codigo\VBS\IMSS_DIARIA.txt"
     
     Set objFSO = CreateObject("Scripting.FileSystemObject")

     If (objFSO.FileExists(filePath)) Then
        'Append
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objTxt = objFSO.opentextfile(filePath,8) 'Appending
        strData = CStr(Campo0) & "|" & CStr(Campo1) & "|" & CStr(Campo2)
        objTxt.WriteLine(strData)
        objTxt.Close
     Else
        'Create
        Set objTxt = objFSO.CreateTextFile(filePath)
        objTxt.WriteLine("Creacion " & Now())
        objTxt.WriteLine("FECHA_ULTIMO_MOVIMIENTO|TOTAL|SYSDATE")
        objTxt.Close
     End If

    iConta=iConta+1
    rs.MoveNext
Loop

    'Close connection and release objects
     Set conn = Nothing

'MsgBox "Termino"
