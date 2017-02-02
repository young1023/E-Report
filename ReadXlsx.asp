<HTML>
<HEAD>
<TITLE></TITLE>
<META name="description" content="">
<META name="keywords" content="">
<META name="generator" content="CuteHTML">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0000FF" VLINK="#800080">


<%

    sFilePath = "d:\data\cp.xlsx"

    sDataDir = "d:\data\"
 
    
     sSheetName = "CP"

  
    sFileConnectionString ="Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};READONLY=TRUE;DBQ="&sFilePath&";DefaultDir="&sDataDir&";"
   
    sFileSQL = "SELECT * FROM [" & sSheetName & "$]"

    set oFileCN = server.createobject("ADODB.Connection") 

    oFileCN.Open sFileConnectionString 'abro el exel
        set oFileRS = oFileCN.Execute(sFileSQL) 'selecciono los registros
        if not oFileRS.EOF then 'si es distinto de EOF obtengo un array con los valores de las columnas
            aSourceData = oFileRS.getRows()
        end if
    oFileRS.Close
    set oFileRS = nothing
    oFileCN.Close
    set oFileCN = nothing
    
    'asigno los mails a la variabls "readDatafile"
    readDatafile = aSourceData


%>

</BODY>
</HTML>
