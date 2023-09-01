# TraitementExcelAsBdd
Traitement de excel comme base de donné

## Le code
Sub UtiliserOleDb()<br>
    Dim conn As Object<br>
    Dim rs As Object<br>
    Dim strSQL As String<br>
    Dim chemin As String<br>
    
    chemin = ThisWorkbook.FullName<br>
    
    Set conn = CreateObject("ADODB.Connection")<br>
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & chemin & ";Extended<br> Properties=""Excel 12.0 Xml;HDR=YES;"""<br>
    
    strSQL = "SELECT nom FROM [Feuil1$] where nom like '%a%'"<br>
    Set rs = conn.Execute(strSQL)<br>
    
    Dim i As Long
    i = 2
    
    Dim str As String
    str = ""
    

    Do Until rs.EOF
        str = str & "," & rs.Fields("nom").Value<br>
        Range("c" & i).Value = rs.Fields("nom").Value<br>
        
        i = i + 1
        
        rs.MoveNext
    Loop
    Range("e3").Value = str<br>

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub

