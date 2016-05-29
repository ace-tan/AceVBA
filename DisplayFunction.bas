Attribute VB_Name = "DisplayFunction"
Option Compare Database
Option Explicit

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long

'Testing for internet connection
Public Function IsInternetConnected() As Boolean
    IsInternetConnected = InternetGetConnectedStateEx(0, "", 254, 0)
End Function

Function ClearList(lst As ListBox) As Boolean
On Error GoTo Err_ClearList
    'Purpose:   Unselect all items in the listbox.
    'Return:    True if successful
    Dim varItem As Variant

    If lst.MultiSelect = 0 Then
        lst = Null
    Else
        For Each varItem In lst.ItemsSelected
            lst.Selected(varItem) = False
        Next
    End If

    ClearList = True

Exit_ClearList:
    Exit Function

Err_ClearList:
    Call LogError(Err.number, Err.Description, "ClearList()")
    Resume Exit_ClearList
End Function

Public Function SelectAll(lst As ListBox) As Boolean
On Error GoTo Err_Handler
    'Purpose:   Select all items in the multi-select list box.
    'Return:    True if successful
    Dim lngRow As Long

    If lst.MultiSelect Then
        For lngRow = 0 To lst.ListCount - 1
            lst.Selected(lngRow) = True
        Next
        SelectAll = True
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    Call LogError(Err.number, Err.Description, "SelectAll()")
    Resume Exit_Handler
End Function


Function CreatePKIndexes(strTableName As String, strID As String)
Dim dbs As DAO.Database
Dim tdf As DAO.TableDef
Dim idx As DAO.index
Dim fld As DAO.Field
Dim strPKey As String
Dim strIdxFldName As String
Dim intCounter As Integer

Set dbs = CurrentDb
Set tdf = dbs.TableDefs(strTableName)

'Check if a Primary Key exists.
'If so, delete it.
strPKey = GetPrimaryKey(tdf)

If Len(strPKey) > 0 Then
   tdf.Indexes.Delete strPKey
   tdf.Fields.Delete strID 'delete primary key column
End If

'add the column back
Set fld = tdf.CreateField(strID, dbLong)
fld.Attributes = dbAutoIncrField 'set to autoincrement
tdf.Fields.Append fld
    
'create primary key
Set idx = tdf.CreateIndex(strID)
idx.Fields.Append idx.CreateField(strID)
idx.CreateField "ContactID"
idx.Primary = True
idx.Unique = True
tdf.Indexes.Append idx

Set fld = Nothing
Set idx = Nothing
Set tdf = Nothing
Set dbs = Nothing
End Function

Function GetPrimaryKey(tdf As DAO.TableDef) As String
'Determine if the specified Primary Key exists
Dim idx As DAO.index

For Each idx In tdf.Indexes
    If idx.Primary Then
        'If a Primary Key exists, return its name
        GetPrimaryKey = idx.Name
        Exit Function
    End If
Next idx

'If no Primary Key exists, return empty string
GetPrimaryKey = vbNullString
End Function


Function FileFolderExists(strFullPath As String) As Boolean
On Error GoTo EarlyExit
If Not Dir(strFullPath, vbDirectory) = vbNullString Then
    FileFolderExists = True
Else
    FileFolderExists = False
End If
EarlyExit:
    On Error GoTo 0
End Function

Function FileExists(stFile As String) As Boolean
'check the file exist
If Dir(stFile) <> "" Then
    FileExists = True
Else
    FileExists = False
End If
End Function

Function LimitChange(ctl As Control, iMaxLen As Integer)
On Error GoTo Err_LimitChange
    ' Purpose:  Limit the text in an unbound text box/combo.
    ' Usage:    In the control's Change event procedure:
    '               Call LimitChange(Me.MyTextBox, 12)
    ' Note:     Requires LimitKeyPress() in control's KeyPress event also.

    If Len(ctl.text) > iMaxLen Then
        MsgBox "Truncated to " & iMaxLen & " characters.", vbExclamation, "Too long"
        ctl.text = Left(ctl.text, iMaxLen)
        ctl.SelStart = iMaxLen
    End If

Exit_LimitChange:
    Exit Function

Err_LimitChange:
    Call LogError(Err.number, Err.Description, "LimitChange()")
    Resume Exit_LimitChange
End Function

Function SetSeed(strTable As String, strAutoNum As String, lngID As Long) As Boolean
    'Purpose:   Set the Seed of an AutoNumber using ADOX.
    Dim cat As New ADOX.Catalog
    
    Set cat.ActiveConnection = CurrentProject.Connection
    cat.Tables(strTable).Columns(strAutoNum).Properties("Seed") = lngID
    Set cat = Nothing
    SetSeed = True
End Function

Function setVariable$(colName As String, tableName As String, colNameCondition As String, para As String)
    Dim counter As Integer
    Dim var As String
    counter = DCount(colName, tableName, colNameCondition & "= " & para)
    If counter > 0 Then
        var = DLookup(colName, tableName, colNameCondition & "= " & para)
        If StrComp(colName, "[Unit1]", vbTextCompare) = 0 Or StrComp(colName, "[Unit2]", vbTextCompare) = 0 Then
            var = Replace(var, "#", "")
            var = Replace(var, "-", "")
        End If
    Else
        var = ""
    End If
    setVariable = var
End Function

Function checkField(rs As DAO.Recordset, fieldName As String) As String
If IsNull(rs.Fields(fieldName)) = False Then
     checkField = rs.Fields(fieldName)
Else
     checkField = ""
End If
End Function

Function checkFieldCurrency(rs As DAO.Recordset, fieldName As String) As String
If IsNull(rs.Fields(fieldName)) = False Then
     checkFieldCurrency = rs.Fields(fieldName)
Else
     checkFieldCurrency = "0"
End If
End Function
'used in frmCurrency page
Function YahooCurrencyConverter(ByVal strFromCurrency, ByVal strToCurrency, Optional ByVal strResultType = "Value") As Double
On Error GoTo ErrorHandler
 
'Init
Dim strURL As String
Dim objXMLHttp As Object
Dim strRes As String, dblRes As Double
 
Set objXMLHttp = CreateObject("MSXML2.ServerXMLHTTP")
strURL = "http://finance.yahoo.com/d/quotes.csv?e=.csv&f=c4l1&s=" & strFromCurrency & strToCurrency & "=X"
 
'Send XML request
With objXMLHttp
    .Open "GET", strURL, False
    .setRequestHeader "Content-Type", "application/x-www-form-URLEncoded"
    .Send
    strRes = .responseText
End With
 
'Parse response
dblRes = Val(Split(strRes, ",")(1))

YahooCurrencyConverter = dblRes

CleanExit:
    Set objXMLHttp = Nothing
Exit Function
 
ErrorHandler:
    YahooCurrencyConverter = 0
    MsgBox "There is no internet connection.", vbInformation, "Internet Connection"
    GoTo CleanExit
End Function


Function DeleteAllRelationships() As String
' WARNING: Deletes all relationships in the current database.
    Dim db As Database      ' Current DB
    Dim rex As Relations    ' Relations of currentDB.
    Dim rel As relation     ' Relationship being deleted.
    Dim iKt As Integer      ' Count of relations deleted.
    Dim sMsg As String      ' MsgBox string.

    sMsg = "About to delete ALL relationships between tables in the current database." & vbCrLf & "Continue?"
   ' If MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Are you sure?") = vbNo Then
    '    DeleteAllRelationships = "Operation cancelled"
  '      Exit Function
 '   End If

    Set db = CurrentDb()
    Set rex = db.Relations
    iKt = rex.count
    Do While rex.count > 0
        'Debug.Print rex(0).Name
        rex.Delete rex(0).Name
    Loop
    DeleteAllRelationships = iKt & " relationship(s) deleted"
End Function

Function CreateKeyAdox(tableFrom As String, tableTo As String, columnName As String)
    'Purpose:   Show how to create relationships using ADOX.
    Dim cat As New ADOX.Catalog
    Dim tbl As ADOX.Table
    Dim ky As New ADOX.Key
    
    Set cat.ActiveConnection = CurrentProject.Connection
    Set tbl = cat.Tables(tableFrom)
    
    'Create as foreign key to tblAdoxContractor.ContractorID
    With ky
        .Type = adKeyForeign
        .Name = tableFrom
        .RelatedTable = tableTo
        .Columns.Append columnName      'Just one field.
        .Columns(columnName).RelatedColumn = columnName
        .DeleteRule = adRICascade   'Cascade to Null on delete.
    End With
    tbl.Keys.Append ky
    
    Set ky = Nothing
    Set tbl = Nothing
    Set cat = Nothing
   ' Debug.Print "Key created."
End Function

Function CreateRelationDAO(primaryTable As String, foreignTable As String, columnName As String, relation As String)
    Dim db As DAO.Database
    Dim rel As DAO.relation
    Dim fld As DAO.Field
    
    'Initialize
    Set db = CurrentDb()
    
    'Create a new relation.
    Set rel = db.CreateRelation(relation)
    
    'Define its properties.
    With rel
        'Specify the primary table.
        .Table = primaryTable
        'Specify the related table.
        .foreignTable = foreignTable
        'Specify attributes for cascading updates and deletes.
        .Attributes = dbRelationDeleteCascade
        
        'Add the fields to the relation.
        'Field name in primary table.
        Set fld = .CreateField(columnName)
        'Field name in related table.
        fld.ForeignName = columnName
        'Append the field.
        .Fields.Append fld
        
        'Repeat for other fields if a multi-field relation.
    End With
    
    'Save the newly defined relation to the Relations collection.
    db.Relations.Append rel
    
    'Clean up
    Set fld = Nothing
    Set rel = Nothing
    Set db = Nothing
    'Debug.Print "Relation created."
End Function



'display psi
Function getPSIFromNEA() As String

'On Error GoTo ErrorHandler
'source code from http://itpscan.info/blog/excel/VBA-XML-01.php
    Dim xmlDoc As MSXML2.DOMDocument
    Dim xEmpDetails As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim xChild As MSXML2.IXMLDOMNode
    Dim info As String
    Dim time As String
    Set xmlDoc = New MSXML2.DOMDocument
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    ' use XML string to create a DOM, on error show error message
    If Not xmlDoc.Load("http://www.haze.gov.sg/data/rss/nea_psi_3hr.xml") Then
        Debug.Print xmlDoc.parseError.ErrorCode & xmlDoc.parseError.reason
    Else
        Set xEmpDetails = xmlDoc.DocumentElement
        Set xParent = xEmpDetails.FirstChild
               
        Dim xmlNodeList As IXMLDOMNodeList
        
        Set xmlNodeList = xmlDoc.SelectNodes("//item")
        
        Dim count As Integer
        count = 0
        For Each xParent In xmlNodeList
        
            For Each xChild In xParent.ChildNodes
                'If xChild.nodeName = "pubDate" Then
                    'time = Right(xChild.text, 3)
                If xChild.nodeName = "psi" Then
                     info = Mid(xChild.text, 4, Len(xChild.text))
                End If
            Next xChild
            count = count + 1
            If count = 1 Then Exit For
        Next xParent
    

    End If
    getPSIFromNEA = info
'ErrorHandler:
 '   MsgBox "There is no internet connection.", vbInformation, "Internet Connection"

End Function


'display heavyRain
Function getHeavyRainFromNEA() As String

'On Error GoTo ErrorHandler
'source code from http://itpscan.info/blog/excel/VBA-XML-01.php
    Dim xmlDoc As MSXML2.DOMDocument
    Dim xEmpDetails As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim xChild As MSXML2.IXMLDOMNode
    Dim Col, Row As Integer
    
    Dim info As String
    info = ""
    Set xmlDoc = New MSXML2.DOMDocument
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    ' use XML string to create a DOM, on error show error message
    If Not xmlDoc.Load("http://wip.weather.gov.sg/wip/pp/rndops/web/rss/rssHeavyRain_new.xml") Then
        Debug.Print xmlDoc.parseError.ErrorCode & xmlDoc.parseError.reason
    Else
        Set xEmpDetails = xmlDoc.DocumentElement
        Set xParent = xEmpDetails.FirstChild
               
        Dim xmlNodeList As IXMLDOMNodeList
        
        Set xmlNodeList = xmlDoc.SelectNodes("//entry")
        
        Dim count As Integer
        count = 0
        For Each xParent In xmlNodeList
        
            For Each xChild In xParent.ChildNodes
                If xChild.nodeName = "summary" Then
                    info = xChild.text
                End If
            Next xChild
            count = count + 1
            If count = 1 Then Exit For
        Next xParent
    

    End If
    getHeavyRainFromNEA = info
'ErrorHandler:
'    MsgBox "There is no internet connection.", vbInformation, "Internet Connection"

End Function
