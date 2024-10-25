Attribute VB_Name = "exemple"
Function FormatTextWithDict(filePath As String, data As Object) As String
    'Opens a text file with placeholders and replaces them with the values from a dictionary
    'Args:
        '- filePath: path of the file to be opened and formatted
        '- data: dictionary with the values to replace the placeholders
    'Raises:
        '- 13: Type Mismatch, the 'data' argument is not a dictionary
        '- 75: Path/File access error, it was not possible to open/read the file
    'Returns:
        '- Formated text file as string

    Dim key As Variant
    Dim fileNumber As Integer
    Dim fileContent As String

    'Checks if the data argument is a dictionary
    If Not TypeName(data) = "Dictionary" Then
        Err.Raise Number:=13, Description:="Type Mismatch: The 'data' argument is not a dictionary."
    End If

    'Informs the space in memory for the file
    fileNumber = FreeFile

    'Reads the file content
    On Error GoTo ReadFileError
    Open filePath For Input As fileNumber
        'LOF lê o arquivo em bytes e aloca na variável
        fileContent = Input(LOF(fileNumber), fileNumber)
    Close fileNumber

    'Formats the text placeholders with the dictionary values
    For Each key In data.Keys
        'Substituir cada placeholder pelo valor correspondente
        fileContent = Replace(fileContent, "{{" & key & "}}", data(key))
    Next key

    'Returns the formatted text
    FormatTextWithDict = fileContent
    Exit Function

'Error treatment
ReadFileError:
    Err.Raise Number:=75, Description:="Path/File access error: It was not possible to open/read the file"
End Function
