Attribute VB_Name = "import_wegli"
Option Explicit

Sub importData()

    Dim a, fso, dic, arr, res() As String
    Dim i&, j%, str$
    
    'At first you need to download your wegli data and save it as 'notices.txt' to your download folder
    'https://www.weg.li/notices/dump.json

    'Open the notices.txt file as textstream and read the textstream to a string
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set a = fso.Opentextfile(VBA.Environ("Userprofile") & "\Downloads\notices.txt", 1)

    str = a.readall
    
    'Terminate the objects as they are no longer needed
    Set a = Nothing
    Set fso = Nothing

    'Create a dictionary in order to reference to the json's properties by index
    Set dic = CreateObject("Scripting.Dictionary")   
    With dic
        .Add 0, "title"
        .Add 1, "status"
        .Add 2, "street"
        .Add 3, "city"
        .Add 4, "zip"
        .Add 5, "latitude"
        .Add 6, "longitude"
        .Add 7, "registration"
        .Add 8, "color"
        .Add 9, "brand"
        .Add 10, "charge"
        .Add 11, "date"
        .Add 12, "duration"
        .Add 13, "severity"
        .Add 14, "photos"
        .Add 15, "created_at"
        .Add 16, "updated_at"
        .Add 17, "sent_at"
        .Add 18, "vehicle_empty"
        .Add 19, "hazard_lights"
        .Add 20, "expired_tuv"
        .Add 21, "expired_eco"
    End With

    'Populate an array by splitting the string on the substring below.
    'I have chosen to split it on the token property because that introduces
    'a new notice in the source string.
    arr = Split(str, "{""token"":")
    

    'Redimension the result array as needed - the second dimension's upper
    'boundary is to be set in concordance with the number of properties
    'defined in the dictionary.
    ReDim res(0 To UBound(arr), 0 To dic.Count - 1)       
    For i = 0 To UBound(res, 2)
        res(0, i) = dic.Item(i)
    Next i
    
    'Populate the result array by looping through both dimensions.
    'In general, the substring to assign to the respective index is
    'found by using the position in the string of the current property
    'and the position of the next property.
    'Special cases for j=0 and j=21 (lower and upper boundary in second dimension)
    'have been defined. Furthermore, there has been defined a special case for j=14,
    'which contains the photos property. As for my needs the photo filenames are not
    'relevant, in that case the loop jumps to next j.
    'Every position in the source array arr is processed by the cleanseString function,
    'which replaces umlauts and - more important - strips the value of the photos property
    'to [{}]
    For i = 1 To UBound(arr)
        arr(i) = cleanseString(arr(i))
        For j = 0 To UBound(res, 2)
            Select Case j
                Case Is = 0
                    res(i, j) = Mid(arr(i), 1, InStr(arr(i), dic.Item(j + 1)) - 3)
                Case Is = 14
                    GoTo next_j
                Case Is = 21
                    res(i, j) = Mid(arr(i), InStr(arr(i), dic.Item(j)) + Len(dic.Item(j)) + 2, InStr(arr(i), "}") - InStr(arr(i), dic.Item(j)) - Len(dic.Item(j)) - 2)
                Case Else
                    res(i, j) = Mid(arr(i), InStr(arr(i), dic.Item(j)) + Len(dic.Item(j)) + 2, InStr(arr(i), dic.Item(j + 1)) - InStr(arr(i), dic.Item(j)) - Len(dic.Item(j)) - 4)
            End Select
            
next_j:
        Next j
    Next i

    'Insert the data from the result array to a new worksheet
    ThisWorkbook.Worksheets.Add.Range("A1").Resize(UBound(res, 1) + 1, UBound(res, 2) + 1) = res
    
    Set dic = Nothing
    Set arr = Nothing

End Sub



Function cleanseString(ByVal str As String) As String

    Dim s$, t$, n&, k&
    s = """photos"":[{"
    t = "}]"
    
    n = InStr(str, s) - 2 'starting point of photos block
    k = InStr(n, str, t) + 1 'end point of photos block
    
    cleanseString = Left(str, n + Len(s)) & Mid(str, k)
    
    cleanseString = Replace(cleanseString, "ÃŸ", "ß")
    cleanseString = Replace(cleanseString, "Ã¤", "ä")
    cleanseString = Replace(cleanseString, "Ã„", "Ä")
    cleanseString = Replace(cleanseString, "Ã¶", "ö")
    cleanseString = Replace(cleanseString, "Ã–", "Ö")
    cleanseString = Replace(cleanseString, "Ã¼", "ü")
    cleanseString = Replace(cleanseString, "Ãœ", "Ü")
    cleanseString = Replace(cleanseString, "Å ", "S")
    

End Function



