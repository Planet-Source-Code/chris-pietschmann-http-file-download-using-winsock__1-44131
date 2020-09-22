Attribute VB_Name = "modGetHTTPInfo"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Written by Chris Pietschmann, MCP
'http://PietschSoft.com
'
'You can find more good code examples and articles at my website.
'
'This was written off of the HTTP 1.1 Specification
'
'If you use this file please either leave this header in it,
' Or please just give me credit.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'You can use the GetHTTPInfo to parse info out of the server response.
'You can use the GetHTTPStatus_Code_Text to return a short english desctiption of the Status Code specified.
'The Enum HTTP_Data is used in the GetHTTPInfo to tell what info to parse out of the server response.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Enum HTTP_Data
    HTTP_Data_Accept_Ranges             ''"Accept-Ranges:"
    HTTP_Data_Cache_Control             ''"Cache-Control:"
    HTTP_Data_Connection                ''"Connection:"
    HTTP_Data_Content_Type              ''"Content-Type:"
    HTTP_Data_Content_Length            ''"Content-Length:"
    HTTP_Data_Date                      ''"Date:"
    HTTP_Data_Expires                   ''"Expires:"
    HTTP_Data_HTTP_Version              ''Get the version of HTTP used
    HTTP_Data_Location                  ''"Location:"
    HTTP_Data_MicrosoftOfficeWebServer  ''"MicrosoftOfficeWebServer:"
    HTTP_Data_MIME_Version              ''"MIME-Version:"
    HTTP_Data_P3P                       ''"P3P:"
    HTTP_Data_RemoteHost                ''"RemoteHost:"
    HTTP_Data_Set_Cookie                ''"Set-Cookie:"
    HTTP_Data_Server                    ''"Server:"
    HTTP_Data_Status_Code               ''Get the Status Code in the header
    HTTP_Data_Data                      ''Get the Data from the server response
End Enum
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'this Enum is used exclusively inside the GetHTTPInfo function
Private Enum HTTP_Data_Type
    HTTP_Data_Type_Other
    HTTP_Data_Type_HTTP_Version
    HTTP_Data_Type_Status_Code
    HTTP_Data_Type_Data
End Enum
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetHTTPInfo(strData As String, HTTP_Data_Name As HTTP_Data) As String
'You can use the GetHTTPInfo to parse info out of the server response.
    Dim strName As String
    Dim strValue As String, intName As Integer, intCRLF As Integer
    Dim Data_Type As HTTP_Data_Type
    
    Select Case HTTP_Data_Name
        Case 0:
            strName = "Accept-Ranges:"
            Data_Type = HTTP_Data_Type_Other
        Case 1:
            strName = "Cache-Control:"
            Data_Type = HTTP_Data_Type_Other
        Case 2:
            strName = "Connection:"
            Data_Type = HTTP_Data_Type_Other
        Case 3:
            strName = "Content-Type:"
            Data_Type = HTTP_Data_Type_Other
        Case 4:
            strName = "Content-Length:"
            Data_Type = HTTP_Data_Type_Other
        Case 5:
            strName = "Date:"
            Data_Type = HTTP_Data_Type_Other
        Case 6:
            strName = "Expires:"
            Data_Type = HTTP_Data_Type_Other
        Case 7: ''get the HTTP Version used
            Data_Type = HTTP_Data_Type_HTTP_Version
        Case 8:
            strName = "Location:"
            Data_Type = HTTP_Data_Type_Other
        Case 9:
            strName = "MicrosoftOfficeWebServer:"
            Data_Type = HTTP_Data_Type_Other
        Case 10:
            strName = "MIME-Version:"
            Data_Type = HTTP_Data_Type_Other
        Case 11:
            strName = "P3P:"
            Data_Type = HTTP_Data_Type_Other
        Case 12:
            strName = "RemoteHost:"
            Data_Type = HTTP_Data_Type_Other
        Case 13:
            strName = "Set-Cookie:"
            Data_Type = HTTP_Data_Type_Other
        Case 14:
            strName = "Server:"
            Data_Type = HTTP_Data_Type_Other
        Case 15: ''get the HTTP Status Code in the header
            Data_Type = HTTP_Data_Type_Status_Code
        Case 16: ''get the Data from the server response
            Data_Type = HTTP_Data_Type_Data

    End Select
    
    If Data_Type = HTTP_Data_Type_Other Then
        ''Get the desired value out of the Header
        intName = InStr(1, strData, strName, vbTextCompare) + (Len(strName))
        intCRLF = InStr(intName, strData, vbCrLf, vbTextCompare)
        If intName > 0 Then
            strValue = Mid(strData, intName, (intCRLF - intName))
        End If
    
        ''make sure there are not leading or trailing spaces
        Do Until Right(strValue, 1) <> Chr(32) And Left(strValue, 1) <> Chr(32)
            If Right(strValue, 1) = Chr(32) Then strValue = Left(strValue, Len(strValue) - 1)
            If Left(strValue, 1) = Chr(32) Then strValue = Right(strValue, Len(strValue) - 1)
        Loop
    
    ElseIf Data_Type = HTTP_Data_Type_HTTP_Version Then
        ''Get the HTTP Version Used
        strValue = Right(Left(strData, 8), 3)
    
    ElseIf Data_Type = HTTP_Data_Type_Status_Code Then
        ''get the HTTP Status Code in the header
        strValue = Right((Left(strData, 12)), 3)
    
    ElseIf Data_Type = HTTP_Data_Type_Data Then
        ''get the Data out of the Server Response
        intName = InStr(1, strData, (vbCrLf & vbCrLf), vbTextCompare)
        strValue = Right(strData, GetHTTPInfo(strData, HTTP_Data_Content_Length))
    End If
    
    'return the value found
    GetHTTPInfo = strValue
End Function
