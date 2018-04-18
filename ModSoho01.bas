Attribute VB_Name = "ModSoho01"
Option Explicit
'Public ws As Workspace
Public cn As ADODB.Connection
Public cnShape As ADODB.Connection
Public RS As ADODB.Recordset
Public rst As ADODB.Recordset
Public c As Control
Public Doc, Char As String
Public From_Menu As String
Public Type_doctor As Boolean
Public id As String
Public Check1 As Boolean
Public Username As String
Public Pass As String
'Public DB_Accounts As Database
Public Wlno_Call As Long
Public CI As Integer
Public ICno_i As String
Public doctor As String
Public Vpub_tkno
Public varAccessStr
Public strRecFlag, strId As Integer
Public BRemarks As String
Public amtpaid As Integer
Public billingID As Integer
Public billflag As String
'************************************************************
Public vPubCompanyName, vPubAddress, vPubRegno, vPubPhone, vPubFax, vPubEmail
'************************************************************

Public pubUserName, pubUserDept
Public strRetrieve As String


Public Function IsEMailAddress(ByVal sEmail As String, Optional ByRef sReason As String) As Boolean

        

   


    Dim nCharacter As Integer

    Dim sBuffer As String



    sEmail = Trim(sEmail)



    If Len(sEmail) < 8 Then

        IsEMailAddress = False

        sReason = "Too short"

        Exit Function

    End If





    If InStr(sEmail, "@") = 0 Then

        IsEMailAddress = False

        sReason = "Missing the @"

        Exit Function

    End If





    If InStr(InStr(sEmail, "@") + 1, sEmail, "@") <> 0 Then

        IsEMailAddress = False

        sReason = "Too many @"

        Exit Function

    End If





    If InStr(sEmail, ".") = 0 Then

        IsEMailAddress = False

        sReason = "Missing the period"

        Exit Function

    End If



    If InStr(sEmail, "@") = 1 Or InStr(sEmail, "@") = Len(sEmail) Or InStr(sEmail, ".") = 1 Or InStr(sEmail, ".") = Len(sEmail) Then

        IsEMailAddress = False

        sReason = "Invalid format"

    Exit Function



End If





For nCharacter = 1 To Len(sEmail)

    sBuffer = Mid$(sEmail, nCharacter, 1)

    If Not (LCase(sBuffer) Like "[a-z]" Or sBuffer = "@" Or sBuffer = "." Or sBuffer = "-" Or sBuffer = "_" Or IsNumeric(sBuffer)) Then: IsEMailAddress = False: sReason = "Invalid character": Exit Function

Next nCharacter



nCharacter = 0



'On Error Resume Next



sBuffer = Right(sEmail, 4)

If InStr(sBuffer, ".") = 0 Then GoTo TooLong:

If Left(sBuffer, 1) = "." Then sBuffer = Right(sBuffer, 3)

If Left(Right(sBuffer, 3), 1) = "." Then sBuffer = Right(sBuffer, 2)

If Left(Right(sBuffer, 2), 1) = "." Then sBuffer = Right(sBuffer, 1)





If Len(sBuffer) < 2 Then

    IsEMailAddress = False

    sReason = "Suffix too short"

    Exit Function

End If



TooLong:



If Len(sBuffer) > 3 Then

    IsEMailAddress = False

    sReason = "Suffix too long"

    Exit Function

End If



sReason = Empty

IsEMailAddress = True



End Function

Public Sub UPPER_CASE(KeyAscii As Integer)
If KeyAscii = "13" Then SendKeys "{TAB}"
Char = Chr(KeyAscii)
KeyAscii = Asc(UCase(Char))
End Sub
Public Sub connect()
'On Error GoTo Xit
Set cn = New ADODB.Connection
With cn
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.ConnectionString = vPubConnectString
cn.CursorLocation = adUseClient
cn.Open
End With
Set cnShape = New ADODB.Connection
With cnShape
cnShape.Provider = "Microsoft.jet.oledb.4.0"
cnShape.ConnectionString = vPubConnectString
'cnShape.
cnShape.CursorLocation = adUseClient
cnShape.Open
End With

Exit Sub
Xit:
Dim fso As New FileSystemObject
fso.DeleteFile App.Path & "\serversetup.ini"
MsgBox "Database Path is invalid. Invoke the application again", vbCritical

End
Exit Sub
End Sub

Public Sub CLEAR_TEXT(a As Form)
'Dim c1 As Control

For Each c In a
If TypeOf c Is TextBox Then
    c.Text = ""
   
End If
If TypeOf c Is CheckBox Then
    c.Value = 0
End If
Next

End Sub

Public Sub ERROR_HANDLING()
 Dim Msg
If Err.Number = "3024" Then
    Msg = MsgBox("File not found for the selected Company" + vbCrLf + "Enter through the ADMINISTRATOR login and Restore the file", vbOKOnly + vbInformation, "Clinic")
ElseIf Err.Number = "94" Then
ElseIf Err.Number = "5" Then
    Err.Clear
ElseIf Err.Number = "3200" Then
    Msg = MsgBox("This Record Can't Be Deleted.Because It Has Some Transactions", vbOKOnly + vbCritical, "CLinic")

ElseIf Err.Number = "53" Then
    Msg = MsgBox("File Not Found In the Given Drive", vbOKOnly + vbCritical, "Clinic")
ElseIf Err.Number = "70" Then
    Msg = MsgBox("Restore Not Done Properly. May Be The File is Already Existing", vbCritical, "Clinic")

ElseIf Err.Number = "3022" Then
    Msg = MsgBox("You Have Tried To Insert Duplicate Record", vbOKOnly + vbInformation, "Clinic")

ElseIf Err.Number = "71" Then
    MsgBox Err.Description, vbOKOnly + vbCritical, "Clinic"
End If
End Sub

Public Sub numericonly(KeyAscii As Integer)
If Not (Chr(KeyAscii) Like "[0-9.]") Then
Beep
KeyAscii = 0
End If
End Sub
Public Function validate(TempKeyAscii As Integer, Optional TempInt As Integer = 0) As Integer

Select Case TempInt
Case 0: ' FOR ACCEPT ONLY A-Z AND a-z  and  backspace and space AND CONVERT UPPER CASE

If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) Then
 
 validate = 0
 Else
validate = Asc((UCase(Chr(TempKeyAscii))))
 
 End If

Case 1: ' FOR ACCEPT ONLY A-Z and a-z  and  backspace and  dot and spase and convert upper case
If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) Then
 
 validate = 0
 Else
validate = Asc((UCase(Chr(TempKeyAscii))))
End If

Case 2: ' FOR ACCEPT ONLY A-Z and a-z  and  backspace and  dot and spase and convert uppercase
If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) And (Not TempKeyAscii Like 44) Then
 
 validate = 0
 Else
validate = Asc((UCase(Chr(TempKeyAscii))))
End If
Case 3: ' FOR ACCEPT ONLY A-Z and a-z  and  backspace and  convert uppercase
If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not TempKeyAscii Like 8) Then
 
 validate = 0
 Else
validate = Asc((UCase(Chr(TempKeyAscii))))
 
 End If
 Case 4: ' FOR ACCEPT ONLY 0-9  and  backspace and space and dot and coma
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) And (Not TempKeyAscii Like 44) Then
 
 
 validate = 0
 Else
validate = TempKeyAscii
 
 End If
Case 5: ' FOR ACCEPT ONLY 0-9  and  backspace
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) Then
 
 
 validate = 0
 Else
validate = TempKeyAscii
 
 End If
 Case 6: ' FOR ACCEPT ONLY 0-9  and  backspace and minus(-)
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 45) Then
 
 
 validate = 0
 Else
validate = TempKeyAscii
 
 End If
  Case 7: ' FOR ACCEPT ONLY 0-9  and  backspace AND COLON(:)
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 58) And (Not TempKeyAscii Like 8) Then
 
 
 validate = 0
 Else
validate = TempKeyAscii
 
 End If
 '59, 58, 34, 39, 44, 60, 46, 62, 47, 61, 43, 63, 123, 125, 92, 124, 96, 126, 33, 37, 94, 38, 95
  Case 8: 'FOR MOMILE NO
If (TempKeyAscii Like 59) Or (TempKeyAscii Like 58) Or (TempKeyAscii Like 34) Or (TempKeyAscii Like 39) _
Or (TempKeyAscii Like 44) Or (TempKeyAscii Like 60) Or (TempKeyAscii Like 46) Or (TempKeyAscii Like 62) _
Or (TempKeyAscii Like 47) Or (TempKeyAscii Like 61) Or (TempKeyAscii Like 43) Or (TempKeyAscii Like 63) _
Or (TempKeyAscii Like 123) Or (TempKeyAscii Like 125) Or (TempKeyAscii Like 92) Or (TempKeyAscii Like 124) _
Or (TempKeyAscii Like 96) Or (TempKeyAscii Like 126) Or (TempKeyAscii Like 33) Or (TempKeyAscii Like 37) _
Or (TempKeyAscii Like 94) Or (TempKeyAscii Like 38) Or (TempKeyAscii Like 95) Then
validate = 0
 Else
validate = TempKeyAscii
End If

 Case 9: ' FOR ACCEPT ONLY 0-9  and  backspace and minus(-) and opening small brakit(  and closing small brakit )
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 45) And (Not TempKeyAscii Like 40) And (Not TempKeyAscii Like 41) Then
 
 
 validate = 0
 Else
validate = TempKeyAscii
 
 End If
  Case 10: 'not accept ' "
If (TempKeyAscii Like 34) Or (TempKeyAscii Like 39) Then
validate = 0
 Else
validate = TempKeyAscii
End If
Case 11: ' FOR ACCEPT ONLY 0-9  and  backspace and space and dot and
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) Then
 
 
 validate = 0
 Else
validate = TempKeyAscii
 
 End If
Case 12: ' FOR ACCEPT ONLY A-Z AND a-z  and 0-9 and  backspace and AND CONVERT UPPER CASE

If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) Then
 
 validate = 0
 Else
validate = Asc((UCase(Chr(TempKeyAscii))))
 
 End If


Case 13: ' FOR ACCEPT ONLY A-Z AND a-z  and 0-9 and  backspace and AND SPACE AND DOT AND minus(-) and opening small brakit(  and closing small brakit )and coma CONVERT UPPER CASE

If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) And (Not TempKeyAscii Like 45) And (Not TempKeyAscii Like 40) And (Not TempKeyAscii Like 41) And (Not TempKeyAscii Like 44) Then
 
 validate = 0
 Else
validate = Asc((UCase(Chr(TempKeyAscii))))
 
 End If
Case 14: ' FOR ACCEPT ONLY A-Z AND a-z  and 0-9 and  backspace and AND SPACE AND DOT AND minus(-) and / and opening small brakit(  and closing small brakit )  and   CONVERT UPPER CASE

If (Not Chr(TempKeyAscii) Like "[A-Za-z]") And (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) And (Not TempKeyAscii Like 46) And (Not TempKeyAscii Like 45) And (Not TempKeyAscii Like 40) And (Not TempKeyAscii Like 41) And (Not TempKeyAscii Like 47) And (Not TempKeyAscii Like 44) Then
 
 validate = 0
 Else
validate = Asc((UCase(Chr(TempKeyAscii))))
 
 End If
Case 15: ' FOR ACCEPT ONLY 0-9  and  backspace and minus(-) and /
If (Not Chr(TempKeyAscii) Like "[0-9]") And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 45) And (Not TempKeyAscii Like 47) Then
 
 
 validate = 0
 Else
validate = TempKeyAscii
 
 End If
Case 16: 'for blood group
If Not TempKeyAscii Like 45 And Not TempKeyAscii Like 43 And Not Chr(TempKeyAscii) Like "[A-Za-z]" And (Not TempKeyAscii Like 8) And (Not TempKeyAscii Like 32) Then
validate = 0
Else
validate = TempKeyAscii
End If
End Select

End Function

Public Function GF_AllowCapitalOnly(L_IntKeyAscii As Integer) As Integer
    If L_IntKeyAscii > 47 And L_IntKeyAscii < 58 Then
        GF_AllowCapitalOnly = 0
    ElseIf L_IntKeyAscii > 96 And L_IntKeyAscii < 123 Then
        GF_AllowCapitalOnly = L_IntKeyAscii - 32
    Else
        GF_AllowCapitalOnly = L_IntKeyAscii
    End If
End Function
Public Function GF_ProcessSingleQuote(L_StrText As String) As String

'Funtion returing a string representing L_StrText after single quotes in the concerned
'string is replaced by two single quotes.
Dim L_IntPosition As Integer        'Integer to store the start position of search for the string to be found.
Dim L_StrQuote As String            'String to store the search string.
'On Error GoTo errorhandler
    L_StrQuote = "'"                'Assign search string as single quote.
    Do
        'Store position of occurance of single quotes in the L_StrText.
        L_IntPosition = InStr(L_IntPosition + 1, L_StrText, L_StrQuote)
        'If no occurance found then exit loop.
        If L_IntPosition = 0 Then
            Exit Do
        End If
        'Replace single quote with two single quote.
        L_StrText = Mid(L_StrText, 1, L_IntPosition - 1) & "'" & Mid(L_StrText, L_IntPosition)
        'Increment the Start Position for search by 1 of the last occurance.
        L_IntPosition = L_IntPosition + 1
    Loop
    GF_ProcessSingleQuote = L_StrText       'Return the replaced string.
Exit Function
errorhandler:
    MsgBox " Error Encountered while Processing Single Quote in text string.." & Chr(13) _
            & " Error Number: " & Err.Number & " Error Description: " & Err.Description, vbInformation, " Process Single Quote"
End Function
Public Sub GS_FillComboBox(LcmbCombo As ComboBox, LRsRecSet As Recordset, Optional LIntFieldNum As Integer)
'Sub to Fill Combo control represented in LcmbCombo parameter using the
'LRsRecset recordset parameter with the field value specified in the LIntFieldNum field number
'if specified else with the first field in the recordset.
  'On Error GoTo errorhandler
  'Clearing the combo content.
    LcmbCombo.Clear
  'Adding field values to the Combo control.
    While Not LRsRecSet.EOF
        If Not IsNull(LRsRecSet.Fields(LIntFieldNum)) Then
            LcmbCombo.AddItem LRsRecSet.Fields(LIntFieldNum)
        End If
        LRsRecSet.MoveNext      'Moving to next record.
            Wend
    LRsRecSet.MoveFirst
Exit Sub
errorhandler:
        MsgBox "No Records Retrived !" & Chr(13) & Chr(13) _
                & "Error Number - " & Err.Number & Chr(13) _
                & "Error Description - " & Err.Description, vbInformation, "Error No Record"
End Sub

Public Function GF_IncrementPriKey(L_StruniqId As String) As String
'Funtion returning a string representing the Incremented primary key value.
Dim L_strPriKey As String     'To Store String part of Primary Key (including 0 (Zero))
Dim L_lngPriKey As Long       'To Store Integr part of Primary Key
Dim L_IntCounter As Integer   'To count the number of iterations
'On Error GoTo errorhandler
    For L_IntCounter = 1 To Len(L_StruniqId)    'Loop to access each character of the L_StruniqId at a time.
    'If the retrieved character is a number, retrieve the complete string from this position
    'and store the same to L_lngPriKey long variable and exit the For loop.
        If Mid(L_StruniqId, L_IntCounter, 1) >= Chr(49) _
           And Mid(L_StruniqId, L_IntCounter, 1) <= Chr(57) Then
                L_lngPriKey = CLng(Mid(L_StruniqId, L_IntCounter))
                If L_IntCounter > 1 Then
                    If Len(CStr(L_lngPriKey + 1)) > Len(CStr(L_lngPriKey)) _
                        And Mid(L_StruniqId, L_IntCounter - 1, 1) = Chr(48) Then
                        L_strPriKey = Mid(L_strPriKey, 1, Len(L_strPriKey) - 1)
                    End If
                End If
                Exit For
        Else
                L_strPriKey = L_strPriKey & Mid(L_StruniqId, L_IntCounter, 1)
        End If
    Next L_IntCounter
    'Increment the Primary numeric portion value by one.
    
    L_lngPriKey = L_lngPriKey + 1
    'Form the new primary key value, if consisted alpha portion concatenate the
    'numeric portion to the alpha string. Else assign only the numeric portion
    'after converting to string.
    If L_strPriKey <> "0" Then
        L_strPriKey = L_strPriKey & CStr(L_lngPriKey)
    Else
        L_strPriKey = CStr(L_lngPriKey)
    End If
'Return the Primary Key string.
    GF_IncrementPriKey = L_strPriKey
Exit Function
errorhandler:
    MsgBox "Error Encountered while Incrementing Primary Key value.." & Chr(13) _
            & "Error Number: " & Err.Number & "Error Description: " & Err.Description, vbInformation, "Incrementing Primary Key"
End Function
Public Function GF_AllowRealNumberOnly(lKeyAscii As Integer) As Integer
'On Error GoTo errorhandler
'Sub allowing Real Number entry only by turning keyascii to value zero.
    If Not (lKeyAscii >= 48 And lKeyAscii <= 57) And lKeyAscii <> 8 And lKeyAscii <> 46 Then
        lKeyAscii = 0
    End If
    If lKeyAscii = 46 Then
        lKeyAscii = 0
    End If
GF_AllowRealNumberOnly = lKeyAscii
Exit Function
errorhandler:
    MsgBox "Error Encountered while Allowing real number entry only.." & Chr(13) _
            & "Error Number: " & Err.Number & "Error Description: " & Err.Description, vbInformation, "Real Number Only"
End Function


Public Function TxtAcceptString(obj As Object, ByVal KeyAscii As Integer) As Integer
    If KeyAscii = 8 Then
            If Len(obj.Text) >= 1 Then
                    TxtAcceptString = KeyAscii
                    Exit Function
            Else
                    TxtAcceptString = 0
            End If
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}": TxtAcceptString = 0: Exit Function
    If KeyAscii = 34 Or KeyAscii = 39 Then TxtAcceptString = 0: Exit Function
 
    TxtAcceptString = KeyAscii
End Function

Public Sub SearchString(ByVal StrQry, ByVal obj As Object)
'Dim rs2 As New ADODB.Recordset
'rs2.Open StrQry, cn
'Dialog.List1.Clear
'While Not rs2.EOF
'If rs2.Fields.Count = 2 Then
'    Dialog.List1.AddItem rs2(1) & "--" & rs2(0)
'Else
'    Dialog.List1.AddItem rs2(0)
'End If
'rs2.MoveNext
'Wend
'rs2.Close
'
'Set Dialog.varfrm = obj
'Dialog.Show
End Sub

Public Function autogen(ByVal qry, ByVal paddstr, con As Object) As String
On Error GoTo Xit
Dim auto1, auto2, i
  Dim rs1 As New ADODB.Recordset
    Dim a() As Variant
   
    rs1.CursorLocation = adUseClient
    rs1.Open qry, cn, 3, 2, 1
    If rs1.RecordCount = 0 Then
        autogen = paddstr & "1"
    Else
        rs1.MoveFirst
        auto1 = 1
        ReDim a(rs1.RecordCount)
        For i = 1 To rs1.RecordCount
          a(i) = Replace(UCase(rs1(0)), UCase(paddstr), "", 1, Len(rs1(0)))
          If CDbl(a(i)) > CDbl(auto1) Then auto1 = CDbl(a(i))
          rs1.MoveNext
        Next i
        auto2 = CInt(auto1) + 1
        autogen = paddstr & auto2
    End If
    rs1.Close
Xit:
End Function



Public Sub flxAcceptMoney(obj As Object, KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 8 Then obj.Text = Mid(obj.Text, 1, Len(obj.Text) - 1): Exit Sub
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        obj.Text = obj.Text & Chr(KeyAscii)
        Exit Sub
    End If
    If KeyAscii = 46 Then
        If InStr(obj.Text, ".") Then
            Exit Sub
        Else
            obj.Text = obj.Text & Chr(KeyAscii)
        End If
    End If
End Sub



Public Function GetCode(ByVal qry) As String
  
    Dim RS As New ADODB.Recordset
    RS.Open qry, cn, adOpenKeyset, adLockReadOnly
    If Not RS.EOF Then GetCode = RS(0)

    RS.Close
   
End Function
Public Sub StockUpdate(ByVal IC, ByVal Iq)
    cn.Execute "Update Itemmaster set stock = stock +" & Iq & " Where itemdescription='" & Trim(IC) & "'"
End Sub

