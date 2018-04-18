Attribute VB_Name = "ConvertRsToWord"
Public strDate As String
Dim Flag As Boolean
Dim flag_6 As Boolean
Dim arr_unit(10) As String
Dim arr_hundred(10) As String
Dim arr_spl(10) As String
Public Function getAmountInWords(str_amount) As String
'On Error Resume Next
Flag = False
flag_6 = False
arr_unit(0) = ""
arr_unit(1) = "One"
arr_unit(2) = "Two"
arr_unit(3) = "Three"
arr_unit(4) = "Four"
arr_unit(5) = "Five"
arr_unit(6) = "Six"
arr_unit(7) = "Seven"
arr_unit(8) = "Eight"
arr_unit(9) = "Nine"
arr_hundred(0) = ""
arr_hundred(1) = "Ten"
arr_hundred(2) = "Twenty"
arr_hundred(3) = "Thirty"
arr_hundred(4) = "Forty"
arr_hundred(5) = "Fifty"
arr_hundred(6) = "Sixty"
arr_hundred(7) = "Seventy"
arr_hundred(8) = "Eighty"
arr_hundred(9) = "Ninety"
arr_spl(0) = ""
arr_spl(1) = "Eleven"
arr_spl(2) = "Twelve"
arr_spl(3) = "Thirteen"
arr_spl(4) = "Fourteen"
arr_spl(5) = "Fifteen"
arr_spl(6) = "Sixteen"
arr_spl(7) = "Seventeen"
arr_spl(8) = "Eighteen"
arr_spl(9) = "Nineteen"
Dim str_part, str_returnvalue, str_words_rs, str_words_ps, str_token_1
    str_part = "":    str_returnvalue = "":    str_words_rs = "":    str_words_ps = "":    str_token_1 = ""
    flag_6 = False

Dim Temp As Double
Temp = CDbl(str_amount)
If Err.Number > 0 Then
    MsgBox "Please Enter some Amount", vbInformation
    getAmountInWords = ""
    Exit Function
Else
    
    Dim Ind
    Ind = InStr(str_amount, ".")

    If (Ind <> 0) Then
        str_token_1 = Mid(str_amount, 1, Ind)
        str_token_2 = Mid(str_amount, Ind + 1, Len(str_amount))
        If Len(str_token_2) = 1 Then str_token_2 = str_token_2 & "0"
    Else
        str_token_1 = str_amount
        str_token_2 = ""
   End If
   For i = 1 To Len(str_token_1)
        str_part = Mid(str_token_1, i, Len(str_token_1))
        str_returnvalue = getWords(str_part)
        If (str_returnvalue <> "") Then str_words_rs = str_words_rs & str_returnvalue & " "
        If (Flag) Then i = i + 1
   Next
   If (str_token_2 <> "" And str_token_2 <> "00") Then
        For i = 1 To Len(str_token_2)
            str_part = Mid(str_token_2, i, Len(str_token_2))
            str_returnvalue = getWords(str_part)
            If (str_returnvalue <> "") Then str_words_ps = str_words_ps & str_returnvalue & " "
            If Flag Then i = i + 1
        Next
    End If
    If (str_token_2 <> "" And str_token_2 <> "00") Then
        getAmountInWords = str_words_rs & " And " & str_words_ps & "Cents Only"
    Else
        getAmountInWords = str_words_rs & " Only"
    End If
End If
End Function


Public Function getWords(str_part) As String
'On Error Resume Next
Dim val_1, val_2
    val_1 = ""
    val_2 = ""
    Dim str_returnvalue
    str_returnvalue = ""
    Flag = False
    Dim k
    k = Len(str_part)
    
    Select Case k
        Case 0:
        Case 1:
                Err.Clear
                val_1 = CLng(Mid(str_part, 1, 1))
                If Err.Number = 0 Then
                    str_returnvalue = arr_unit(val_1)
                End If
        Case 2:
                Err.Clear
                val_1 = CLng(Mid(str_part, 1, 1))
                val_2 = CLng(Mid(str_part, 2, 1))
                If (val_1 = 1 And val_2 <> 0) Then
                    str_returnvalue = arr_spl(val_2)
                    Flag = True
                Else
                    str_returnvalue = arr_hundred(val_1)
                End If

        Case 3:
                val_1 = CLng(Mid(str_part, 1, 1))
                If (val_1 = 0) Then
                    str_returnvalue = ""
                Else
                    str_returnvalue = arr_unit(val_1) & " " & "Hundred"
                End If
        Case 4:
                val_1 = CLng(Mid(str_part, 1, 1))
                If (val_1 = 0 And flag_6) Then
                    str_returnvalue = arr_unit(val_1)
                Else
                    str_returnvalue = arr_unit(val_1) & " " & "Thousand"
                End If
        Case 5:
                Err.Clear
                val_1 = CLng(Mid(str_part, 1, 1))
                val_2 = CLng(Mid(str_part, 2, 1))
                
                If (val_1 = 1 And val_2 <> 0) Then
                    str_returnvalue = arr_spl(val_2) & " " & "Thousand"
                    Flag = True
                Else
                    str_returnvalue = arr_hundred(val_1)
                End If

        Case 6:
                Err.Clear
                val_1 = CLng(Mid(str_part, 1, 1))
                val_2 = CLng(Mid(str_part, 2, 1))
                str_returnvalue = arr_unit(val_1) & " " & "Thousand"
                If (val_2 = 0) Then flag_6 = True
              
        Case 7:
                Err.Clear
                val_1 = CLng(Mid(str_part, 1, 1))
                val_2 = CLng(Mid(str_part, 2, 1))
                If (val_1 = 1 And val_2 <> 0) Then
                    str_returnvalue = arr_spl(val_2) & " " & "Million"
                    Flag = True
                    flag_6 = True
                Else
                    str_returnvalue = arr_hundred(val_1)
                End If
    End Select

getWords = str_returnvalue
End Function
Function NumToWords(ByVal nNumber As Double) As String
   cWords = ""
    
   If nNumber < 0 Then
      cWords = "Negative " + NumToWords(-1 * nNumber)
    
   ElseIf nNumber <> Int(nNumber) Then
      cWords = NumToWords(Int(nNumber))
      cWords = cWords + " point"
      nFrac = nNumber - Int(nNumber)
      Do
         nFrac = nFrac * 10
         nDigit = Int(nFrac)
         nFrac = nFrac - nDigit
          
         If nDigit > 0 Then
            cDigitWord = NumToWords(nDigit)
            cWords = cWords + " " + cDigitWord
         Else
            Exit Do
         End If
      Loop
    
   ElseIf nNumber < 20 Then
      Select Case nNumber
         Case 0: cWords = "Zero"
         Case 1: cWords = "One"
         Case 2: cWords = "Two"
         Case 3: cWords = "Three"
         Case 4: cWords = "Four"
         Case 5: cWords = "Five"
         Case 6: cWords = "Six"
         Case 7: cWords = "Seven"
         Case 8: cWords = "Eight"
         Case 9: cWords = "Nine"
         Case 10: cWords = "Ten"
         Case 11: cWords = "Eleven"
         Case 12: cWords = "Twelve"
         Case 13: cWords = "Thirteen"
         Case 14: cWords = "Fourteen"
         Case 15: cWords = "Fifteen"
         Case 16: cWords = "Sixteen"
         Case 17: cWords = "Seventeen"
         Case 18: cWords = "Eighteen"
         Case 19: cWords = "Ninteen"
      End Select
    
   ElseIf nNumber < 100 Then
      nTensPlace = Int(nNumber / 10)
      nOnesPlace = nNumber Mod 10
      Select Case nTensPlace * 10
         Case 20: cWords = "Twenty"
         Case 30: cWords = "Thirty"
         Case 40: cWords = "Forty"
         Case 50: cWords = "Fifty"
         Case 60: cWords = "Sixty"
         Case 70: cWords = "Seventy"
         Case 80: cWords = "Eighty"
         Case 90: cWords = "Ninty"
      End Select
      If nOnesPlace > 0 Then
         cWords = cWords + " " + NumToWords(nOnesPlace)
      End If
    
   ElseIf nNumber < 1000 Then
      nHundredsPlace = Int(nNumber / 100)
      nRest = nNumber Mod 100
      cWords = NumToWords(nHundredsPlace)
      cWords = cWords + " Hundred"
      If nRest > 0 Then
         cWords = cWords + " " + NumToWords(nRest)
      End If
    
   ElseIf nNumber < 1000000 Then
      nThousands = Int(nNumber / 1000)
      nRest = nNumber Mod 1000
      cWords = NumToWords(nThousands)
      cWords = cWords + " Thousand"
      If nRest > 0 Then
         cWords = cWords + " " + NumToWords(nRest)
      End If
    
   ElseIf nNumber < 1000000000 Then
      nMillions = Int(nNumber / 1000000)
      nRest = Int(nNumber Mod 1000000)
      cWords = NumToWords(nMillions)
      cWords = cWords + " Million"
      If nRest > 0 Then
         cWords = cWords + " " + NumToWords(nRest)
      End If
    
   ElseIf nNumber < 1000000000000# Then
      nBillions = Int(nNumber / 1000000000)
      nRest = Int(nNumber Mod 1000000000)
      cWords = NumToWords(nBillions)
      cWords = cWords + " Billion"
      If nRest > 0 Then
         cWords = cWords + " " + NumToWords(nRest)
      End If
    
   ElseIf nNumber < 1E+15 Then
      nTrillions = Int(nNumber / 1000000000000#)
      nRest = Int(nNumber Mod 1000000000000#)
      cWords = NumToWords(nTrillions)
      cWords = cWords + " Trillion"
      If nRest > 0 Then
         cWords = cWords + " " + NumToWords(nRest)
      End If
       
      ' You can follow the pattern of the Millions / Billions / Trillions
      '  if you need bigger numbers.
       
   End If
    
   NumToWords = cWords
End Function
Public Function ConverttoWord(ByVal pText As Double) As String
On Error Resume Next
Dim fin_s
fin_s = ""
c = 1
s = pText
S1 = CDbl(Mid(s, 1, 1))
ss2 = CDbl(Mid(s, 2, 1))
ss3 = CDbl(Mid(s, 3, 1))
ss4 = CDbl(Mid(s, 4, 1))
ss5 = CDbl(Mid(s, 5, 1))
ss6 = CDbl(Mid(s, 6, 1))
ss7 = CDbl(Mid(s, 7, 1))
ss8 = CDbl(Mid(s, 8, 1))
ss9 = CDbl(Mid(s, 9, 1))
ss10 = CDbl(Mid(s, 10, 1))
ss11 = CDbl(Mid(s, 11, 1))
ss12 = CDbl(Mid(s, 12, 1))
ss13 = CDbl(Mid(s, 13, 1))
ss14 = CDbl(Mid(s, 14, 1))
If (S1 = "." Or S1 = "=" Or S1 = "/") Or (ss2 = "." Or ss2 = "=" Or ss2 = "/") Or (ss3 = "." Or ss3 = "=" Or ss3 = "/") Or (ss4 = "." Or ss4 = "=" Or ss4 = "/") Or (ss5 = "." Or ss5 = "=" Or ss5 = "/") Or (ss6 = "." Or ss6 = "=" Or ss6 = "/") Or (ss7 = "." Or ss7 = "=" Or ss7 = "/") Or (ss8 = "." Or ss8 = "=" Or ss8 = "/") Or (ss9 = "." Or ss9 = "=" Or ss9 = "/") Or (ss10 = "." Or ss10 = "=" Or ss10 = "/") Or (ss11 = "." Or ss11 = "=" Or ss11 = "/") Or (ss12 = "." Or ss12 = "=" Or ss12 = "/") Or (ss13 = "." Or ss13 = "=" Or ss13 = "/") Or (ss14 = "." Or ss14 = "=" Or ss14 = "/") Then
   Exit Function
Else
   crf = Right(Str(Int(Val(s) / 1000000000#)), 2) 'billion
   hlkf = Right(Str(Int(Val(s) / 100000000)), 1) 'million
   lkf = Right(Str(Int(Val(s) / 1000000)), 2) 'million
  ' hthf = Right(str(Int(Val(s) / 100000)), 1) 'Thousand
   thf = Right(Str(Int(Val(s) / 1000)), 3) 'Thousand
   hnf = Right(Str(Int(Val(s) / 100)), 1) 'Hundred
   unf = Right(s, 2)
   If crf > 0 Then cr = spellit(crf) & " BILLION"
   If hlkf > 0 Then hlk = spellit(hlkf) & " HUNDRED"
   If lkf > 0 Then lk = spellit(lkf) & " MILLION"
   'If hthf > 0 Then hth = spellit(hthf) & " HUNDRED"
   If thf > 0 Then th = spellit(thf) & " THOUSAND"
   If thf > 0 And hnf > 0 Then th = spellit(thf) & " THOUSAND"
   If hnf > 0 Then hn = spellit(hnf) & " HUNDRED"
   If unf > 0 Then un = "AND " & spellit(unf)
   fin_s = cr & " " & hlk & " " & lk & " " & hth & " " & th & " " & hn & " " & un
   AmtFigure = fin_s
End If

Dim St, st1, st2
St = "."
st1 = Val(Total.Text)
st2 = InStr(1, st1, St, 1)

If Val(st2) < 1 Then
 If Total.Text <> "" Then
  Total.Text = st1 & "." & "00"
 End If
End If
ConverttoWord = fin_s
End Function
Public Function spellit(ByVal s As String) As String
'On Error Resume Next

Dim a(10) As String
a(0) = "ZERO"
a(1) = "ONE"
a(2) = "TWO"
a(3) = "THREE"
a(4) = "FOUR"
a(5) = "FIVE"
a(6) = "SIX"
a(7) = "SEVEN"
a(8) = "EIGHT"
a(9) = "NINE"

Dim b(10) As String
b(0) = "TEN"
b(1) = "ELEVEN"
b(2) = "TWENTY"
b(3) = "THIRTY"
b(4) = "FORTY"
b(5) = "FIFTY"
b(6) = "SIXTY"
b(7) = "SEVENTY"
b(8) = "EIGHTY"
b(9) = "NINETY"

Dim c(10) As String
c(0) = "TEN"
c(1) = "ELEVEN"
c(2) = "TWELVE"
c(3) = "THIRTEEN"
c(4) = "FOURTEEN"
c(5) = "FIFTEEN"
c(6) = "SIXTEEN"
c(7) = "SEVENTEEN"
c(8) = "EIGHTEEN"
c(9) = "NINETEEN"

s = Str(Val(s))
s = Trim(s)
plc = Len(s)
final_str = ""
up = Val(Right(s, 1))
tp = Val(Left(s, 1))

If Len(s) = 3 Then
 final_str = ConverttoWord(s)
ElseIf Len(s) = 2 Then
    If up = 0 And tp = 1 Then
        final_str = b(up)
    ElseIf up = 0 And tp > 1 Then
        final_str = b(tp)
    ElseIf tp = 1 Then
        final_str = c(up)
    Else
        final_str = b(tp) & " " & a(up)
    End If
Else
    final_str = a(up)
End If
If up = 0 And tp = 0 Then
    final_str = ""
End If
If Len(s) <= 0 Then
  AmtFigure = ""
End If
spellit = final_str
End Function


