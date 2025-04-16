Sub CombineEmailAddresses()
    Dim fixedEmails As String
    Dim additionalEmails As String
    Dim combinedEmails As String
    
    ' Define fixed email addresses
    fixedEmails = "john.doe@example.com; sarah.connor@domain.test; michael.smith@company.org; anna.brown@samplemail.net; mark.jones@fakemail.co; lisa.taylor@nowhere.com"
    
    ' Get additional emails from cell A2 in the "EMAIL" sheet
    additionalEmails = Sheets("EMAIL").Range("A2").Value
    
    ' Combine fixed and additional emails
    If additionalEmails <> "" Then
        combinedEmails = fixedEmails & "; " & additionalEmails
    Else
        combinedEmails = fixedEmails
    End If
    
    ' Output the combined result to cell A3 on the "EMAIL" sheet
    Sheets("EMAIL").Range("A3").Value = combinedEmails
End Sub

