
Sub FormatPhoneNumber()
Dim oFolder As MAPIFolder
Set oFolder = Application.ActiveExplorer.CurrentFolder
If Left(UCase(oFolder.DefaultMessageClass), 11) <> "IPM.CONTACT" Then
MsgBox "Select contact folder", vbExclamation
Exit Sub
End If

Dim objOL As Outlook.Application
Dim currentExplorer As Explorer
Dim Selection As Selection

Set objOL = Outlook.Application
Set currentExplorer = objOL.ActiveExplorer
Set Selection = currentExplorer.Selection
    
Dim nCounter As Integer
nCounter = 0
On Error GoTo handle

Dim oItem
' For Each oItem In oFolder.Items
For Each oItem In Selection
    Dim oContact As ContactItem
    Set oContact = oItem
    If Not oContact Is Nothing Then
        With oContact
            .AssistantTelephoneNumber = FixFormat(.AssistantTelephoneNumber)
            .Business2TelephoneNumber = FixFormat(.Business2TelephoneNumber)
            .BusinessFaxNumber = FixFormat(.BusinessFaxNumber)
            .BusinessTelephoneNumber = FixFormat(.BusinessTelephoneNumber)
            .CallbackTelephoneNumber = FixFormat(.CallbackTelephoneNumber)
            .CarTelephoneNumber = FixFormat(.CarTelephoneNumber)
            .CompanyMainTelephoneNumber = FixFormat(.CompanyMainTelephoneNumber)
            .Home2TelephoneNumber = FixFormat(.Home2TelephoneNumber)
            .HomeFaxNumber = FixFormat(.HomeFaxNumber)
            .HomeTelephoneNumber = FixFormat(.HomeTelephoneNumber)
            .ISDNNumber = FixFormat(.ISDNNumber)
            .MobileTelephoneNumber = FixFormat(.MobileTelephoneNumber)
            .OtherFaxNumber = FixFormat(.OtherFaxNumber)
            .OtherTelephoneNumber = FixFormat(.OtherTelephoneNumber)
            .PagerNumber = FixFormat(.PagerNumber)
            .PrimaryTelephoneNumber = FixFormat(.PrimaryTelephoneNumber)
            .RadioTelephoneNumber = FixFormat(.RadioTelephoneNumber)
            .TelexNumber = FixFormat(.TelexNumber)
            .TTYTDDTelephoneNumber = FixFormat(.TTYTDDTelephoneNumber)
           
            .AssistantTelephoneNumber = RemoveChars(.AssistantTelephoneNumber)
            .Business2TelephoneNumber = RemoveChars(.Business2TelephoneNumber)
            .BusinessFaxNumber = RemoveChars(.BusinessFaxNumber)
            .BusinessTelephoneNumber = RemoveChars(.BusinessTelephoneNumber)
            .CallbackTelephoneNumber = RemoveChars(.CallbackTelephoneNumber)
            .CarTelephoneNumber = RemoveChars(.CarTelephoneNumber)
            .CompanyMainTelephoneNumber = RemoveChars(.CompanyMainTelephoneNumber)
            .Home2TelephoneNumber = RemoveChars(.Home2TelephoneNumber)
            .HomeFaxNumber = RemoveChars(.HomeFaxNumber)
            .HomeTelephoneNumber = RemoveChars(.HomeTelephoneNumber)
            .ISDNNumber = RemoveChars(.ISDNNumber)
            .MobileTelephoneNumber = RemoveChars(.MobileTelephoneNumber)
            .OtherFaxNumber = RemoveChars(.OtherFaxNumber)
            .OtherTelephoneNumber = RemoveChars(.OtherTelephoneNumber)
            .PagerNumber = RemoveChars(.PagerNumber)
            .PrimaryTelephoneNumber = RemoveChars(.PrimaryTelephoneNumber)
            .RadioTelephoneNumber = RemoveChars(.RadioTelephoneNumber)
            .TelexNumber = RemoveChars(.TelexNumber)
            .TTYTDDTelephoneNumber = RemoveChars(.TTYTDDTelephoneNumber)

            .Save
           
            nCounter = nCounter + 1
        End With
    End If
nextItem:
Next

Set currentExplorer = Nothing
Set obj = Nothing
Set Selection = Nothing

MsgBox nCounter & " contacts processed.", vbInformation
Exit Sub

handle:
    Resume nextItem
       
End Sub

Private Function RemovePrefix(strPhone As String) As String

strPhone = Trim(strPhone)
RemovePrefix = strPhone
If strPhone = "" Then Exit Function
Dim prefix As String
prefix = Left(strPhone, 1)

' Configured for US
' Enter the correct prefix here
Do While (prefix = "+" Or prefix = "91")
' if the prefix is 2 digits, change to 4;
' if 3 digits, change to 5
    strPhone = Mid(strPhone, 4)
    prefix = Left(strPhone, 4)
Loop


' RemovePrefix = strPhone

End Function

Private Function RemoveChars(strPhone As String) As String

strPhone = Trim(strPhone)
RemoveChars = strPhone
If strPhone = "" Then Exit Function
strPhone = Replace(strPhone, "(", "")
strPhone = Replace(strPhone, ")", "")
strPhone = Replace(strPhone, ".", "")
strPhone = Replace(strPhone, " ", "")
strPhone = Replace(strPhone, "-", "")
strPhone = Replace(strPhone, "+91", "+91 ")
RemoveChars = strPhone

End Function


Private Function FixFormat(strPhone As String) As String

FixFormat = strPhone
strPhone = Trim(strPhone)

If strPhone = "" Then Exit Function
'If Left(strPhone, 1) = "+" Then Exit Function
If Left(strPhone, 2) = "(+" Then Exit Function
If Left(strPhone, 2) = "+1" Then Exit Function
If Left(strPhone, 3) = "+44" Then Exit Function
If Left(strPhone, 4) = "+974" Then Exit Function
If Left(strPhone, 4) = "+971" Then Exit Function
If Left(strPhone, 3) = "+91" Then Exit Function
If Left(strPhone, 2) = "00" Then Exit Function
If Left(strPhone, 3) = "(00" Then Exit Function
If Left(strPhone, 1) = "1" Then Exit Function
If Left(strPhone, 2) = "(1" Then Exit Function

FixFormat = "+91" + strPhone

End Function
