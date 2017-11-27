Attribute VB_Name = "Leased_Renewal_Utilities"
Public Sub mailMngLease()
Attribute mailMngLease.VB_ProcData.VB_Invoke_Func = " \n14"

     
Dim MailObj As Outlook.MailItem
Dim emailBody As String
Dim emailSubject As String
Dim emailTo As String
Dim emailName As String
Dim emailVIN As String
Dim emailPlate As String
Dim varRow As String

varRow = ActiveCell.Row
emailPlate = Range("D" & varRow).Value
emailVIN = Range("F" & varRow).Value
emailTo = Range("M" & varRow).Value
emailName = Range("N" & varRow).Value


'Instantiate outlook email object
Set MailObj = Outlook.Application.CreateItem(olMailItem)


If Trim(UCase(Range("K" & varRow).Value)) = "MANAGER" Then
    emailBody = emailBodyMaker("MANAGER")
    emailSubject = "Manager's Lease Registration Renewal - " & emailPlate & " - VIN " & emailVIN
ElseIf Trim(UCase(Range("K" & varRow).Value)) = "STAFF" Then
    emailBody = emailBodyMaker("STAFF")
    emailSubject = "Staff Lease Registration Renewal - " & emailPlate & " - VIN " & emailVIN
End If

    
With MailObj
  .To = emailTo
  .CC = "aeapen@altayer.com;ssoriano@altayer-motors.com;jcerna@altayer-motors.com"
  .Subject = emailSubject
  .BodyFormat = olFormatHTML
  .HTMLBody = emailBody
  .Display 'Can be .Send but prompts for user intervention before sending without 3rd party software like ClickYes
End With

 
End Sub

Private Function emailBodyMaker(emailType As String) As String

Select Case emailType

Case "MANAGER"

    emailBodyMaker = "Dear " & emailName & "," & "<br />" & "<br />" & _
    "Greetings." & "<br />" & "<br />" & _
    "This is to bring to your attention that the subject vehicle registration is due for renewal." & "<br />" & _
    "We will arrange for the vehicle test and renewal." & "<br />" & "<br />" & _
    "Kindly arrange to deliver the car at your convenience within this month to ATMC (SZR) (LandRover reception - Ms. Sim/Ms. Jil) c/o myself." & "<br />" & _
    "Time between 8:am - 12 noon; Saturday through Thursday." & "<br />" & "<br />" & _
    "Note:" & "<br />" & _
    "1)  If there is tint applied on the vehicle, kindly arrange to have the same removed from your end, as this would hinder vehicle test." & "<br />" & _
    "2)  Also, leave the original registration card in the vehicle itself (pref. Glove Compartment)." & "<br />" & "<br />" & _
    "Best Regards," & "<br />" & _
    "Philip Jacob" & "<br />" & _
    "Supervisor (Sales Administration)" & "<br />" & _
    "JLR"

Case "STAFF"

    emailBodyMaker = "Dear " & emailName & "," & "<br />" & "<br />" & _
    "Greetings." & "<br />" & "<br />" & _
    "This is to bring to your attention that the subject vehicle registration is due for renewal." & "<br />" & _
    "We will arrange for the vehicle test and renewal." & "<br />" & "<br />" & _
    "Kindly arrange to deliver the car at your convenience within this month to ATMC (SZR) (LandRover reception - Ms. Sim/Ms. Jil) c/o myself." & "<br />" & _
    "Time between 8:am - 12 noon; Saturday through Thursday." & "<br />" & "<br />" & _
    "Note:" & "<br />" & _
    "1)  Kindly leave the original registration card in the vehicle itself (pref. Glove Compartment)." & "<br />" & _
    "2)  Arrange to clear any pending traffic fines to date." & "<br />" & "<br />" & _
    "Best Regards," & "<br />" & _
    "Philip Jacob" & "<br />" & _
    "Supervisor (Sales Administration)" & "<br />" & _
    "JLR"
       
End Select

End Function
