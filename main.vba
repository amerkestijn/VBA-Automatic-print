' VBA7 is voor 64Bits systemen. Als deze niet gevonden wordt gebruiken we de functie onder "Else"
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare PtrSafe Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" _
        (ByVal pszPrinter As String) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" _
        (ByVal pszPrinter As String) As Long
#End If

Private WithEvents Items As Outlook.Items

Private Sub Application_Startup()
  Dim Ns As Outlook.NameSpace
  Dim Folder As Outlook.MAPIFolder

' objFolder not required
  Set Ns = Application.GetNamespace("MAPI")
  Set Folder = Ns.Folders("OUTLOOKFOLDER").Folders("Postvak IN")
  Set Items = Folder.Items
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
  If TypeOf Item Is Outlook.MailItem Then
    PrintAttachments Item
  End If
End Sub


Public Sub PrintAttachments(oMail As Outlook.MailItem)
  On Error Resume Next
  Dim colAtts As Outlook.Attachments
  Dim oAtt As Outlook.Attachment
  Dim sFile As String
  Dim sDirectory As String
  Dim sFileType As String
  Dim strSubject As String

' Pak de map & defineer de random number generator
  sDirectory = "PATH\TO\PRINTMAP"
  srandoname = Int((9001 - 1 + 1) * Rnd + 1)

  Set colAtts = oMail.Attachments

  If colAtts.Count Then
    For Each oAtt In colAtts

' Selecteer de laatste 4 letters van het bestand bijv. .pdf
      sFileType = LCase$(Right$(oAtt.FileName, 4))

      Select Case sFileType
' Alleen PDF bestanden selecteren
      Case ".pdf"

' Maak bestand met de random number + underscore + originele bestandsnaam
        Dim currenttime As Date
        currenttime = Now
        sFile = sDirectory & CStr(srandoname) + "_" + Format(CStr(Now), "ssmmhhddmmyyyy") + "_" + oAtt.FileName
        oAtt.SaveAsFile sFile
        
' Defineer de huidige tijd. Loop deze functie tot de huidige tijd gelijk is aan de huidige tijd + 1 seconde. Dit voorkomt vastlopers!
' Als we dit niet doen wordt niet alles geprint als er 2 mails tegelijk binnenkomen.
        Do Until currenttime + TimeValue("00:00:01") <= Now
        Loop
        
' Lanceer de Shell & open PDFtoPrinter.exe, vervolgens geven we het pad naar het zojuist opgeslagen bestand & de printer die we ervoor willen gebruiken
        Shell ("PATH\TO\PDFtoPrinter.exe " & sFile & " " & Chr(34) & "PRINTERNAAM" & Chr(34))
        
      End Select
    Next
  End If
End Sub
