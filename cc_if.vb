Option Compare Text

Private Function MatchAddress(ByRef Item As Object, strMatchPattern As String) As Boolean
    
    MatchAddress = False
    
    For Each curRecipient In Item.Recipients
    
        If MatchAddressEntry(curRecipient.AddressEntry, strMatchPattern) Then
            MatchAddress = True
            Exit For
        End If
        
    Next
        
End Function

Private Function MatchAddressEntry(ByVal entry As AddressEntry, strMatchPattern As String)

    MatchAddressEntry = False
    
    Select Case entry.Type
    
        Case "SMTP"
            MatchAddressEntry = MatchAddressString(entry.Address, strMatchPattern)
        
        Case Else
            If Not entry.Members Is Nothing Then
                For Each distroListEntry In entry.Members
                    If MatchAddressEntry(distroListEntry, strMatchPattern) Then
                        MatchAddressEntry = True
                        Exit For
                    End If
                Next
            End If
    End Select
    
End Function

Private Function MatchAddressString(strAddress As String, strMatchPattern As String)
    MatchAddressString = strAddress Like strMatchPattern
End Function


Private Function AddBccRecipient(ByRef Item As Object, strBcc As String) As Boolean
    
    AddBccRecipient = True
    
    Dim objRecip As recipient
    Dim strMsg As String
    Dim res As Integer
    
    Set objRecip = Item.Recipients.Add(strBcc)
    objRecip.Type = olBCC
    
    If Not objRecip.Resolve Then
        strMsg = "Could not resolve the Bcc recipient. " & _
        "Do you want still to send the message?"
        res = MsgBox(strMsg, vbYesNo + vbDefaultButton1, _
        "Could Not Resolve Bcc Recipient")
        If res = vbNo Then
            AddBccRecipient = False
        End If
    End If
        
    Set objRecip = Nothing

End Function

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
   
    Dim strBcc As String
    Dim strFilter As String
    Dim addressMatches As Boolean
    
    strBcc = "davidbashirov@gmail.com"
    strFilter = "*@sitesell.com"
        
    On Error Resume Next
    
    addressMatches = MatchAddress(Item, strFilter)
    
    If Not addressMatches Then
        Cancel = Not AddBccRecipient(Item, strBcc)
    End If
    
End Sub