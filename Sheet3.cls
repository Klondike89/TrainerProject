Private Sub cmdClear_Click()


    Dim msgValue As VbMsgBoxResult
        
    msgValue = MsgBox("Do you want to Clear the Form?", vbYesNo + vbQuestion, "Clear Form")
    
    If msgValue = vbYes Then
    
        Call TraineeReset
    
    End If
End Sub

Private Sub cmdSaveTnE_Click()

    If Validate = True Then
    
        Dim msgValue As VbMsgBoxResult
        
        msgValue = MsgBox("Do you want to save Trainee Schedule?", vbYesNo + vbQuestion, "Save")
        
            If msgValue = vbYes Then
            
                Call Save
                Call Schedule
                Call TraineeReset
            
            End If
    
    End If

End Sub

Private Sub cmdLoadTnE_Click()

    Dim msgValue As VbMsgBoxResult
        
    msgValue = MsgBox("Do you want to Update a Trainee Schedule?", vbYesNo + vbQuestion, "Update")
    
    If msgValue = vbYes Then
    
        Call Update
        Call LoadSchedule
    
    End If

End Sub

Private Sub CommandButton1_Click()
    Call LoadSchedule
End Sub
