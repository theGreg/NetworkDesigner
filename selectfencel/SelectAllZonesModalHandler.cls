VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectAllZonesModalHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IModalDialogEvents
Private Sub IModalDialogEvents_OnDialogClosed(ByVal DialogBoxName As String, ByVal DialogResult As MsdDialogBoxResult)

End Sub

Private Sub IModalDialogEvents_OnDialogOpened(ByVal DialogBoxName As String, DialogResult As MsdDialogBoxResult)

    If DialogBoxName = "Alert" Then

    '   Remove the following line to let the user close the dialog box.
        DialogResult = msdDialogBoxResultOK

    End If  ' Alert

End Sub

