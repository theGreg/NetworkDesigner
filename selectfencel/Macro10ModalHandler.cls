VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Macro10ModalHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IModalDialogEvents
Private Sub IModalDialogEvents_OnDialogClosed(ByVal DialogBoxName As String, ByVal DialogResult As MsdDialogBoxResult)

End Sub

Private Sub IModalDialogEvents_OnDialogOpened(ByVal DialogBoxName As String, DialogResult As MsdDialogBoxResult)

    If DialogBoxName = "Copy Model" Then

        CadInputQueue.SendCommand "DMSG CLIPBOARD PASTE"

    '   Set a variable associated with a dialog box
        SetCExpressionValue "modManInfoP->copyModelName", "1657 - SC275 - Sheet 6 of 11 Rev 1", "MODELMANAGER"

    '   Remove the following line to let the user close the dialog box.
        DialogResult = msdDialogBoxResultOK

    End If  ' Copy Model

End Sub

