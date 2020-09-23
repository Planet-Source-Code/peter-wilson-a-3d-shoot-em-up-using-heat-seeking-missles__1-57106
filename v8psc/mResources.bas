Attribute VB_Name = "mResources"
Option Explicit

Public Sub LoadResourceStrings(CurrentForm As Form)
    
    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim intN As Integer
    
    ' Loop through all Controls on the Form 'CurrentForm')
    For Each ctl In CurrentForm.Controls
        
        ' What 'type' is the Control? ListBox, CommandButton, etc.
        sCtlType = TypeName(ctl)
        
        Select Case sCtlType
            Case "Label", "Menu", "CommandButton", "Frame"
                If IsNumeric(ctl.Caption) = True Then ctl.Caption = LoadResString(CInt(ctl.Caption))
                
            Case "ListView"
                For Each obj In ctl.ColumnHeaders
                    If IsNumeric(obj.Text) = True Then obj.Text = LoadResString(CInt(obj.Text))
                Next
            
            Case "SSTab"
                For intN = 0 To ctl.Tabs - 1
                    If IsNumeric(ctl.TabCaption(intN)) = True Then ctl.TabCaption(intN) = LoadResString(CInt(ctl.TabCaption(intN)))
                Next intN
            
            Case "Image"
                If IsNumeric(ctl.ToolTip) = True Then ctl.ToolTip = LoadResString(CInt(ctl.ToolTip))
                If IsNumeric(ctl.Tag) = True Then ctl.Icon = LoadResPicture(CInt(ctl.Tag), vbResIcon)
                If IsNumeric(ctl.Tag) = True Then ctl.Picture = LoadResPicture(CInt(ctl.Tag), vbResIcon)
                
        End Select
    Next ctl
    
    If IsNumeric(CurrentForm.Caption) = True Then CurrentForm.Caption = LoadResString(Val(CurrentForm.Caption))
    

End Sub

Private Function LoadResString_New(p_intResourceID As Variant, resType As Integer)

End Function


