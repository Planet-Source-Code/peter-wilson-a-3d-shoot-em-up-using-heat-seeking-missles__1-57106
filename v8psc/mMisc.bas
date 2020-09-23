Attribute VB_Name = "mMisc"
Option Explicit

Public Sub Debug_PrintMatrix(Canvas As PictureBox, m1 As mdrMatrix4)
    
    With m1
        
        If (Canvas Is Nothing) = False Then
            
            Canvas.ForeColor = RGB(0, 0, 0)
            Canvas.CurrentX = 0
            Canvas.CurrentY = 0
            
            Canvas.Font = "Small Fonts"
            Canvas.FontSize = 7
            
            Canvas.Print Format(.rc11, "0.0000"), Format(.rc12, "0.0000"), Format(.rc13, "0.0000"), Format(.rc14, "0.0000")
            Canvas.Print Format(.rc21, "0.0000"), Format(.rc22, "0.0000"), Format(.rc23, "0.0000"), Format(.rc24, "0.0000")
            Canvas.Print Format(.rc31, "0.0000"), Format(.rc32, "0.0000"), Format(.rc33, "0.0000"), Format(.rc34, "0.0000")
            Canvas.Print Format(.rc41, "0.0000"), Format(.rc42, "0.0000"), Format(.rc43, "0.0000"), Format(.rc44, "0.0000")
        End If
        
'        Debug.Print Format(.rc11, "0.0000"), Format(.rc12, "0.0000"), Format(.rc13, "0.0000"), Format(.rc14, "0.0000")
'        Debug.Print Format(.rc21, "0.0000"), Format(.rc22, "0.0000"), Format(.rc23, "0.0000"), Format(.rc24, "0.0000")
'        Debug.Print Format(.rc31, "0.0000"), Format(.rc32, "0.0000"), Format(.rc33, "0.0000"), Format(.rc34, "0.0000")
'        Debug.Print Format(.rc41, "0.0000"), Format(.rc42, "0.0000"), Format(.rc43, "0.0000"), Format(.rc44, "0.0000")
'        Debug.Print
        
    End With
    
End Sub



