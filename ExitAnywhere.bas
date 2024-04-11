Attribute VB_Name = "ExitAnywhere"
Option Explicit ' Explicitly define variables

' this will make the form move accross
' the screen left to right right to left up and down
' call it like this ExitUp Formname
' i was exeperimenting with a timer to slow it down
' but it was a little buggy

' make the form exit downwards
Sub ExitDown(Form As Form)

    Do
        
        Form.Top = Trim(Str(Int(Form.Top) + 300))
        
        DoEvents
                            
        Loop Until Form.Top > 7200

        If Form.Top > 7200 Then Form.Hide 'End
                                    
End Sub

' make the form exit left
Sub ExitLeft(Form As Form)

    Do
        
        Form.Left = Trim(Str(Int(Form.Left) - 300))
        
        DoEvents
        
        Loop Until Form.Left < -6300
        
        If Form.Left < -6300 Then End
        
End Sub

' make the form exit right
Sub ExitRight(Form As Form)

    Do
        
        Form.Left = Trim(Str(Int(Form.Left) + 300))

        DoEvents
        
        Loop Until Form.Left > 9600
        
        If Form.Left > 9600 Then End
        
End Sub

' make the form exit upwards
Sub ExitUp(Form As Form)

    Do
        
        Form.Top = Trim(Str(Int(Form.Top) - 300))

        DoEvents
        
        Loop Until Form.Top < -4500

        If Form.Top < -4500 Then End
                                
End Sub
