Attribute VB_Name = "context_menu"
Option Private Module
Option Explicit

' -----------------------------------------------
' context menu title
Const MENU_NAME As String = "MENU TEST"
' -----------------------------------------------

' Excel version
Const EXCEL2013 As Integer = 15

Sub show_msg()
    MsgBox "Context menu test message", Title:="Context menu test"
End Sub


'右クリックメニュー追加
Sub add_menu()
    On Error Resume Next

    If is_overlap(ActiveWindow) > 0 Then
        delete_menu ("full")
    End If
    
    With Application.CommandBars("Cell").Controls.Add(temporary:=True)
        .Caption = MENU_NAME
        .OnAction = "show_msg"
        .BeginGroup = True
    End With
    
End Sub


'右クリックメニュー削除
Sub delete_menu(Optional status As String = "normal")
    On Error Resume Next
    
    Dim myWindow As Variant
    If status = "full" Then
        If CInt(Application.Version) >= EXCEL2013 Then
            For Each myWindow In Application.Windows
                While is_overlap(myWindow) > 0
                    myWindow.Activate
                    Application.CommandBars("Cell").Controls(MENU_NAME).Delete
                Wend
            Next myWindow
        Else
            Application.CommandBars("Cell").Controls(MENU_NAME).Delete
        End If
    Else
        If is_overlap(ActiveWindow) > 0 Then
            Application.ActiveWindow.Activate
            Application.CommandBars("Cell").Controls(MENU_NAME).Delete
        End If
    End If

End Sub



'戻り値:　メニュータイトルの重複数
Function is_overlap(myWindow As Variant) As Integer
    On Error Resume Next

    Dim cb As CommandBarControl
    Dim menu_count As Integer
    menu_count = 0
    myWindow.Activate
    
    For Each cb In CommandBars("Cell").Controls
        If cb.Caption = MENU_NAME Then
            menu_count = menu_count + 1
        End If
    Next cb
    
    is_overlap = menu_count
End Function



