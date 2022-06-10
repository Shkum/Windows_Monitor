Imports System.Text
Imports System.Windows.Forms
Public Class Form1

    Const STRING_BUFFER_LENGTH As Integer = 255
    Dim WinHwnd As Long

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        If CLng(WinApi.WinFromCursor()) = CLng(WinHwnd) Then Exit Sub
        If e.Button = MouseButtons.Left Then
            If CLng(WinApi.WinFromCursor()) = CLng(WinHwnd) Then Exit Sub
            Cursor = Cursors.Cross
            WinHwnd = WinApi.WinFromCursor()
            TextBox2.Text = CStr(WinHwnd)
            WinApi.GetWindowText(WinHwnd, windowText, STRING_BUFFER_LENGTH)
            TextBox1.Text = windowText.ToString
            WinApi.GetClassName(WinHwnd, className, STRING_BUFFER_LENGTH)
            TextBox3.Text = className.ToString
            ListBox1.Items.Clear()
            WinApi.EnumChildWindows(WinHwnd, New WinApi.EnumWindowsCallback(AddressOf FillList), 0)
            Label4.Text = "Child windows: " & CStr(ListBox1.Items.Count)
        End If
        If CheckBox2.Checked = True And e.Button = MouseButtons.Left Then
            Me.Width = 300
            Me.Height = 1
        End If
    End Sub

    Function FillList(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        Dim c_cnt As Integer, c_txt As String
        WinApi.GetWindowText(hWnd, windowText, STRING_BUFFER_LENGTH)
        WinApi.GetClassName(hWnd, className, STRING_BUFFER_LENGTH)
        If CheckBox3.Checked = True Then Me.Opacity = 0.4
        If ListBox1.Items.Count > 0 Then
            For c_cnt = 0 To ListBox1.Items.Count - 1
                c_txt = c_txt & CStr(ListBox1.Items.Item(c_cnt))
            Next
        End If
        If InStr(c_txt, CStr(hWnd)) = 0 Then ListBox1.Items.Add(hWnd & vbTab & className.ToString & vbTab & windowText.ToString)
        WinApi.EnumChildWindows(hWnd, New WinApi.EnumWindowsCallback(AddressOf FillNode), 0)
        Return True
    End Function

    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        Cursor = Cursors.Arrow
        Me.Opacity = 1
        Me.Width = 1004
        Me.Height = 393
        TxTClick()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
    End Sub

    Function FillNode(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        WinApi.GetWindowText(hWnd, windowText, STRING_BUFFER_LENGTH)
        WinApi.GetClassName(hWnd, className, STRING_BUFFER_LENGTH)
        ListBox1.Items.Add("> " & hWnd & vbTab & className.ToString & vbTab & windowText.ToString)
        WinApi.EnumChildWindows(hWnd, New WinApi.EnumWindowsCallback(AddressOf Fillnode1), 0)
        Return True
    End Function

    Function Fillnode1(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        WinApi.GetWindowText(hWnd, windowText, STRING_BUFFER_LENGTH)
        WinApi.GetClassName(hWnd, className, STRING_BUFFER_LENGTH)
        ListBox1.Items.Add("> > " & hWnd & vbTab & className.ToString & vbTab & windowText.ToString)
        WinApi.EnumChildWindows(hWnd, New WinApi.EnumWindowsCallback(AddressOf Fillnode2), 0)
        Return True
    End Function
    Function Fillnode2(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        WinApi.GetWindowText(hWnd, windowText, STRING_BUFFER_LENGTH)
        WinApi.GetClassName(hWnd, className, STRING_BUFFER_LENGTH)
        ListBox1.Items.Add("> > > " & hWnd & vbTab & className.ToString & vbTab & windowText.ToString)
        WinApi.EnumChildWindows(hWnd, New WinApi.EnumWindowsCallback(AddressOf Fillnode3), 0)
        Return True
    End Function
    Function Fillnode3(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        WinApi.GetWindowText(hWnd, windowText, STRING_BUFFER_LENGTH)
        WinApi.GetClassName(hWnd, className, STRING_BUFFER_LENGTH)
        ListBox1.Items.Add("> > > > " & hWnd & vbTab & className.ToString & vbTab & windowText.ToString)
        WinApi.EnumChildWindows(hWnd, New WinApi.EnumWindowsCallback(AddressOf Fillnode4), 0)
        Return True
    End Function
    Function Fillnode4(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        WinApi.GetWindowText(hWnd, windowText, STRING_BUFFER_LENGTH)
        WinApi.GetClassName(hWnd, className, STRING_BUFFER_LENGTH)
        ListBox1.Items.Add("> > > > > " & hWnd & vbTab & className.ToString & vbTab & windowText.ToString)
        Return True
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox("Make left mouse button down on the picture." & vbCrLf & "Move mouse to any window." & vbCrLf & "Release mouse button." & vbCrLf _
               & "Select window in the 'Child windows list' or 'Parrent windoow text'." & vbCrLf & "And see window information on the right top panel." & vbCrLf _
               & "You may send modifying message from right bottom panel." & vbCrLf & vbCrLf & "You may test sending messages to this button." & vbCrLf _
               & "This button automaticaly selected when program starts." & vbCrLf & vbCrLf & "Created by Sergrey Shkumat (pepelnici@meta.ua)", vbInformation, "About...")
    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        ListClick()
    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        TxTClick()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If RadioButton1.Checked = True Then
            WinApi.EnableWindow(CLng(Label6.Text), 1)

        ElseIf RadioButton2.Checked = True Then
            WinApi.EnableWindow(CLng(Label6.Text), 0)

        ElseIf RadioButton3.Checked = True Then
            WinApi.SetWindowPos(CLng(Label6.Text), 0, 0, 0, 0, 0, &H10 Or &H80 Or &H2 Or &H1)

        ElseIf RadioButton4.Checked = True Then
            WinApi.SetWindowPos(CLng(Label6.Text), 0, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1)


        ElseIf RadioButton5.Checked = True Then
            WinApi.SetWindowPos(CLng(Label6.Text), -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1)

        ElseIf RadioButton6.Checked = True Then
            WinApi.SetWindowPos(CLng(Label6.Text), -2, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1)

        ElseIf RadioButton7.Checked = True Then
            WinApi.SetWindowText(CLng(Label6.Text), TextBox4.Text)

        ElseIf RadioButton8.Checked = True Then
            WinApi.SetForegroundWindow(CLng(Label6.Text))

        ElseIf RadioButton9.Checked = True Then
            WinApi.SendMessage(CLng(Label6.Text), &H10, 0, 0)

        Else
            MsgBox("Select message to send.", vbInformation, "Message not send...")
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Visible = True
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        Dim wHwnd As Integer
        wHwnd = CInt(Button1.Handle)
        WinApi.GetWindowText(CLng(wHwnd), windowText, STRING_BUFFER_LENGTH)
        WinApi.GetClassName(CLng(wHwnd), className, STRING_BUFFER_LENGTH)
        Label6.Text = CStr(wHwnd)
        Label9.Text = windowText.ToString
        Label11.Text = className.ToString
        Label13.Text = CStr(CBool(WinApi.IsWindowVisible(wHwnd)))
        Label15.Text = CStr(CBool(WinApi.IsWindowEnabled(wHwnd)))
        Label16.Text = CStr(GetParent(CInt(wHwnd)))
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ListClick()
    End Sub

    Private Sub ListClick()
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        Dim wHwnd As String, c_Txt() As String
        If ListBox1.Text <> "" Then
            c_Txt = Split(ListBox1.Text, vbTab)
            wHwnd = c_Txt(0)
            Do While Strings.Left(wHwnd, 2) = "> "
                wHwnd = Strings.Right(wHwnd, wHwnd.Length - 2)
            Loop
            WinApi.GetWindowText(CLng(wHwnd), windowText, STRING_BUFFER_LENGTH)
            WinApi.GetClassName(CLng(wHwnd), className, STRING_BUFFER_LENGTH)
            Label6.Text = wHwnd
            Label9.Text = windowText.ToString
            Label11.Text = className.ToString
            Label13.Text = CStr(CBool(WinApi.IsWindowVisible(CInt(wHwnd))))
            Label15.Text = CStr(CBool(WinApi.IsWindowEnabled(CInt(wHwnd))))
            Label16.Text = CStr(GetParent(CInt(wHwnd)))
        End If
    End Sub


    Private Sub TxTClick()
        If TextBox2.Text = "" Then Exit Sub
        Dim windowText, className As New StringBuilder(STRING_BUFFER_LENGTH)
        Dim wHwnd As String
        wHwnd = TextBox2.Text
        WinApi.GetWindowText(CLng(wHwnd), windowText, STRING_BUFFER_LENGTH)
        WinApi.GetClassName(CLng(wHwnd), className, STRING_BUFFER_LENGTH)
        Label6.Text = wHwnd
        Label9.Text = windowText.ToString
        Label11.Text = className.ToString
        Label13.Text = CStr(CBool(WinApi.IsWindowVisible(CInt(wHwnd))))
        Label15.Text = CStr(CBool(WinApi.IsWindowEnabled(CInt(wHwnd))))
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ListBox1.Items.Clear()
        Label4.Text = "All windows:"
        WinApi.EnumWindowsDllImport(New WinApi.EnumWindowsCallback(AddressOf FillActiveWindowsList), 0)
    End Sub


    Function FillActiveWindowsList(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        Dim windowText As New StringBuilder(STRING_BUFFER_LENGTH)
        WinApi.GetWindowText(CLng(hWnd), windowText, STRING_BUFFER_LENGTH)
        Dim windowIsOwned As Boolean
        Dim windowStyle As Integer
        windowIsOwned = WinApi.GetWindow(hWnd, 4) <> 0
        windowStyle = WinApi.GetWindowLong(hWnd, -20)
        If windowText.ToString <> "" Then
            If WinApi.GetParent(hWnd) = 0 Then
                If windowStyle <> 128 And windowIsOwned = False Then
                    If CheckBox4.Checked = True Then
                        If CBool(WinApi.IsWindowVisible(hWnd)) = True Then
                            ListBox1.Items.Add(hWnd & vbTab & windowText.ToString)
                        End If
                    Else
                        ListBox1.Items.Add(hWnd & vbTab & windowText.ToString)
                    End If

                End If
            End If
        End If

        Return True
    End Function
End Class
