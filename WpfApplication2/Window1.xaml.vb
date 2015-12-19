Imports System
Imports System.IO
Imports System.Text
Imports System.Collections
Imports System.Collections.Specialized  'Need to add this in order to use my.settings specialized string collection


Public Class Window1
    Public WalletKeysPath As String 'declare global variable to store file path to wallet.keys
    Public IsCheckedSimpleWallet As Boolean 'stores state of whether to run simplewallet in hidden mode or not: checked is show, unchecked is hide

    Private Sub Form_Load()
        TerminateApp()    'Call Terminate function to kill any instances of simplewallet.exe before starting wallet manager
        TextWalletPath.Text = "Please Select Wallet File"
        IsCheckedSimpleWallet = False

    End Sub

    Private Sub ButtonSetWalletPath_Click(sender As Object, e As RoutedEventArgs) Handles ButtonSetWalletPath.Click

        ' Configure open file dialog box
        Dim fDialog As New Microsoft.Win32.OpenFileDialog()
        TextEnterPassword.Password = "" 'Clear Password Input TextBox
        TextWalletPath.Text = ""   'Clear WalletPath TextBox

        With fDialog
            ' Allow the user to make multiple selections in the dialog box.
            .Multiselect = False

            ' Set the title of the dialog box.
            .Title = "Select Wallet.keys File"
            .DefaultExt = ".keys" ' Default file extension
            .Filter = "Wallet.KEYS|*.keys" ' Filter files by extension
        End With

        ' Show the dialog box. If the .Show method returns True, the
        ' user picked at least one file. If the .Show method returns
        ' False, the user clicked Cancel.
        Dim result? As Boolean = fDialog.ShowDialog()

        ' Process open file dialog box results
        If result = True Then
            ' Open document
            WalletKeysPath = fDialog.FileName
            TextWalletPath.Text = "Current Selected Wallet.Keys File Path:" & vbCrLf & fDialog.FileName

        Else
            WalletKeysPath = ""
            TextWalletPath.Text = "Please Select Wallet File"
            'MessageBox.Show("You Clicked Cancel")
        End If

    End Sub

    Function GetPathName(ByVal strFilePathName As String) As String
        'Function returns File Name, given File Path Name (as string) minus the file extention part

        Dim LPosition As String 'variable that denotes position of char starting from right of string
        Dim DirFilePathName As String
        Dim DirPathNameFull As String   'Name of the File imported, with file extentions part of name

        DirFilePathName = strFilePathName   'Store File Path Name

        'Return position (from the Beginning/Left of string) of first occurence of '\', from the right of string
        LPosition = InStrRev(DirFilePathName, "\")

        'Extract portion of File Path string that denotes File Name
        DirPathNameFull = Mid(DirFilePathName, 1, LPosition) 'With extentions still intact
        GetPathName = DirPathNameFull

    End Function

    Function GetFileName(ByVal strFilePathName As String) As String
        'Function returns File Name, given File Path Name (as string) minus the file extention part

        Dim LPosition As String 'variable that denotes position of char starting from right of string
        Dim DirFilePathName As String   'File Path Name of Director File imported
        Dim DirFileNameFull As String   'Name of the File imported, with file extentions part of name

        DirFilePathName = strFilePathName   'Store File Path Name

        'Return position (from the Beginning/Left of string) of first occurence of '\', from the right of string
        LPosition = InStrRev(DirFilePathName, "\")

        'Extract portion of File Path string that denotes File Name
        DirFileNameFull = Mid(DirFilePathName, LPosition + 1, Len(DirFilePathName) - LPosition) 'With extentions still intact
        'GetFileName = Mid(DirFileNameFull, 1, Len(DirFileNameFull) - 5)    'removing file extentions with 5 char such as '.keys'
        GetFileName = DirFileNameFull
    End Function

    Private Sub RunSimplewallet(ByVal SimpleWalletPath As String, ByVal WalletName As String, ByVal Password As String)
        Dim objShell As Object
        objShell = CreateObject("WScript.Shell")
        Dim CommandStr As String

        CommandStr = SimpleWalletPath & " --daemon-address node.moneroclub.com:8880 --rpc-bind-port 8082 --wallet-file=" & WalletName & " --password=" & Password
        On Error GoTo ErrorHandler
        objShell.CurrentDirectory = GetPathName(SimpleWalletPath) 'Set current directory of Shell
        If IsCheckedSimpleWallet Then
            objShell.Run(CommandStr, 1)
        Else
            objShell.Run(CommandStr, 0)
        End If
        Exit Sub

ErrorHandler:
        HandleError()
    End Sub

    Sub HandleError()

        If Err.Number <> 0 Then
            MessageBox.Show("Error occurred: " & Err.Number & " - " & Err.Description)
        End If

        End ' Exit the macro entirely

    End Sub

    Public Sub WaitSomeTime()
        For iCount = 1 To 50000000    'loop through some values to wait
        Next iCount
    End Sub

    Function IsProcessRunning(process As String)
        Dim objList As Object

        objList = GetObject("winmgmts:") _
        .ExecQuery("select * from win32_process where name='" & process & "'")

        If objList.Count > 0 Then
            IsProcessRunning = True
        Else
            IsProcessRunning = False
        End If

    End Function

    Private Sub ButtonEnterWallet_Click(sender As Object, e As RoutedEventArgs) Handles ButtonEnterWallet.Click
        Dim PathSimpleWallet As String
        If (WalletKeysPath IsNot Nothing) And WalletKeysPath <> "" Then
            If TextWalletPath.Text = "Please Select Wallet File" Then
                MessageBox.Show("Error! Please Select Wallet File Path")
            Else
                PathSimpleWallet = GetPathName(WalletKeysPath) & "simplewallet.exe"
                If Dir(PathSimpleWallet) <> "" Then
                    If (TextEnterPassword.Password Is Nothing) Then
                        MessageBox.Show("Please Enter Password")
                    ElseIf TextEnterPassword.Password = "" Then
                        MessageBox.Show("Please Enter Password")
                    Else
                        Dim myPassword As String = TextEnterPassword.Password
                        TextEnterPassword.Password = "" 'Clear the Password in User Input
                        TextWalletPath.Text = "Please Wait! Loading Wallet..."
                        Call RunSimplewallet(PathSimpleWallet, GetFileName(WalletKeysPath), myPassword) 'Launch simplewallet executable
                        WaitSomeTime() 'Pause for a few seconds to wait to see if simplewallet.exe exists because of bad password
                        If IsProcessRunning("simplewallet.exe") Then    'check if simplewallet.exe is running
                            myPassword = ""
                            Me.Hide()
                            Dim WalletGUI As New Window2
                            WalletGUI.Show()
                            PathSimpleWallet = ""
                            Me.Close()
                        Else
                            myPassword = ""
                            MessageBox.Show("Error! Wrong Password Entered")
                            TextWalletPath.Text = "Please Select Wallet File"
                            PathSimpleWallet = ""
                        End If
                    End If
                Else
                    MessageBox.Show("Error! Simplewallet.exe Not Located in Same Folder Path as wallet.keys file")
                End If
            End If
        Else
            MessageBox.Show("Please Select Wallet Keys File Path")
        End If
    End Sub

    Sub TerminateApp()
        'Terminates the exe process specificed
        'Terminates ALL instances of the exe process held in variable strTerminateThis

        Dim strTerminateThis As String  'variable hold process to terminate

        Dim objWMIcimv2 As Object, objProcess As Object, objList As Object
        Dim intError As Integer

        'Process to terminate – you could specify and .exe program name here
        strTerminateThis = "simplewallet.exe"

        'Connect to CIMV2 Namespace and then find the .exe process

        objWMIcimv2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
        objList = objWMIcimv2.ExecQuery("select * from win32_process where name='" & strTerminateThis & "'")
        For Each objProcess In objList
            intError = objProcess.Terminate 'Terminates a process and all of its threads.
            'Return value is 0 for success. Any other number is an error.
            If intError <> 0 Then Exit For
        Next

        'ALL instances of exe (strTerminateThis) have been terminated
        objWMIcimv2 = Nothing
        objList = Nothing
        objProcess = Nothing

    End Sub

    Private Sub HandleCheck(ByVal sender As Object, ByVal e As RoutedEventArgs)
        IsCheckedSimpleWallet = True
    End Sub

    Private Sub HandleUnchecked(ByVal sender As Object, ByVal e As RoutedEventArgs)
        IsCheckedSimpleWallet = False
    End Sub

    Private Sub ButtonNewWallet_Click(sender As Object, e As RoutedEventArgs) Handles ButtonNewWallet.Click
        Dim CreateNewWallet As New Window3
        AddHandler CreateNewWallet.Closing, AddressOf CreateNewWallet_WindowClosed
        CreateNewWallet.Show()
        Me.Hide()
    End Sub

    Private Sub CreateNewWallet_WindowClosed()

        If IsProcessRunning("simplewallet.exe") Then    'check if simplewallet.exe is running
            Dim WalletGUI As New Window2
            WalletGUI.Show()
            Me.Close()
        Else
            MessageBox.Show("Choose to Not Create New Wallet")
            Me.Show()
        End If
    End Sub
End Class
