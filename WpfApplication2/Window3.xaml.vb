Imports System.IO 'need this to use File.Exists
Imports System.Diagnostics
'Imports System.Windows.Forms

Public Class Window3
    Dim simplewalletpath As String
    Dim walletpath As String
    Public IsCheckedSimpleWallet As Boolean

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'Clear all Textboxes
        TextBoxCreateWalletName.Text = ""
        PasswordBox.Password = ""
        IsCheckedSimpleWallet = False

    End Sub

    Private Sub ButtonFolderSelection_Click(sender As Object, e As RoutedEventArgs) Handles ButtonFolderSelection.Click
        'Clear all Textboxes
        TextBoxCreateWalletName.Text = ""
        PasswordBox.Password = ""

        ' Configure open file dialog box
        Dim fDialog As New Microsoft.Win32.OpenFileDialog()

        With fDialog
            ' Allow the user to make multiple selections in the dialog box.
            .Multiselect = False

            ' Set the title of the dialog box.
            .Title = "Select simplewallet.exe"
            .DefaultExt = ".exe" ' Default file extension
            .Filter = "simplewallet.exe|*.exe" ' Filter files by extension
        End With

        ' Show the dialog box. If the .Show method returns True, the
        ' user picked at least one file. If the .Show method returns
        ' False, the user clicked Cancel.
        Dim result? As Boolean = fDialog.ShowDialog()

        ' Process open file dialog box results
        If result = True Then
            ' Open document
            simplewalletpath = fDialog.FileName

        Else
            simplewalletpath = ""
            'MessageBox.Show("You Clicked Cancel")
        End If

        If (simplewalletpath IsNot Nothing) And simplewalletpath <> "" Then
            Dim i As Integer
            i = InStrRev(simplewalletpath, "\")
            walletpath = Strings.Left(simplewalletpath, i)
        End If

    End Sub

    Private Sub ButtonCreateWallet_Click(sender As Object, e As RoutedEventArgs) Handles ButtonCreateWallet.Click
        Dim NewWalletName As String = TextBoxCreateWalletName.Text
        Dim NewWalletPassword As String = PasswordBox.Password
        PasswordBox.Password = "" 'Clear PasswordBox

        If (simplewalletpath Is Nothing) Or simplewalletpath = "" Then
            MessageBox.Show("Error! Please Select File Path to simplewallet.exe")
            Exit Sub
        End If
        If NewWalletName Is Nothing Or NewWalletName = "" Then
            MessageBox.Show("Error! Please Enter Wallet Name")
            Exit Sub
        ElseIf NewWalletName.Contains("\") Then
            MessageBox.Show("Error! Wallet Name Cannot Contain '\' Character")
            Exit Sub
        End If
        If NewWalletPassword Is Nothing Or NewWalletPassword = "" Then
            MessageBox.Show("Error! Password is Null or Empty Char")
            Exit Sub
        End If

        If File.Exists(walletpath & NewWalletName & ".keys") Then
            MessageBox.Show("Error! Wallet With Name: '" & NewWalletName & "' Already Exists! Please Choose Another Wallet Name")
            Exit Sub
        End If

        If File.Exists(simplewalletpath) Then
            Call RunSimplewallet(simplewalletpath, NewWalletName, NewWalletPassword) 'Launch simplewallet executable
            WaitSomeTime()

            If File.Exists(walletpath & NewWalletName & ".keys") Then
                TerminateApp() 'kill simplewallet.exe
                Call RPCModeRunSimplewallet(simplewalletpath, NewWalletName, NewWalletPassword) 'Launch simplewallet executable; in RPC mode
                NewWalletPassword = ""  'clear password variable
                Me.Hide()
                Me.Close()
            Else
                TerminateApp() 'kill simplewallet.exe
                NewWalletPassword = ""  'clear password variable
                Me.Close()
            End If

        Else
            MessageBox.Show("Error! Cannot Find simplewallet.exe, Please Select It Again")
            Exit Sub
        End If

    End Sub

    Private Sub RunSimplewallet(ByVal SimpleWalletPath As String, ByVal WalletName As String, ByVal Password As String)
        Dim objShell As Object
        objShell = CreateObject("WScript.Shell")
        Dim CommandStr As String

        CommandStr = SimpleWalletPath '& " --generate-new-wallet=" & WalletName & " --password=" & Password
        On Error GoTo ErrorHandler
        objShell.CurrentDirectory = GetPathName(SimpleWalletPath) 'Set current directory of Shell
        objShell.Run(CommandStr, 1)
        WaitSomeTime()
        My.Computer.Keyboard.SendKeys(WalletName & "{ENTER}", True)
        WaitSomeTime()
        My.Computer.Keyboard.SendKeys(Password & "{ENTER}", True)
        WaitSomeTime()
        My.Computer.Keyboard.SendKeys("0" & "{ENTER}", True)

        Exit Sub

ErrorHandler:
        HandleError()
    End Sub

    Private Sub RPCModeRunSimplewallet(ByVal SimpleWalletPath As String, ByVal WalletName As String, ByVal Password As String)
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
End Class
