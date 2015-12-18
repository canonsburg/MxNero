Imports System
Imports System.IO
Imports System.Text
Imports System.ComponentModel ' CancelEventArgs
Imports System.Windows ' window
Imports System.Windows.Threading
Imports System.Net
Imports System.Net.Sockets
Imports System.Diagnostics 'Need this for stopwatch()


Public Class Window2
    Inherits Window

    Public url As String = "http://localhost:8082/json_rpc"    'The address of the simplewallet

    Public Sub dispatcherTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)

        If URLExists(url) Then
            GetBalance()
        End If
    End Sub

    Private Sub DataWindow_Closing(ByVal sender As Object, ByVal e As CancelEventArgs)
        'MessageBox.Show("Exiting Wallet Manager. Do You Want to Save Wallet?")
        If IsProcessRunning("simplewallet.exe") Then    'Check if simplewallet is still running

            If URLExists(url) Then
                SaveWallet()
                TerminateApp()
                Exit Sub
            End If
            TerminateApp()
        End If
    End Sub

    Private Sub Form_Load()
        If Not IsProcessRunning("simplewallet.exe") Then
            MessageBox.Show("Error! Simplewallet.exe is not running. Please Close Wallet Manager and Restart")
            Exit Sub
        Else

            'Check If simplewallet binding is effective
            MessageBox.Show("Loading Wallet! Please Let Wallet Sychronize..." & vbCrLf & "This May Take Up to 20 Mins to Sychronize from Scratch... Please Wait...")
            Dim timer As New Stopwatch
            timer.Start()
            Do While Not URLExists(url)
                Threading.Thread.Sleep(100)
                If Not IsProcessRunning("simplewallet.exe") Then
                    MessageBox.Show("Error! Simplewallet.exe is not running. Please Close Wallet Manager and Restart")
                    Exit Sub
                End If

                If timer.ElapsedMilliseconds > 100 * 60 * 1000 Then 'break out of loop after 100 minutes
                    MessageBox.Show("Error: Simplewallet Synchronization Taking Too Long. Exited Process")
                    TerminateApp()  'Terminate Simplewallet.exe
                    Exit Sub
                End If
            Loop

            TextAddress.Text = GetAddress()
            TextDaemonAddress.Text = "node.moneroclub.com:8880"
            TextSendToAddress.Text = ""
            TextSendAmount.Text = ""
            TextPaymentID.Text = ""
            If URLExists(url) Then
                GetBalance()
            Else
                MessageBox.Show("Error! Cannot Connect to Remote Daemon. Please Check Network Connection.")
                Exit Sub
            End If
        End If

        Dim dt As DispatcherTimer = New DispatcherTimer()   'Create a new DispatcherTimer
        AddHandler dt.Tick, AddressOf dispatcherTimer_Tick  'Set it to the defined subroutine
        dt.Interval = New TimeSpan(0, 0, 120)   'updates ever 120 seconds
        dt.Start()
    End Sub

    Function GetAddress() As String

        Dim Request As String = "{""jsonrpc"":""2.0"",""id"":""0"",""method"":""getaddress""}"

        On Error GoTo ErrorHandler

        Dim MyRequest As Object = CreateObject("MSXML2.XMLHTTP")
        MyRequest.Open("POST", url, False)
        MyRequest.setRequestHeader("Content-Type", "application/json")
        MyRequest.Send(Request)

        Dim rpcRspon_Address As String = MyRequest.ResponseText  'Store json RPC response
        Dim Srt_Address As Integer = InStr(1, rpcRspon_Address, """address""", vbTextCompare)  'variable to hold start of the string "address" in the json rpc response
        Dim str_BfrAddress As String = Mid(rpcRspon_Address, Srt_Address, Len(rpcRspon_Address) - Srt_Address + 1)   'string that contains json rpc response that begins with "address"
        Dim End_Address As Integer = InStr(1, str_BfrAddress, "}", vbTextCompare)  'variable that index end of json rpc "address" portion of response
        Dim str_Address As String   ' get the string that contains address
        str_Address = Mid(str_BfrAddress, 12, End_Address - 12) 'Note: json rpc '"address": ' contains 11 char but need to add 1 more to get start of the string representing numeric balance
        str_Address = Replace(str_Address, """", "")    'the wallet address is enclosed with quotations, get rid of them
        str_Address = Trim(str_Address) 'get rid of extra space chars before and after address

        GetAddress = str_Address
        Exit Function

ErrorHandler:
        HandleError()

    End Function

    Public Function GetBalance()

        Dim Request As String = "{""jsonrpc"":""2.0"",""id"":""0"",""method"":""getbalance"", ""params"": {}}"

        On Error GoTo ErrorHandler

        Dim MyRequest As Object = CreateObject("MSXML2.XMLHTTP")
        With MyRequest
            .Open("POST", url, False)
            .setRequestHeader("Content-Type", "application/json")
            .Send(Request)
        End With

        Dim rpcRspon_Balance As String = MyRequest.ResponseText
        Dim Srt_Balance As Integer = InStr(1, rpcRspon_Balance, """balance""", vbTextCompare)  'variable to hold start of the string "balance" in the json rpc response
        Dim str_BfrBalance As String = Mid(rpcRspon_Balance, Srt_Balance, Len(rpcRspon_Balance) - Srt_Balance + 1)   'string that contains json rpc response that begins with "balance"
        Dim End_Balance As Integer = InStr(1, str_BfrBalance, ",", vbTextCompare) 'variable that index end of json rpc "balance" portion of response
        Dim str_Balance As String = Mid(str_BfrBalance, 12, End_Balance - 12) ' get the string that contains the numeric balance
        'Note: json rpc '"balance": ' contains 11 char but need to add 1 more to get start of the string representing numeric balance
        Dim Srt_UnlockedBalance As Integer = InStr(1, rpcRspon_Balance, """unlocked_balance""", vbTextCompare) 'variable to hold start of the string "unlocked_balance" in the json rpc response
        Dim str_BfrUnlockedBalance As String = Mid(rpcRspon_Balance, Srt_UnlockedBalance, Len(rpcRspon_Balance) - Srt_UnlockedBalance + 1) 'string that contains json rpc response that begins with "unlocked_balance"
        Dim End_UnlockedBalance As Integer = InStr(1, str_BfrUnlockedBalance, "}", vbTextCompare) 'variable that index end of json rpc "unlocked_balance" portion of response
        Dim str_UnlockedBalance As String = Mid(str_BfrUnlockedBalance, 21, End_UnlockedBalance - 21) ' get the string that contains the numeric unlocked_balance
        'Note: json rpc '"unlocked_balance":  ' contains 12 char but need to add 1 more to get start of the string representing numeric balance
        str_UnlockedBalance = Trim(str_UnlockedBalance)

        Dim num_Balance As Double = Val(str_Balance) / 1000000000000.0# ' convert from str to num and display in moneros
        Dim num_UnlockedBalance As Double = Val(str_UnlockedBalance) / 1000000000000.0# ' convert from str to num and display in moneros

        TextBalance.Text = CStr(num_Balance)
        TextUnlockedBalance.Text = CStr(num_UnlockedBalance)
        TextAvailableBalance.Text = CStr(num_UnlockedBalance)

        Exit Function

ErrorHandler:
        HandleError()
    End Function

    Function URLExists(url As String) As Boolean
        URLExists = False
        Dim myurl As New System.Uri(url)
        Dim wRequest As System.Net.WebRequest = System.Net.WebRequest.Create(myurl)
        wRequest.Timeout = 1000000000  'make timeout to be 100 minutes
        Dim wResponse As System.Net.WebResponse
        Try
            wResponse = wRequest.GetResponse()

            'Is the responding address the same as HostAddress to avoid false positive from an automatic redirect.
            If wResponse.ResponseUri.AbsoluteUri().ToString = url Then 'include query strings
                URLExists = True
            End If
            wResponse.Close()
            wRequest = Nothing
        Catch ex As Exception
            wRequest = Nothing
            'MessageBox.Show(ex.ToString)   This currently throws an socket exception but going to ignore it for now
        End Try

        Return URLExists
    End Function

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

    Sub HandleError()
        If Err.Number <> 0 Then
            MessageBox.Show("Error occurred: " & Err.Number & " - " & Err.Description)
        End If

        End ' Exit the macro entirely
    End Sub

    Private Sub ButtonSendMoney_Click(sender As Object, e As RoutedEventArgs) Handles ButtonSendMoney.Click
        If TextSendToAddress.Text IsNot Nothing And TextSendAmount.Text IsNot Nothing Then
            If TextSendToAddress.Text <> "" And TextSendAmount.Text <> "" Then
                If Not IsNumVal(TextSendAmount.Text) Then
                    MessageBox.Show("Error! Amount Entered is NOT a valid number")
                    Exit Sub
                End If

                GetBalance()  'Get updated wallet balance
                If Val(TextUnlockedBalance.Text) > Val(TextSendAmount.Text) Then    'Only allow SEND if Unlocked Balance is greater than Send Amount
                    Dim myAddress As String
                    Dim myAmount As String
                    Dim SentGood As Boolean
                    Dim myMixin As String
                    Dim myPaymentID As String
                    myAddress = TextSendToAddress.Text
                    myAmount = TextSendAmount.Text
                    myAmount = myAmount * 1000000000000.0#  'rescale amount for use in simplewallet

                    If TextPaymentID.Text Is Nothing Or TextPaymentID.Text = "" Then
                        myMixin = Math.Floor(sliderMixin.Value)
                        myMixin = CStr(myMixin)
                        If CInt(myMixin) < 2 Or CInt(myMixin) > 20 Then
                            myMixin = "2"
                        End If
                        If IsProcessRunning("simplewallet.exe") Then    'Check if simplewallet is still running
                            If URLExists(url) Then
                                SentGood = SendMoney(myAddress, myAmount, myMixin, "")  'Send Money with null paymentID
                                If SentGood Then
                                    MessageBox.Show("Success! Money Sent")
                                Else
                                    MessageBox.Show("Error! Cannot Send Funds")
                                    Exit Sub
                                End If
                            End If
                        End If
                    Else
                        myMixin = CStr(sliderMixin.Value)
                        If CInt(myMixin) < 2 Or CInt(myMixin) > 20 Then
                            myMixin = "2"
                        End If

                        If Len(TextPaymentID.Text) <> 64 Then
                            MessageBox.Show("Error! PaymentID has to be 64-character long Hex")
                        Else
                            If Not IsHexVal(TextPaymentID.Text) Then
                                MessageBox.Show("Error! PaymentID must have a Hex String")
                            Else
                                myPaymentID = TextPaymentID.Text
                                If IsProcessRunning("simplewallet.exe") Then    'Check if simplewallet is still running
                                    If URLExists(url) Then
                                        SentGood = SendMoney(myAddress, myAmount, myMixin, myPaymentID)  'Send Money with paymentID
                                        If SentGood Then
                                            MessageBox.Show("Success! Money Sent")
                                        Else
                                            MessageBox.Show("Error! Cannot Send Funds")
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    MessageBox.Show("Error! Cannot Send, Unlocked Balance Less Than Send Amount Request")
                End If
            Else
                MessageBox.Show("Error! Address or Amount Value Entered Wrong")
            End If
        Else
            MessageBox.Show("Error! Address or Amount Value Entered Wrong")
        End If

        If IsProcessRunning("simplewallet.exe") Then    'Check if simplewallet is still running

            If URLExists(url) Then
                GetBalance() 'Update New Balance
            End If
        End If

        Dim IsSaved As Boolean = SaveWallet()
    End Sub

    Public Function SendMoney(ByVal Address As String, ByVal Amount As String, ByVal Mixin As String, ByVal PaymentID As String) As Boolean

        Dim Request As String = "{""jsonrpc"":""2.0"",""id"":""0"",""method"":""transfer"", ""params"": {""destinations"": [{""amount"":" & Amount & ",""address"":""" & Address & """}],""payment_id"":""" & PaymentID & """,""fee"":0,""mixin"":" & Mixin & ",""unlock_time"":0}}"

        On Error GoTo ErrorHandler

        Dim MyRequest As Object = CreateObject("MSXML2.XMLHTTP")
        With MyRequest
            .Open("POST", url, False)
            .setRequestHeader("Content-Type", "application/json")
            .Send(Request)
        End With

        Dim rpcRspon_Transfer As String = MyRequest.ResponseText
        If rpcRspon_Transfer.Contains("error") Then
            SendMoney = False
        Else
            SendMoney = True
        End If
        Exit Function

ErrorHandler:
        HandleError()
    End Function

    Public Function SaveWallet() As Boolean

        Dim Request As String
        Request = "{""jsonrpc"":""2.0"",""id"":""0"",""method"":""store""}"
        Dim MyRequest As Object = CreateObject("MSXML2.XMLHTTP")
        With MyRequest
            .Open("POST", url, False)
            .setRequestHeader("Content-Type", "application/json")
            .Send(Request)
        End With

        If MyRequest.ResponseText IsNot Nothing Then
            SaveWallet = True
        Else
            SaveWallet = False
        End If

    End Function

    Private Sub ButtonSave_Click(sender As Object, e As RoutedEventArgs) Handles ButtonSave.Click
        If IsProcessRunning("simplewallet.exe") Then    'Check if simplewallet is still running

            If URLExists(url) Then
                Dim IsSaved As Boolean = SaveWallet()
            Else
                MessageBox.Show("Error! Cannot Save Wallet! Cannot Connect to Simplewallet. Please Check Network Connection." & vbCrLf & "Exit Current Program and Restart")
                Exit Sub
            End If
        Else
            MessageBox.Show("Error! Cannot Find Simplewallet.exe! Please Restart Program")
            Exit Sub
        End If
        MessageBox.Show("Wallet Saved!")
    End Sub

    Public Function IsHexVal(ByVal str As String) As Boolean
        Dim isHex As Boolean

        For i = 1 To Len(str)
            isHex = (Mid(str, i, 1) >= "0" And Mid(str, i, 1) <= "9") Or (Mid(str, i, 1) >= "a" And Mid(str, i, 1) <= "f") Or (Mid(str, i, 1) >= "A" And Mid(str, i, 1) <= "F")
            If (Not isHex) Then
                IsHexVal = False
                Exit Function
            End If
        Next i

        IsHexVal = True

    End Function

    Public Function IsNumVal(ByVal Number As String) As Boolean
        'Function Checks if Number is a valid interpretation of acceptable number, in decimal format
        Dim isNum As Boolean
        Dim Count As Integer
        Count = 0

        For i = 1 To Len(Number)
            isNum = (Mid(Number, i, 1) >= "0" And Mid(Number, i, 1) <= "9") Or Mid(Number, i, 1) = "."
            If (Not isNum) Then
                IsNumVal = False
                Exit Function
            End If
            If Mid(Number, i, 1) = "." Then 'If More than one occurence of ".", invalid
                Count = Count + 1
                If Count > 1 Then
                    IsNumVal = False
                    Exit Function
                End If
            End If
        Next i

        If Mid(Number, 1, 1) = "." And Mid(Number, Len(Number), 1) = "." Then   'If first or last char is ".", invalid
            IsNumVal = False
            Exit Function
        End If

        If InStr(1, Number, ".") <> 0 Then   'Check if number has decimals that are smaller than allowed denomination, i.e. 1 / 1000000000000
            Dim LastPart As String  'get the part of num string that is to the right of the decimal place
            Dim pastSmall As Boolean 'if the portion of LastPart is smaller than 12 decimal places
            LastPart = Mid(Number, InStr(1, Number, ".") + 1, Len(Number) - InStr(1, Number, "."))
            If Len(LastPart) > 12 Then
                For i = 1 To Len(LastPart)
                    pastSmall = (Mid(LastPart, i, 1) >= "1" And Mid(LastPart, i, 1) <= "9") 'if the part past 12 decimal places is non zero, invalid
                    If (Not pastSmall) Then
                        IsNumVal = False
                        MessageBox.Show("Error! Input Send Amount contains digits that exceed smallest denomination")
                        Exit Function
                    End If
                Next i
            End If
        End If

        IsNumVal = True
    End Function

End Class
