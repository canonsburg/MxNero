Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.
    Static Sub Main()
        app = New window1()
        app.InitializeComponent()
        app.Run()
    End Sub

End Class
