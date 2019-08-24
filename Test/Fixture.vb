Public Class Fixture : Inherits AsyncTest.Fixture

    Private Timer As System.Timers.Timer

    Private TheForm As Form1

    Public Sub New(ByVal TheForm As Form1)

        Me.TheForm = TheForm

    End Sub

    <AsyncTest.Test()> Public Sub TestMyException()

        Throw New Exception("Blah")

    End Sub

    <AsyncTest.Test()> Public Sub TestExceptionThrown()

        Dim Result As IAsyncResult = _
        Net.Dns.BeginGetHostByName(Nothing, Nothing, Nothing)

    End Sub

    <AsyncTest.Test()> Public Sub TestUnableToResolve()

        Dim Result As IAsyncResult = _
        Net.Dns.BeginGetHostByName("ausdiasasadasd", WrapCallback(AddressOf UnableToResolveCallback), Nothing)

        If Result.CompletedSynchronously Then
            Debug.WriteLine("Didn't find any addresses")
            EndTest()
        End If

    End Sub

    <AsyncTest.Test()> Public Sub TestResolveOk()

        Dim Result As IAsyncResult = _
        Net.Dns.BeginGetHostByName("rikkus.info", WrapCallback(AddressOf ResolveOkCallback), Nothing)

        If Result.CompletedSynchronously Then
            Debug.WriteLine("Didn't find any addresses")
            EndTest()
        End If

    End Sub

    Private Sub UnableToResolveCallback(ByVal result As IAsyncResult)

        Try

            Dim Entry As Net.IPHostEntry = Net.Dns.EndGetHostByName(result)

        Catch ex As Exception

            EndTest()
            Return

        End Try

        AssertFail("Shouldn't have resolved anything")

    End Sub

    Private Sub ResolveOkCallback(ByVal result As IAsyncResult)

        Dim Entry As Net.IPHostEntry = Net.Dns.EndGetHostByName(result)

        AssertEqual(Entry.AddressList.Length.ToString, 1.ToString)

    End Sub

    <AsyncTest.Test()> Public Sub TestEventHandler()

        Timer = New System.Timers.Timer

        AddEventHandler(Timer, "Elapsed", "TimerCallback")

        Timer.AutoReset = False
        Timer.Interval = 100
        Timer.Start()

    End Sub

    Public Sub TimerCallback(ByVal Sender As Object, ByVal Args As Timers.ElapsedEventArgs)

        AssertEqual("abc", "abc")

    End Sub

    <AsyncTest.Test()> Public Sub TestAssertFailEventHandler()

        Timer = New System.Timers.Timer

        AddEventHandler(Timer, "Elapsed", "AssertFailTimerCallback")

        Timer.AutoReset = False
        Timer.Interval = 100
        Timer.Start()

    End Sub

    Public Sub AssertFailTimerCallback(ByVal Sender As Object, ByVal Args As Timers.ElapsedEventArgs)

        AssertFail("I am failing an assertion on purpose")

    End Sub

    <AsyncTest.Test()> Public Sub TestBrokenEventHandler()

        Timer = New System.Timers.Timer

        AddEventHandler(Timer, "Elapsed", "BrokenTimerCallback")

        Timer.AutoReset = False
        Timer.Interval = 100
        Timer.Start()

    End Sub

    Public Sub BrokenTimerCallback(ByVal Sender As Object, ByVal Args As Timers.ElapsedEventArgs)

        Throw New System.Exception("This exception is expected")

    End Sub

    <AsyncTest.Test()> Public Sub TestCheckBoxChangesButtonEnabled()

        AddEventHandler(Me.TheForm.CheckBox1, "CheckedChanged", "CheckedChangedCallback")

        Me.TheForm.CheckBox1.Checked = True

    End Sub

    Public Sub CheckedChangedCallback(ByVal sender As System.Object, ByVal e As System.EventArgs)

        AssertEqual(Me.TheForm.CheckBox1.Checked, Me.TheForm.Button1.Enabled)

    End Sub

End Class
