Imports System.Reflection

Public Class Fixture

#Region "Public API"

    Public Sub RunAllTests()

        For Each TestMethod As MethodInfo In MyTestMethods
            MyTestMethodsRemaining.Add(TestMethod)
        Next

        RunNextTest()

    End Sub

    Public Enum TestResult
        Success
        Failure
        AssertionFailed
        Timeout
        ExceptionThrownBySetup
        ExceptionThrown
    End Enum

    Public Event StartingTest(ByVal Name As String)
    Public Event TestFinished(ByVal Name As String, ByVal Result As TestResult, ByVal Explanation As String)
    Public Event FinishedRunningTests()

#End Region

#Region "API for subclasses"

    Public Sub New()

        InitTimers()

        Dim Methods As MethodInfo() = _
        MyClass.GetType.GetMethods _
        ( _
            BindingFlags.DeclaredOnly _
            Or BindingFlags.Public _
            Or BindingFlags.Instance _
        )

        For Each Info As MethodInfo In Methods

            If Info.IsDefined(GetType(AsyncTest.Test), True) Then
                MyTestMethods.Add(Info)
            End If

        Next

    End Sub

    Protected Property MaxTestTimeSeconds() As Integer
        Get
            Return MyMaxTestTimeSeconds
        End Get
        Set(ByVal Value As Integer)

            MyMaxTestTimeSeconds = Value
            If MyCurrentCaseName <> "" Then
                TestTimeoutTimer.Stop()
                TestTimeoutTimer.Interval = MyMaxTestTimeSeconds * 1000
                TestTimeoutTimer.Start()
            End If
        End Set
    End Property

    Protected Function WrapCallback(ByVal Callback As AsyncCallback) As AsyncCallback

        MyCallbackMap.Add(MyCurrentCaseName, Callback)

        Return New AsyncCallback(AddressOf CallbackWrapper)

    End Function

    Protected Sub AddEventHandler(ByVal TheObject As Object, ByVal EventName As String, ByVal CallbackName As String)

        Dim EventProducerType As Type = TheObject.GetType

        Dim AddMethodInfo As MethodInfo = Utils.AddMethod(EventProducerType, EventName)

        Dim EventHandlerType As Type = AddMethodInfo.GetParameters()(0).ParameterType

        Dim CompiledAssembly As System.Reflection.Assembly = _
        GenerateAssemblyForEventHandlerWrapperClass(EventHandlerType, EventName, CallbackName, Me)

        Dim CallbackWrapperClassType As Type = CompiledAssembly.GetType(CallbackWrapperClassName(EventHandlerType.Name))

        Dim CallbackObjectFactory As MethodInfo = CallbackWrapperClassType.GetMethod("Create")

        MyCallbackObject = CallbackObjectFactory.Invoke(Nothing, New Object() {Me})

        Dim Wrapper As System.Delegate = _
        System.Delegate.CreateDelegate(EventHandlerType, MyCallbackObject, CallbackMethodName)

        Utils.AddMethod(EventProducerType, EventName).Invoke(TheObject, New Object() {Wrapper})

    End Sub

#Region "SetUp and TearDown stubs"

    Public Overridable Sub SetUp()
        ' Empty
    End Sub

    Public Overridable Sub TearDown()
        ' Empty
    End Sub

#End Region

#Region "Assertions"

    Public Sub AssertEqual(ByVal Expected As String, ByVal Actual As String)

        If Actual <> Expected Then
            EndTest(TestResult.AssertionFailed, MyCurrentCaseName & ": " & Expected & " != " & Actual)
        End If

    End Sub

    Public Sub AssertEqual(ByVal Expected As Boolean, ByVal Actual As Boolean)

        If Actual <> Expected Then
            EndTest(TestResult.AssertionFailed, MyCurrentCaseName & ": " & Expected.ToString & " != " & Actual.ToString)
        End If

    End Sub

    Public Sub AssertFail(Optional ByVal Explanation As String = "")

        EndTest(TestResult.Failure, Explanation)

    End Sub

#End Region

#End Region

    ' -----------------------------------------------------------------------------------------------------------

#Region "Timers"

    Private WithEvents TestTimeoutTimer As New System.Timers.Timer
    Private WithEvents ResetAndStartNextTestTimer As New System.Timers.Timer

#End Region

#Region "Constants"

    Private Const CallbackMethodName As String = "CallbackMethod"
    Private Const CallbackWrapperClassNamePrefix As String = "CallbackWrapperFor_"
    Private Const CallbackWrapperDelegateNamePrefix As String = "CallbackDelegateFor_"

#End Region

    Private WaitingForCallback As Boolean

    Private MyTestMethods As New Collection
    Private MyCallbackMethods As New Collection
    Private MyTestMethodsRemaining As New Collection

    Private MyCurrentCaseName As String
    Private MyMaxTestTimeSeconds As Integer = 2000

    Private MyCallbackMap As Hashtable = New Hashtable

    Private MyCallbackObject As Object

    Public Sub CallbackWrapper(ByVal Result As IAsyncResult)

        Dim RealCallback As AsyncCallback = DirectCast(MyCallbackMap.Item(MyCurrentCaseName), AsyncCallback)

        Try

            RealCallback.Invoke(Result)

        Catch ex As Exception

            EndTest(TestResult.ExceptionThrown, ex.Message)

        End Try

        EndTest()

    End Sub

    Public Sub PublicEndTest()

        EndTest(TestResult.Success)

    End Sub

    Public Sub PublicEndTestWithException(Optional ByVal Explanation As String = "")

        EndTest(TestResult.ExceptionThrown, Explanation)

    End Sub

    Protected Sub EndTest(Optional ByVal Result As TestResult = TestResult.Success, Optional ByVal Explanation As String = "")

        If MyCurrentCaseName = "" Then Return

        TestTimeoutTimer.Stop()
        WaitingForCallback = False

        RaiseEvent TestFinished(MyCurrentCaseName, Result, Explanation)

        Try
            TearDown()
        Catch ex As Exception
            Debug.WriteLine("TearDown threw an exception: " & ex.Message)
        End Try

        MyCurrentCaseName = ""
        ResetAndStartNextTestTimer.Start()

    End Sub

    Private Sub RunNextTest()

        If MyTestMethodsRemaining.Count = 0 Then
            RaiseEvent FinishedRunningTests()
            Exit Sub
        End If

        Dim TestMethod As MethodInfo = DirectCast(MyTestMethodsRemaining.Item(1), MethodInfo)
        MyTestMethodsRemaining.Remove(1)

        RunTest(TestMethod)

    End Sub

    Private Sub RunTest(ByVal TestMethod As MethodInfo)

        RaiseEvent StartingTest(TestMethod.Name)

        MyCurrentCaseName = TestMethod.Name
        WaitingForCallback = True

        TestTimeoutTimer.Start()

        RunSetup()
        RunTestMethod(TestMethod)

    End Sub

    Private Sub RunSetup()

        Try
            SetUp()
        Catch ex As Exception
            If ex.InnerException Is Nothing Then
                EndTest(TestResult.ExceptionThrownBySetup, ex.Message)
            Else
                EndTest(TestResult.ExceptionThrownBySetup, ex.InnerException.Message)
            End If
        End Try

    End Sub

    Private Sub RunTestMethod(ByVal TestMethod As MethodInfo)

        Try
            TestMethod.Invoke(Me, Nothing)
        Catch ex As Exception
            If ex.InnerException Is Nothing Then
                EndTest(TestResult.ExceptionThrown, ex.Message)
            Else
                EndTest(TestResult.ExceptionThrown, ex.InnerException.Message)
            End If
        End Try

    End Sub

    Private Sub Timeout(ByVal Sender As Object, ByVal Args As System.Timers.ElapsedEventArgs) Handles TestTimeoutTimer.Elapsed

        EndTest(TestResult.Timeout)

    End Sub

    Private Sub ResetAndStartNextTest(ByVal Sender As Object, ByVal Args As System.Timers.ElapsedEventArgs) Handles ResetAndStartNextTestTimer.Elapsed

        TearDown()
        RunNextTest()

    End Sub

    Private Sub InitTimers()

        TestTimeoutTimer.AutoReset = False
        TestTimeoutTimer.Interval = MaxTestTimeSeconds * 1000

        ResetAndStartNextTestTimer.AutoReset = False
        ResetAndStartNextTestTimer.Interval = 100

    End Sub

#Region "Event handling stuff"


    Private Function GenerateAssemblyForEventHandlerWrapperClass _
    ( _
        ByVal EventHandlerType As Type, _
        ByVal EventName As String, _
        ByVal RealCallbackName As String, _
        ByVal ProxyObject As Object _
    ) As System.Reflection.Assembly

        Dim EventHandlerInvokeMethod As MethodInfo = EventHandlerType.GetMethod("Invoke")

        Dim ParamSpecString As String = Utils.ParametersSpecAsString(EventHandlerInvokeMethod)
        Dim ParamNamesString As String = Utils.ParameterNamesAsString(EventHandlerInvokeMethod)

        Dim Code As String = CallbackClassDefinitionCode _
        ( _
            EventHandlerType.Name, _
            ParamSpecString, _
            ParamNamesString, _
            ProxyObject.GetType.FullName, _
            CallbackMethodName, _
            RealCallbackName _
        )

        Dim Refs() As String = {"System.dll", "Microsoft.VisualBasic.dll"}

        Return Utils.CompileToAssembly(Code, Refs)

    End Function

    Private Function CallbackClassDefinitionCode _
    ( _
        ByVal EventHandlerName As String, _
        ByVal ParameterSpec As String, _
        ByVal ParameterNames As String, _
        ByVal CallbackClassName As String, _
        ByVal CallbackMethodName As String, _
        ByVal RealCallbackName As String _
    ) As String

        Return _
        "imports System" & vbCrLf _
        & "public class " & CallbackWrapperClassName(EventHandlerName) & vbCrLf _
        & "Public shared CallbackObject as Object" & vbCrLf _
        & "public sub New(RealCallbackObject as Object)" & vbCrLf _
        & "  CallbackObject = RealCallbackObject" & vbCrLf _
        & "end sub" & vbCrLf _
        & "public shared function Create(RealCallbackObject as object) as " & CallbackWrapperClassName(EventHandlerName) & vbCrLf _
        & "  return new " & CallbackWrapperClassName(EventHandlerName) & "(RealCallbackObject)" & vbCrLf _
        & "end function" & vbCrLf _
        & "public sub " & CallbackMethodName & "(" & ParameterSpec & ")" & vbCrLf _
        & "  if callbackobject is nothing then" & vbCrLf _
        & "  end if" & vbCrLf _
        & "  dim CallbackMethod as system.reflection.MethodInfo = Callbackobject.GetType().GetMethod(""" & RealCallbackName & """)" & vbCrLf _
        & "  try" & vbCrLf _
        & "    callbackmethod.Invoke(CallbackObject, New Object() { " & ParameterNames & " })" & vbCrLf _
        & "  catch ex as Exception" & vbCrLf _
        & "    callbackobject.gettype.getmethod(""PublicEndTestWithException"").invoke(callbackobject, new Object() { ex.innerexception.Message })" & vbCrLf _
        & "  end try" & vbCrLf _
                & "  try" & vbCrLf _
        & "  callbackobject.gettype.getmethod(""PublicEndTest"").invoke(callbackobject, new Object() { })" & vbCrLf _
        & "  catch ex as Exception" & vbCrLf _
        & "    system.console.writeline(""Unexpected exception:"" & ex.message)" & vbCrLf _
        & "  end try" & vbCrLf _
        & "end sub" & vbCrLf _
        & "end class" & vbCrLf

    End Function

    Private Shared Function CallbackWrapperClassName(ByVal EventHandlerName As String) As String

        Return CallbackWrapperClassNamePrefix & Utils.DotsToDoubleUnderscores(EventHandlerName)

    End Function

    Private Shared Function CallbackWrapperDelegateName(ByVal EventHandlerName As String) As String

        Return CallbackWrapperDelegateNamePrefix & Utils.DotsToDoubleUnderscores(EventHandlerName)

    End Function

#End Region

End Class


