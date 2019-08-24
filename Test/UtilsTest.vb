Imports NUnit.Framework

Public Class FakeEventHandler

End Class

<TestFixture()> Public Class UtilsTest

    Public Sub ParametersSpecTestMethodNoParameters()

        ' Does nothing.

    End Sub

    Public Sub ParametersSpecTestMethodOneParameter(ByVal ParameterOne As String)

        ' Does nothing.

    End Sub

    Public Sub ParametersSpecTestMethodTwoParameters(ByVal ParameterOne As String, ByVal ParameterTwo As Integer)

        ' Does nothing.

    End Sub

    Public Sub ParametersSpecTestMethodThreeParameters(ByVal ParameterOne As String, ByVal ParameterTwo As Integer, ByVal ParameterThree As Boolean)

        ' Does nothing.

    End Sub

    Public Sub add_FakeEventName(ByVal TheEventHandler As FakeEventHandler)

        ' Does nothing.

    End Sub

    Public Sub remove_FakeEventName()

        ' Does nothing.

    End Sub

    <Test()> Public Sub TestCompileToAssembly()

        Dim Code As String = "public class TestClass" & vbCrLf _
        & "public shared function TestMe() as string" & vbCrLf _
        & "return ""hello, world""" & vbCrLf _
        & "end function" & vbCrLf _
        & "end class" & vbCrLf

        Dim Refs() As String = {"System.dll"}

        Dim TheAssembly As Reflection.Assembly = Utils.CompileToAssembly(Code, Refs)

        Assert.IsNotNull(TheAssembly, "Couldn't compile code to assembly")

        Dim TestClassType As Type = TheAssembly.GetType("TestClass")

        Assert.IsNotNull(TestClassType, "Couldn't find TestClass type in assembly")

        Dim Text As String = CStr(TestClassType.GetMethod("TestMe").Invoke(Nothing, New Object() {}))

        Assert.AreEqual("hello, world", Text, "TestClass.TestMe didn't say hello")

    End Sub

    <Test()> Public Sub TestParametersSpecAsString()

        Dim NoParametersMethodInfo As Reflection.MethodInfo = _
        Me.GetType.GetMethod("ParametersSpecTestMethodNoParameters")

        Assert.AreEqual _
        ( _
            "", _
            AsyncTest.Utils.ParametersSpecAsString(NoParametersMethodInfo) _
        )

        Dim OneParameterMethodInfo As Reflection.MethodInfo = _
        Me.GetType.GetMethod("ParametersSpecTestMethodOneParameter")

        Assert.AreEqual _
        ( _
            "ParameterOne as System.String", _
            AsyncTest.Utils.ParametersSpecAsString(OneParameterMethodInfo) _
        )

        Dim TwoParametersMethodInfo As Reflection.MethodInfo = _
        Me.GetType.GetMethod("ParametersSpecTestMethodTwoParameters")

        Assert.AreEqual _
        ( _
            "ParameterOne as System.String, ParameterTwo as System.Int32", _
            AsyncTest.Utils.ParametersSpecAsString(TwoParametersMethodInfo) _
        )

        Dim ThreeParametersMethodInfo As Reflection.MethodInfo = _
        Me.GetType.GetMethod("ParametersSpecTestMethodThreeParameters")

        Assert.AreEqual _
        ( _
            "ParameterOne as System.String, ParameterTwo as System.Int32, ParameterThree as System.Boolean", _
            AsyncTest.Utils.ParametersSpecAsString(ThreeParametersMethodInfo) _
        )

    End Sub

    <Test()> Public Sub TestParameterNamesAsString()

        Dim NoParametersMethodInfo As Reflection.MethodInfo = _
        Me.GetType.GetMethod("ParametersSpecTestMethodNoParameters")

        Assert.AreEqual _
        ( _
            "", _
            AsyncTest.Utils.ParameterNamesAsString(NoParametersMethodInfo) _
        )

        Dim OneParameterMethodInfo As Reflection.MethodInfo = _
        Me.GetType.GetMethod("ParametersSpecTestMethodOneParameter")

        Assert.AreEqual _
        ( _
            "ParameterOne", _
            AsyncTest.Utils.ParameterNamesAsString(OneParameterMethodInfo) _
        )

        Dim TwoParametersMethodInfo As Reflection.MethodInfo = _
        Me.GetType.GetMethod("ParametersSpecTestMethodTwoParameters")

        Assert.AreEqual _
        ( _
            "ParameterOne, ParameterTwo", _
            AsyncTest.Utils.ParameterNamesAsString(TwoParametersMethodInfo) _
        )

        Dim ThreeParametersMethodInfo As Reflection.MethodInfo = _
        Me.GetType.GetMethod("ParametersSpecTestMethodThreeParameters")

        Assert.AreEqual _
        ( _
            "ParameterOne, ParameterTwo, ParameterThree", _
            AsyncTest.Utils.ParameterNamesAsString(ThreeParametersMethodInfo) _
        )

    End Sub

    <Test()> Public Sub TestDotsToDoubleUnderscores()

        Assert.AreEqual("", Utils.DotsToDoubleUnderscores(""))
        Assert.AreEqual("__", Utils.DotsToDoubleUnderscores("."))
        Assert.AreEqual("____", Utils.DotsToDoubleUnderscores(".."))
        Assert.AreEqual("a__b__c", Utils.DotsToDoubleUnderscores("a.b.c"))
        Assert.AreEqual("____bc", Utils.DotsToDoubleUnderscores("..bc"))
        Assert.AreEqual("bc____", Utils.DotsToDoubleUnderscores("bc.."))
        Assert.AreEqual("__", Utils.DotsToDoubleUnderscores("__"))
        Assert.AreEqual("______", Utils.DotsToDoubleUnderscores("__.__"))
        Assert.AreEqual("  __  __  ", Utils.DotsToDoubleUnderscores("  .  .  "))

    End Sub

    <Test()> Public Sub TestTypeOfEventHandler()

        'Assert.AreEqual(GetType(FakeEventHandler), Utils.TypeOfEventHandler(Me.GetType, "FakeEventName"))

    End Sub

    <Test()> Public Sub TestAddMethod()

        Assert.IsNull(Utils.AddMethod(Me.GetType, "BogusEventName"))

        Assert.AreEqual(Utils.AddMethod(Me.GetType, "FakeEventName"), Me.GetType.GetMethod("add_FakeEventName"))

    End Sub

    <Test()> Public Sub TestRemoveMethod()

        Assert.IsNull(Utils.RemoveMethod(Me.GetType, "BogusEventName"))

        Assert.AreEqual(Utils.RemoveMethod(Me.GetType, "FakeEventName"), Me.GetType.GetMethod("remove_FakeEventName"))

    End Sub

    <Test()> Public Sub TestFindMethodInfo()

        Dim NoParametersMethodInfo As Reflection.MethodInfo = _
        Utils.FindMethodInfo("ParametersSpecTestMethodNoParameters", Me.GetType)

        Assert.IsNotNull(NoParametersMethodInfo, "Can't find method ParametersSpecTestMethodNoParameters")

        Assert.AreEqual(0, NoParametersMethodInfo.GetParameters.Length, "Should have no parameters")

        Dim OneParameterMethodInfo As Reflection.MethodInfo = _
        Utils.FindMethodInfo("ParametersSpecTestMethodOneParameter", Me.GetType)

        Assert.IsNotNull(OneParameterMethodInfo, "Can't find method ParametersSpecTestMethodOneParameter")

        Dim ParameterInfo As Reflection.ParameterInfo = _
        DirectCast(OneParameterMethodInfo.GetParameters.GetValue(0), Reflection.ParameterInfo)

        Assert.AreEqual("ParameterOne", ParameterInfo.Name, "First parameter's name is incorrect")

    End Sub

End Class
