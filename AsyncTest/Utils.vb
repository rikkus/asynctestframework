Imports System.Reflection
Imports System.CodeDom.Compiler

Public Class Utils

    Public Shared Function CompileToAssembly(ByVal Code As String, ByVal Refs() As String) As System.Reflection.Assembly

        Dim CodeProvider As New VBCodeProvider

        Dim Compiler As ICodeCompiler = CodeProvider.CreateCompiler

        Dim Params As New CompilerParameters

        Params.GenerateInMemory = True
        Params.TreatWarningsAsErrors = True

        Params.ReferencedAssemblies.AddRange(Refs)

        Dim Results As CompilerResults = Compiler.CompileAssemblyFromSource(Params, Code)

        If Results.Errors.Count > 0 Then
            For Each TheError As CompilerError In Results.Errors
                Debug.WriteLine(TheError.ErrorText)
            Next
            Return Nothing
        End If

        Return Results.CompiledAssembly

    End Function

    Public Shared Function ParametersSpecAsString(ByVal Info As MethodInfo) As String

        Dim Ret As String = ""

        For Each Param As ParameterInfo In Info.GetParameters

            If Ret <> "" Then
                Ret += ", "
            End If

            Ret += Param.Name & " as " & Param.ParameterType.FullName

        Next

        Return Ret

    End Function

    Public Shared Function ParameterNamesAsString(ByVal Info As MethodInfo) As String

        Dim Ret As String = ""

        For Each Param As ParameterInfo In Info.GetParameters

            If Ret <> "" Then
                Ret += ", "
            End If

            Ret += Param.Name

        Next

        Return Ret

    End Function

    Public Shared Function DotsToDoubleUnderscores(ByVal Text As String) As String

        Return Text.Replace(".", "__")

    End Function

    Public Shared Function TypeOfEventHandler(ByVal TheType As Type, ByVal EventName As String) As Type

        Return AddMethod(TheType, EventName).GetParameters()(0).ParameterType

    End Function

    Public Shared Function AddMethod(ByVal TheType As Type, ByVal EventName As String) As MethodInfo

        Return FindMethodInfo("add_" & EventName, TheType)

    End Function

    Public Shared Function RemoveMethod(ByVal TheType As Type, ByVal EventName As String) As MethodInfo

        Return FindMethodInfo("remove_" & EventName, TheType)

    End Function

    Public Shared Function FindMethodInfo(ByVal Name As String, ByVal TheType As Type) As MethodInfo

        Dim InfoList() As MethodInfo = TheType.GetMethods

        For Each Info As MethodInfo In InfoList

            If Info.Name = Name Then Return Info

        Next

        Return Nothing

    End Function

End Class
