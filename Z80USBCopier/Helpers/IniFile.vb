Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Reflection
Imports System.Text

Namespace Helpers

    Public Class IniFile

#Region "Win API"

        <DllImport("kernel32")>
        Private Shared Function WritePrivateProfileString(Section As String, Key As String, Value As String, FilePath As String) As Boolean
        End Function

        <DllImport("kernel32")>
        Private Shared Function GetPrivateProfileString(Section As String, Key As String, [Default] As String, RetVal As StringBuilder, Size As Integer, FilePath As String) As Integer
        End Function

#End Region

#Region "Variables"
        Private Path As String
        Private EXE As String = Assembly.GetExecutingAssembly().GetName().Name
#End Region

#Region "Constructor"

        Public Sub New(Optional IniPath As String = Nothing)
            Path = New FileInfo(If(IniPath, EXE & Convert.ToString(".ini"))).FullName.ToString()
        End Sub

#End Region

        Public Function Read(key As String, Optional section As String = Nothing, Optional [default] As String = Nothing) As String
            Dim RetVal = New StringBuilder(255)
            GetPrivateProfileString(If(section, EXE), key, [default], RetVal, 255, Path)
            Return RetVal.ToString()
        End Function

        Public Sub Write(key As String, value As String, Optional section As String = Nothing)
            Dim ret As Boolean = WritePrivateProfileString(If(section, EXE), key, value, Path)
            If (Not ret) Then
                Throw New System.Exception(String.Format("Something went wrong writting value {0} on file {1}. The folder is OK? The file is read/only?", key, Path))
            End If
        End Sub

        Public Sub DeleteKey(key As String, Optional section As String = Nothing)
            Write(key, Nothing, If(section, EXE))
        End Sub

        Public Sub DeleteSection(Optional section As String = Nothing)
            Write(Nothing, Nothing, If(section, EXE))
        End Sub

        Public Function KeyExists(key As String, Optional section As String = Nothing) As Boolean
            Return Read(key, section).Length > 0
        End Function

    End Class

End Namespace
