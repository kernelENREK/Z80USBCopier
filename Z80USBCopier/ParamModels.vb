Public Class ParamModels
    Public Extensions As ExtensionsPrm
    Public SizeLimit As SizeLimitPrm
    Public DestinationFolder As DestinationPrm
    Public ExcludeDrives As ExcludeDrivesPrm
    Public Debug As DebugPrm
End Class

Public Class ExtensionsPrm
    ''' <summary>
    ''' Copy regardless of file extension. If its false only files with <see cref="Extensions"/> list will be copied
    ''' </summary>
    ''' <returns></returns>
    Public Property AllExtensions As Boolean
    ''' <summary>
    ''' Only copy files with this extensions. Not applicable if <see cref="AllExtensions"/> is set to true
    ''' </summary>
    ''' <returns></returns>
    Public Property Extensions As List(Of String)
End Class

Public Class SizeLimitPrm
    ''' <summary>
    ''' Copy regardless of file size. If its false only files with size match <see cref="MinSizeBytes"/> and <see cref="MaxSizeBytes"/> will
    ''' be copied
    ''' </summary>
    ''' <returns></returns>
    Public Property AnySize As Boolean
    ''' <summary>
    ''' Minimun file size (in bytes) for being copying. Not applicable if <see cref="AnySize"/> is set to true 
    ''' </summary>
    ''' <returns></returns>
    Public Property MinSizeBytes As Integer
    ''' <summary>
    ''' Maximun file size (in bytes) for being copying. Not applicable if <see cref="AnySize"/> is set to true
    ''' </summary>
    ''' <returns></returns>
    Public Property MaxSizeBytes As Integer
End Class

Public Class DestinationPrm
    ''' <summary>
    ''' Root folder from HDD where USB content will be copied.
    ''' The hierarchy will be: RootFolder\yyyyMMdd\HHmmss_[usb_drive]\
    ''' </summary>
    ''' <returns></returns>
    Public Property RootFolder As String
    ''' <summary>
    ''' Once the copy is completed remove all destination empty folders
    ''' </summary>
    ''' <returns></returns>
    Public Property RemoveEmptyFolders As Boolean
End Class
Public Class ExcludeDrivesPrm
    ''' <summary>
    ''' Drive letters (without :\) excluded from plugged stuff
    ''' </summary>
    ''' <returns></returns>
    Public Property Drives As List(Of String)
End Class

Public Class DebugPrm
    ''' <summary>
    ''' Save info (skipped files and deleted empty folders) to RootFolder\yyyyMMdd\HHmmss_[usb_drive]_log.txt file
    ''' </summary>
    ''' <returns></returns>
    Public Property SaveToLog As Boolean
End Class

