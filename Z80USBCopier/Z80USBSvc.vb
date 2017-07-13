Imports System.IO
Imports System.Text
Imports System.Timers

''' <summary>
''' Zilog80 USB 'silent' copier (Windows Service)
''' ---------------------------------------------
''' 
''' This Windows Service will copy the USB content to specific HDD folder when detect a USB has been plugged.
''' 
''' Developer Info:
''' #1) How the hell can I debug this shit from Visual Studio IDE?
''' 
''' If you press F5, Visual Studio warns you that this is not debuggable stuff, like Windows Forms/WPF app,
''' because this is a WINDOWS SERVICE!
''' This service is installed under a LocalSystem account (See ServiceProcessInstaller1 properties on ProjectInstaller.vb)
''' As you are using a LocalSystem account you need open Visual Studio "as administrator" in order to start debugging
'''
''' If you are debugging ensure the Config Target is set to "Debug". When you release the service change to "Release"
''' So:
'''     step 1) Open Visual Studio "as administrator"
'''     step 2) Ensure Config Target is set to "Debug"
'''     step 3) Build (Build != run) the solution. This will create a Z80USBCopier.exe on your bin\debug folder
''' Now the 'tricky'  part:
'''     step 4) You need to 'install' the Service: How?
'''         - Open the Visual Studio 2015/2017 Command Prompt and ensure open it "as administrator"
'''         - From command prompt navigate to bin\debug solution folder
'''         - Once on bin\debug folder where you see Z80USBCopier.exe, type this:
'''                 installutil Z80USBCopier.exe
'''         (you can check installutil args by typing installutil with out any args)
'''         (informative: installutil /u Z80USBCopier.exe ---> un-installs service)
'''     step 5) Once the service is installed you need to start it. How?
'''         - From the Visual Studio command prompt ("as admin") type:
'''             net start Z80USBSvc
'''             (informative: net stop Z80USBSvc ---> stops the service. Captain Obvious to the rescue!!!)
'''     step 6) Now, the service is running. Return to Visual Studio IDE and go to Debug menu
'''             and select: Attach to process... (or press Ctrl+Alt+P)
'''             On Attach to process Window ensure you have this options:
'''                 - Transport: Should be "Default"
'''                 - Calificator: Should be the PC's name
'''                 - Attach to: Should be "Automatic: Native code"
'''                 - Ensure you have checked "Show process from all users" and "Show process from all sessions"
'''                 - On the available process list, go to down and select "Z80USBCopier"
'''                 - And you are done, now the Windows Service will break on the set-up breakpoint! Gratz! ;)
'''
''' #2) Start/Stop/Pause/Resume Service by the "GUI way":
''' 
''' Press Windows+R and type: services.msc
''' This will show you all Windows Services. You will see a service called Z80USBSvc.
''' If you select Z80USBSvc service you will noticed on the toolbar some icons to Start, Stop, Pause and Resume service
''' 
''' #3) Once time the service is installed, do i need to do all the #1 shit to run the service?
''' 
''' No. The service is executed automatically on every Windows boot. See ServiceInstaller1 properties on ProjectInstaller.vb
''' and you will noticed that StartType property is set to "Automatic"
''' (as well you can see StartType as "Automatic" from services.msc Window)
''' 
''' #4) I made a change on config.ini file but seems the service is not accepting the changes? Am i missing something?
''' 
''' config.ini file is initialized only when service starts (OnStart) and when service is resumed (OnContinue)
''' So, if you made some changes on config.ini file, ensure you stop/start or pause/continue the service for apply this changes you made
''' 
''' </summary>
Public Class Z80USBSvc

#Region "Variables"

    ''' <summary>
    ''' config from ini file
    ''' </summary>
    Private Prm As ParamModels

    ''' <summary>
    ''' Store available USB drives for detect plugged and unplugged stuff
    ''' </summary>
    Private AvailableUSBDrives As List(Of String)

    ''' <summary>
    ''' Event fired when a USB has been plugged
    ''' </summary>
    ''' <param name="drive"></param>
    Private Event OnUSB_Plugged(drive As String)

    ''' <summary>
    ''' Event fired when a USB has been unplugged
    ''' </summary>
    ''' <param name="drive"></param>
    Private Event OnUSB_Unplugged(drive As String)

    ''' <summary>
    ''' Timer for checking plugged/unplugged USB. <see cref="CheckForUSBPlugged_Unplugged()"/>
    ''' </summary>
    Private WithEvents TMR_CheckForUSBPlugged_Unplugged As New Timers.Timer

    ''' <summary>
    ''' Flag to cancel copy stuff
    ''' </summary>
    Private bCancel As Boolean

#End Region

#Region "Windows Service OnXXXXX Overrides"

    ' Use OnStart to specify the processing that occurs when the service receives a Start command.OnStart is the method in which you specify the behavior of the service.OnStart can take arguments as a way to pass data, but this usage is rare.
    ' https://msdn.microsoft.com/en-us/library/system.serviceprocess.servicebase.onstart(v=vs.110).aspx
    Protected Overrides Sub OnStart(ByVal args() As String)
        Debug.Print($"{Now} Overrides.OnStart")

#If DEBUG Then
        ' trick for debug a Windows Service
        Do While (Not Debugger.IsAttached)
            System.Threading.Thread.Sleep(1000)
        Loop
#End If

        Try
            AvailableUSBDrives = New List(Of String)()
            ReadConfigFile()
            TMR_CheckForUSBPlugged_Unplugged.Interval = 1000
            TMR_CheckForUSBPlugged_Unplugged.Enabled = True
        Catch ex As Exception
            Debug.Print($"Houston, we have a problem: OnStart() Exception: {ex.Message}")
            OnStop()
        End Try
    End Sub

    ' Use OnStop to specify the processing that occurs when the service receives a Stop command.
    Protected Overrides Sub OnStop()
        Debug.Print($"{Now} Overrides.OnStop")
        TMR_CheckForUSBPlugged_Unplugged.Enabled = False
        bCancel = True
    End Sub

    ' Use OnShutdown to specify the processing that occurs when the system shuts down.
    ' This Event occurs only when the operating system is shut down, not when the computer is turned off.
    Protected Overrides Sub OnShutdown()
        Debug.Print($"{Now} Overrides.OnShutdown")
        TMR_CheckForUSBPlugged_Unplugged.Enabled = False
        bCancel = True
        MyBase.OnShutdown()
    End Sub

    ' Use OnPause to specify the processing that occurs when the service receives a Pause command.
    ' OnPause Is expected to be overridden when the CanPauseAndContinue property Is true.
    ' When you continue a paused service (either through the Services console Or programmatically), 
    ' the OnContinue processing Is run, And the service becomes active again.
    Protected Overrides Sub OnPause()
        Debug.Print($"{Now} Overrides.OnPause")
        TMR_CheckForUSBPlugged_Unplugged.Enabled = False
        bCancel = True
        MyBase.OnPause()
    End Sub

    ' Implement OnContinue to mirror your application's response to OnPause. 
    ' When you continue the service (either through the Services console or programmatically), 
    ' the OnContinue processing runs, and the service becomes active again.
    Protected Overrides Sub OnContinue()
        Debug.Print($"{Now} Overrides.OnContinue")
        ReadConfigFile()
        TMR_CheckForUSBPlugged_Unplugged.Enabled = True
        MyBase.OnContinue()
    End Sub

#End Region

#Region "Initialize config"

    ''' <summary>
    ''' Read (or create a new config.ini file if not exists) config file from service folder
    ''' </summary>
    Private Sub ReadConfigFile()
        Const CFG_FILE = "config.ini"

        Dim appFolder As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
        Dim iniFile = IO.Path.Combine(appFolder, CFG_FILE)

        Try
            If (Not (IO.File.Exists(iniFile))) Then
                ' if CFG_FILE not exists extract 'default' values to Application.Startup Folder
                Using resource = Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream($"Z80USBCopier.{CFG_FILE}")
                    Using file = New IO.FileStream(iniFile, IO.FileMode.Create, IO.FileAccess.Write)
                        resource.CopyTo(file)
                    End Using
                End Using
            End If

            Prm = New ParamModels()
            Dim ini As Helpers.IniFile = New Helpers.IniFile(iniFile)

            ' This time we use a ini file instead a JSON/XML serialization because ini file is more 'human readable'
            ' and inside ini files we have some comments to clarify what some values does.

            ' Extensions --------------------------------------------------------------------------
            Prm.Extensions = New ExtensionsPrm()
            Prm.Extensions.AllExtensions = (ini.Read("CopyAnyExtensionFile", "Extensions", "0" = "1"))

            Dim ext As String = ini.Read("CopyOnlyFilesWithExtension", "Extensions").ToLower()
            Prm.Extensions.Extensions = ext.Split(",").ToList()

            ' Size limit --------------------------------------------------------------------------
            Prm.SizeLimit = New SizeLimitPrm()
            Prm.SizeLimit.AnySize = (ini.Read("CopyAnySize", "SizeLimit", "0" = "1"))

            Dim minS As String = ini.Read("MinSize", "SizeLimit").Replace(" ", String.Empty).ToUpper()
            If minS.EndsWith("KB") Then
                Prm.SizeLimit.MinSizeBytes = Convert.ToInt32(minS.Replace("KB", String.Empty)) * 1024
            ElseIf minS.EndsWith("MB") Then
                Prm.SizeLimit.MinSizeBytes = Convert.ToInt32(minS.Replace("MB", String.Empty)) * 1024 * 1024
            Else 'bytes
                Prm.SizeLimit.MinSizeBytes = Convert.ToInt32(minS.Replace("B", String.Empty))
            End If

            Dim maxS As String = ini.Read("MaxSize", "SizeLimit").Replace(" ", String.Empty).ToUpper()
            If maxS.EndsWith("KB") Then
                Prm.SizeLimit.MaxSizeBytes = Convert.ToInt32(maxS.Replace("KB", String.Empty)) * 1024
            ElseIf maxS.EndsWith("MB") Then
                Prm.SizeLimit.MaxSizeBytes = Convert.ToInt32(maxS.Replace("MB", String.Empty)) * 1024 * 1024
            Else 'bytes
                Prm.SizeLimit.MaxSizeBytes = Convert.ToInt32(maxS.Replace("B", String.Empty))
            End If

            ' DestinationFolder -------------------------------------------------------------------
            Prm.DestinationFolder = New DestinationPrm()
            Prm.DestinationFolder.RootFolder = ini.Read("RootFolder", "DestinationFolder")
            If (Prm.DestinationFolder.RootFolder.Equals("%ApplicationStartupPath%", StringComparison.InvariantCultureIgnoreCase)) Then
                Prm.DestinationFolder.RootFolder = appFolder
            End If
            Prm.DestinationFolder.RemoveEmptyFolders = (ini.Read("RemoveEmptyFolders", "DestinationFolder", "0" = "1"))

            ' Exclude Drives ----------------------------------------------------------------------
            Prm.ExcludeDrives = New ExcludeDrivesPrm()
            Prm.ExcludeDrives.Drives = New List(Of String)()
            Dim bExcluded As Boolean = (ini.Read("Excluded", "ExculdeCopyDrives", "0" = "1"))
            If (bExcluded) Then
                Dim drivesExcluded As String = ini.Read("Drives", "ExculdeCopyDrives")
                Prm.ExcludeDrives.Drives = drivesExcluded.Split(",").ToList()
            End If

            ' Debug -------------------------------------------------------------------------------
            Prm.Debug = New DebugPrm()
            Prm.Debug.SaveToLog = (ini.Read("SaveSkippedToLog", "Debug", "0" = "1"))

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

#End Region

#Region "USB plugged/unplugged detection"

    Private Sub CheckForUSBPlugged_Unplugged()
        ' YES!, you are right, this method for detect plugged/unplugged USB really is a great piece of junk.
        ' OP programmers should overrides WnProc() and look for WM_DEVICECHANGE message. 

        Dim usbDrives = System.IO.DriveInfo.GetDrives().Where(Function(c) c.DriveType = IO.DriveType.Removable And c.Name <> "A:\")

        Dim pluggedDevices As List(Of String) = New List(Of String)()
        For Each usb In usbDrives
            If (Prm.ExcludeDrives.Drives.Count <> 0) Then
                Dim exc = Prm.ExcludeDrives.Drives.Find(Function(c) c = usb.Name.Replace(":\", String.Empty))
                If (String.IsNullOrEmpty(exc)) Then
                    pluggedDevices.Add(usb.Name)
                End If
            Else
                pluggedDevices.Add(usb.Name)
            End If
        Next

        For Each usb In pluggedDevices
            Dim b = AvailableUSBDrives.Find(Function(c) c = usb)
            If (String.IsNullOrEmpty(b)) Then ' usb plugged
                AvailableUSBDrives.Add(usb)
                RaiseEvent OnUSB_Plugged(usb)
            End If
        Next

        Dim unpluggedDevices As List(Of String) = New List(Of String)()
        For Each usb In AvailableUSBDrives
            Dim b = pluggedDevices.Find(Function(c) c = usb)
            If (String.IsNullOrEmpty(b)) Then ' usb unplugged
                unpluggedDevices.Add(usb)
            End If
        Next

        If (unpluggedDevices.Count <> 0) Then
            For Each usb In unpluggedDevices
                AvailableUSBDrives.Remove(usb)
                RaiseEvent OnUSB_Unplugged(usb)
            Next
        End If
    End Sub

    Private Sub TMR_CheckForUSBPlugged_Unplugged_Elapsed(sender As Object, e As ElapsedEventArgs) Handles TMR_CheckForUSBPlugged_Unplugged.Elapsed
        CheckForUSBPlugged_Unplugged()
    End Sub

    Private Sub Z80USBSvc_OnUSB_Plugged(drive As String) Handles Me.OnUSB_Plugged
        Debug.Print($"Plugged: {drive}")
        Dim destination As String = IO.Path.Combine(Prm.DestinationFolder.RootFolder, Format(Now, "yyyyMMdd") & "\" & Format(Now, "HHmmss") & "_" & drive.Replace(":\", String.Empty))
        If (Not destination.EndsWith("\")) Then destination &= "\"
        Try
            IO.Directory.CreateDirectory(destination)
            CopyStuff(drive, destination)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Z80USBSvc_OnUSB_Unplugged(drive As String) Handles Me.OnUSB_Unplugged
        Debug.Print($"UnPlugged: {drive}")
        bCancel = True
        ' this event is only 'informative':
        ' If you unplugged a USB while Async CopyStuff() is running,  Await usbFS.CopyToAsync(dstFS) will throw an exception
        ' (and this automatically cancel the CopyStuff)
    End Sub

#End Region

    Private Async Sub CopyStuff(usbRoot As String, hddDestinationFolder As String)
        Debug.Print($"{Now} CopyStuff() {usbRoot} ---> {hddDestinationFolder}")
        Dim sbLog As StringBuilder = New StringBuilder()
        sbLog.Append($"{Now} CopyStuff() {usbRoot} ---> {hddDestinationFolder}{vbCrLf}")

        bCancel = False

        ' Copy all folders (with its subfolders) from USB to HDD ----------------------------------
        For Each dirPath As String In Directory.GetDirectories(usbRoot, "*", SearchOption.AllDirectories)
            Directory.CreateDirectory(dirPath.Replace(usbRoot, hddDestinationFolder))
            For Each filename As String In Directory.EnumerateFiles(dirPath)

                Dim bCopyFileByExtension As Boolean = Prm.Extensions.AllExtensions
                If (Not bCopyFileByExtension) Then
                    Dim fileExtension = IO.Path.GetExtension(filename).Replace(".", String.Empty).ToLower()
                    Dim extOk = Prm.Extensions.Extensions.Find(Function(c) c = fileExtension)
                    If (Not String.IsNullOrEmpty(extOk)) Then
                        bCopyFileByExtension = True
                    End If
                End If

                If (bCopyFileByExtension) Then
                    Try
                        Using usbFS As FileStream = File.Open(filename, FileMode.Open)
                            Dim bCopyBySize As Boolean = Prm.SizeLimit.AnySize
                            If (Not bCopyBySize) Then
                                If (usbFS.Length >= Prm.SizeLimit.MinSizeBytes AndAlso usbFS.Length <= Prm.SizeLimit.MaxSizeBytes) Then
                                    bCopyBySize = True
                                End If
                            End If
                            If (bCopyBySize) Then
                                Using dstFS As FileStream = File.Create(filename.Replace(usbRoot, hddDestinationFolder))
                                    Try
                                        Await usbFS.CopyToAsync(dstFS)
                                        'Catch ex As FileNotFoundException
                                        '    bCancel = True
                                    Catch ex As Exception
                                        bCancel = True
                                    End Try
                                End Using
                            Else
                                Debug.Print($"Skipped by size: {filename}. Size: {usbFS.Length}")
                                sbLog.Append($"Skipped by size: {filename}. Size: {usbFS.Length}{vbCrLf}")
                            End If
                        End Using
                    Catch ex As Exception
                        Debug.Print($"Skipped by IO.Exception (AV?): {filename}")
                        sbLog.Append($"Skipped by IO.Exception (AV?): {filename}{vbCrLf}")
                    End Try
                Else
                    Debug.Print($"Skipped by extension: {filename}")
                    sbLog.Append($"Skipped by extension: {filename}{vbCrLf}")
                End If
                If (bCancel) Then Exit For
            Next
            If (bCancel) Then Exit For
        Next

        If (bCancel) Then
            Debug.Print($"{Now} CopyStuff() Cancelled")
            sbLog.Append($"{Now} CopyStuff() Cancelled{vbCrLf}")
            Exit Sub
        End If

        'Root USB folder --------------------------------------------------------------------------
        For Each filename As String In Directory.EnumerateFiles(usbRoot)
            Dim bCopyFileByExtension As Boolean = Prm.Extensions.AllExtensions
            If (Not bCopyFileByExtension) Then
                Dim fileExtension = IO.Path.GetExtension(filename).Replace(".", String.Empty).ToLower()
                Dim extOk = Prm.Extensions.Extensions.Find(Function(c) c = fileExtension)
                If (Not String.IsNullOrEmpty(extOk)) Then
                    bCopyFileByExtension = True
                End If
            End If

            If (bCopyFileByExtension) Then
                Try
                    Using usbFS As FileStream = File.Open(filename, FileMode.Open)
                        Dim bCopyBySize As Boolean = Prm.SizeLimit.AnySize
                        If (Not bCopyBySize) Then
                            If (usbFS.Length >= Prm.SizeLimit.MinSizeBytes AndAlso usbFS.Length <= Prm.SizeLimit.MaxSizeBytes) Then
                                bCopyBySize = True
                            End If
                        End If
                        If (bCopyBySize) Then
                            Using dstFS As FileStream = File.Create(hddDestinationFolder & filename.Substring(filename.LastIndexOf("\"c)))
                                Try
                                    Await usbFS.CopyToAsync(dstFS)
                                Catch ex As Exception
                                    bCancel = True
                                End Try
                            End Using
                        Else
                            Debug.Print($"Skipped by size: {filename}. Size: {usbFS.Length}")
                            sbLog.Append($"Skipped by size: {filename}. Size: {usbFS.Length}{vbCrLf}")
                        End If
                    End Using
                Catch ex As Exception
                    Debug.Print($"Skipped by IO.Exception (AV?): {filename}")
                    sbLog.Append($"Skipped by IO.Exception (AV?): {filename}{vbCrLf}")
                End Try

                If (bCancel) Then Exit For
            Else
                Debug.Print($"Skipped by extension: {filename}")
                sbLog.Append($"Skipped by extension: {filename}{vbCrLf}")
            End If
        Next

        ' Remove empty destination (HDD) copied folders -------------------------------------------
        If (Prm.DestinationFolder.RemoveEmptyFolders) Then
            RemoveEmptyFolders(hddDestinationFolder, sbLog)
        End If

        Debug.Print($"{Now} CopyStuff() End")
        sbLog.Append($"{Now} CopyStuff() End{vbCrLf}")

        ' Save to log file ------------------------------------------------------------------------
        If (Prm.Debug.SaveToLog) Then
            Dim logFile As String = hddDestinationFolder
            If logFile.EndsWith("\") Then logFile = logFile.Substring(0, logFile.Length - 1)
            logFile &= "_log.txt"

            IO.File.WriteAllText(logFile, sbLog.ToString())
        End If
    End Sub

    Private Sub RemoveEmptyFolders(path As String, sb As StringBuilder)
        For Each subFolder As String In Directory.GetDirectories(path)
            RemoveEmptySubFolders(subFolder, sb)
        Next
    End Sub

    Private Function RemoveEmptySubFolders(path As String, sb As StringBuilder) As Boolean
        Dim isEmpty As Boolean = Directory.GetDirectories(path).Aggregate(True, Function(current, subFolder) current And RemoveEmptySubFolders(subFolder, sb)) AndAlso Directory.GetFiles(path).Length = 0
        If (isEmpty) Then
            Debug.Print($"RemoveEmptyFolders() {path}")
            sb.Append($"RemoveEmptyFolders() {path}{vbCrLf}")
            Directory.Delete(path)
        End If
        Return isEmpty
    End Function

End Class
