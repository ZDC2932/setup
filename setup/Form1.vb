Imports System.Net
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.IO
Imports System.Diagnostics
Imports System.Drawing.Text
Imports Microsoft.Win32
Imports System.Configuration
Imports System.Security.Cryptography.X509Certificates
Imports Microsoft.Office.Interop.Excel
Imports System.Net.Http.Headers


Public Class Form1
    '主要配置变量
    Public ConfigFilepath As String
    Public winver As Integer
    Public winname As String
    Public IEinstalled As Boolean
    Public IE_version As String
    Public Pageinstalled As Boolean
    Public Page_version As String
    Public Google_installed As Boolean
    Public Google_version As String
    Public Excelinstalled As Boolean
    Public Excel_version As String
    Public Wpsinstalled As Boolean
    Public Wps_version As String
    Public Wps_rootpath As String
    Public Excel_rootpath As String
    Dim currentDirectory As String = My.Computer.FileSystem.CurrentDirectory

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Checkpage()
        Excelck()
        Wpsck()
        Ggck()
        Renewstat()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/阿里巴巴普惠体 R.ttf", Path.Combine(currentDirectory, "阿里巴巴普惠体 R.ttf"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/s1.bat", Path.Combine(currentDirectory, "s1.bat"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/s2.bat", Path.Combine(currentDirectory, "s2.bat"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/s3.bat", Path.Combine(currentDirectory, "s3.bat"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/s4.bat", Path.Combine(currentDirectory, "s4.bat"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/s5.bat", Path.Combine(currentDirectory, "s5.bat"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/s63.bat", Path.Combine(currentDirectory, "s63.bat"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/s7.bat", Path.Combine(currentDirectory, "s7.bat"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/99.ico", Path.Combine(currentDirectory, "99.ico"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/linkcv3.vbs", Path.Combine(currentDirectory, "linkcv3.vbs"))
        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/鹏燊云平台3.bat", Path.Combine(currentDirectory, "鹏燊云平台3.bat"))
        Config()
        MessageBox.Show("遥遥领先，执行成功！！！")

    End Sub


    'Public Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.MouseDown

    '    RichTextBox1.Text = "主版本号 ：" & GetOSVersion() & vbNewLine & "版本名称：" & GetOSproductName()
    '    RichTextBox1.AppendText(vbNewLine & "page_version：" & Checkpage())
    '    Excelck()
    '    RichTextBox1.AppendText(vbNewLine & "excelinstalled：" & Excelinstalled)
    '    RichTextBox1.AppendText(vbNewLine & "excel_version：" & Excel_version)
    '    Ggck()
    '    RichTextBox1.AppendText(vbNewLine & "googleinstalled：" & Google_installed)
    '    RichTextBox1.AppendText(vbNewLine & "google_version：" & Google_version)
    '    Wpsck()
    '    RichTextBox1.AppendText(vbNewLine & "Wpsinstalled：" & Wpsinstalled)
    '    RichTextBox1.AppendText(vbNewLine & "Wps_version：" & Wps_version)
    '    RichTextBox1.AppendText(vbNewLine & "Wps_rootpath：" & Wps_rootpath)
    'End Sub

    'ST1 生成配置文件 对主要变量赋值

    Public Sub Renewstat()
        If Pageinstalled = True Then
            inpage.Text = "重新安装"
        End If
        If Google_installed = True Then
            inchrome.Text = "重新安装"
        End If
        If Excelinstalled = True Then
            inoffice.Text = "重新安装"
        End If
        If Wpsinstalled = True Then
            inwps.Text = "重新安装"
        End If
        If IEinstalled = True Then
            inIE.Text = "清理缓存"
        End If
    End Sub
    Public Sub Config()

        winver = GetOSVersion()
        winname = GetOSproductName()


        Pageinstalled = False
        Page_version = Checkpage()
        Google_installed = False
        Google_version = 1
        Excelinstalled = False
        Excel_version = 1
        Wpsinstalled = False
        Wps_version = 11


        '  使用  File.AppendAllText  方法将内容追加到文件   


        'File.AppendAllText(Filepath, content)

        'RichTextBox1.Text = GetOSVersion() & GetOSproductName()

        'GetOSVersion()
        'Getwinver()


        '检查IE浏览器

        '检查page

        '检查谷歌浏览器

        '检查Excel

        '检查WPS
        ConfigFilepath = Path.Combine(currentDirectory, "config.ini")  '指定要写入的文件路径   
        Dim content As String = "PC版本号：" & GetOSVersion() & vbNewLine & "PC产品名称：" & GetOSproductName()  '要写入的内容
        File.WriteAllText(ConfigFilepath, content)

    End Sub
    'page 检查方法
    Public Function Checkpage()
        Dim Pv As String
        Pv = 0
        Pageinstalled = False
        Dim Pagepath = "C:\Program Files (x86)\Zhuozhengsoft\PageOfficeClient\POBrowser.exe"
        If File.Exists(Pagepath) Then
            If Getappver(Pagepath) > 0 Then
                Pageinstalled = True
                Pv = Getappver(Pagepath)
            End If
        Else
            Pagepath = "C:\Program Files\Zhuozhengsoft\PageOfficeClient\POBrowser.exe"
            If File.Exists(Pagepath) Then
                If Getappver(Pagepath) > 0 Then
                    Pv = Getappver(Pagepath)
                End If

            End If
        End If
        Page_version = Pv

    End Function
    '检查谷歌浏览器
    Public Sub Ggck()
        ' 检查谷歌浏览器是否已安装
        Dim chromeKey As String = "Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"
        Dim chromePath As String = Registry.GetValue("HKEY_LOCAL_MACHINE\" & chromeKey, "", Nothing)

        If chromePath IsNot Nothing Then
            ' 获取谷歌浏览器版本信息
            Getappver(chromePath)
            Console.WriteLine(Getappver(chromePath))
            If IO.File.Exists(chromePath) Then
                Google_version = Getappver(chromePath)
                Google_installed = True

            Else
                Console.WriteLine("无法获取谷歌浏览器版本信息")
                Google_version = 0
                Google_installed = True

            End If
        Else
            Console.WriteLine("谷歌浏览器未安装")
            Google_installed = False
        End If
    End Sub
    '检查wps
    Public Sub Wpsck()
        ' 检查 WPS Office 是否已安装
        '计算机\HKEY_CURRENT_USER\Software\Kingsoft\Office\6.0\wpsoffice\Application Settings
        Dim wpsKey As String = "Software\Kingsoft\Office\6.0\wpsoffice\Application Settings"
        Dim Wpsversion As String = Registry.GetValue("HKEY_CURRENT_USER\" & wpsKey, "version", Nothing)
        Wps_rootpath = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Kingsoft\Office\6.0\Common", "InstallRoot", Nothing)

        If Wpsversion IsNot Nothing Then
            ' 获取 WPS Office 版本信息
            'Dim versionInfoPath As String = IO.Path.Combine(wpsPath, "office6\wps.ini")
            Wps_version = Wpsversion
            Wpsinstalled = True

        Else
            Wps_version = 0
            Wpsinstalled = False

            Console.WriteLine("WPS Office 未安装")
        End If
    End Sub

    Sub Excelck()
        ' 创建 Excel 应用程序对象
        'Dim excelApp As New Application()

        '' 检查 Excel 是否已安装
        'If excelApp Is Nothing Then
        '    Excelinstalled = False
        '    Console.WriteLine("Excel 未安装！")
        'Else
        '    ' 获取 Excel 版本信息
        '    Excelinstalled = True
        '    Excel_version = excelApp.Version & " (com)"
        '    Console.WriteLine("Excel 版本: " & Excel_version)
        'End If
        '' 关闭 Excel 应用程序
        'excelApp.Quit()
        Dim Excelreg As String = "SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot"
        Excel_rootpath = Registry.GetValue("HKEY_LOCAL_MACHINE\" & Excelreg, "Path", Nothing)
        If Excel_rootpath IsNot Nothing Then
            Excel_version = Getappver(Excel_rootpath & "excel.exe")
            Excelinstalled = True
        Else
            Excel_version = 0
            Excelinstalled = False

        End If





    End Sub

    '检查ie是否安装
    Sub IEck（）
        IE_version = Getappver("C:\Program Files\Internet Explorer\iexplore.exe")

    End Sub
    '获取exe版本的方法
    Public Function Getappver(exePath As String)
        Dim fileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(exePath)  '获取文件版本信息   
        Dim exever As String = fileVersionInfo.ProductVersion
        Dim index As Integer = exever.IndexOf(".")
        Dim outstr As String
        If index <> -1 Then
            outstr = exever.Substring(0, index)
        Else
            outstr = "No match"
        End If

        Return outstr  '返回产品版本信息 
    End Function


    '版本判断，大于win10
    'for /f "tokens=4,5 delims=. " %%a in ('ver') do if %%a%%b geq 60 goto new

    '安装字体
    Public Sub Fontinstall()
        Dim fontName As String = "阿里巴巴普惠体 R"
        If CheckFontExists(fontName) Then
            Console.WriteLine($"字体  '{fontName}'  存在。")
        Else
            Console.WriteLine($"字体  '{fontName}'  不存在。")
            DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/阿里巴巴普惠体 R.ttf", Path.Combine(currentDirectory, "阿里巴巴普惠体 R.ttf"))
            StartLocalProcess(Path.Combine(currentDirectory, "installfont.bat"))
        End If
    End Sub

    '安装page
    Public Sub Installpage()
        '下载file，如果存在可以不下载
        If File.Exists(Path.Combine(currentDirectory, "posetup.msi")) Then
            StartLocalProcess(Path.Combine(currentDirectory, "Installpage.bat"))
        Else
            'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/posetup.msi", Path.Combine(currentDirectory, "posetup.msi"))
            Console.WriteLine(currentDirectory)
            StartLocalProcess(Path.Combine(currentDirectory, "Installpage.bat"))
        End If

        'Dim Pagecmd As String = "msiexec /i " & Path.Combine(currentDirectory, "posetup.msi") & "/quiet"
        'runCmd(Pagecmd)

    End Sub

    '安装谷歌
    Public Sub Installgoogle()
        '安装谷歌浏览器
        If File.Exists(Path.Combine(currentDirectory, "ChromeSetup.exe")) And Google_installed = False Then
            StartLocalProcess(Path.Combine(currentDirectory, "ChromeSetup.exe"))
        ElseIf Google_installed = False And File.Exists(Path.Combine(currentDirectory, "ChromeSetup.exe")) = False Then
            DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/ChromeSetup.exe", Path.Combine(currentDirectory, "ChromeSetup.exe"))
            StartLocalProcess(Path.Combine(currentDirectory, "ChromeSetup.exe"))
        End If

    End Sub
    '安装excel
    Public Sub InstallExcle()


    End Sub
    '安装ie
    Public Sub InstallIE()
        If IEinstalled = False Then
            StartLocalProcess(Path.Combine(currentDirectory, "Installie.bat"))
        End If

    End Sub

    '获取系统的版本
    Public Function GetOSVersion() As String
        Dim keyPath As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(keyPath)
            Dim majorVersion As Integer = Integer.Parse(key.GetValue("CurrentMajorVersionNumber").ToString())
            Return $" {majorVersion}"
        End Using
    End Function

    '获取Windows产品名称
    Public Function GetOSproductName() As String
        Dim keyPath As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(keyPath)
            Dim productName As String = key.GetValue("ProductName").ToString()
            Return $" {productName}"
        End Using
    End Function
    '获取操作系统名称
    Public Sub Getwinver()
        Dim osMajor As Integer = GetOSVersion()
        Select Case osMajor

            Case 6
                Console.WriteLine("操作系统：Windows Vista 或 Windows Server 2008")

            Case 6.1
                Console.WriteLine("操作系统：Windows 7 或 Windows Server 2008 R2")

            Case 6.2
                Console.WriteLine("操作系统：Windows 8 或 Windows Server 2012")

            Case 6.3
                Console.WriteLine("操作系统：Windows 8.1 或 Windows Server 2012 R2")

            Case 10
                Console.WriteLine("操作系统：Windows 10 或 Windows Server 2016 或 Windows Server 2019")

            Case 11
                Console.WriteLine("操作系统：Windows 11")

            Case Else

                Console.WriteLine("操作系统：未知")

        End Select

    End Sub


    Sub Wtf(message As String)
        Dim logDirectory As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs")
        Dim logFile As String = Path.Combine(logDirectory, "log.txt")

        ' 检查logs目录是否存在，如果不存在则创建
        If Not Directory.Exists(logDirectory) Then
            Directory.CreateDirectory(logDirectory)
        End If

        ' 使用StreamWriter写入log文件
        Using writer As New StreamWriter(logFile, True)
            writer.WriteLine(DateTime.Now.ToString() & " - " & message)
        End Using
    End Sub


    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub
    Private Sub DownloadFile(ByVal fileUrl As String, ByVal fileName As String)
        Using client As New WebClient()
            Try
                client.DownloadFile(fileUrl, fileName)
                Console.WriteLine("文件下载成功！")
            Catch ex As Exception
                Console.WriteLine("文件下载失败：" & ex.Message)
            End Try
        End Using
    End Sub


    Public Function CheckFontExists(ByVal fontName As String) As Boolean
        Dim fontFamilies As FontFamily() = FontFamily.Families
        For Each fontFamily As FontFamily In fontFamilies
            If fontFamily.Name.Equals(fontName, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    Sub ExtractAndWriteVersion(filePath As String)
        Try
            Dim versionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(filePath)
            Dim softwareName As String = versionInfo.ProductName
            Dim softwareVersion As String = versionInfo.ProductVersion

            Dim output As String = $"{softwareName}: {softwareVersion}"
            File.WriteAllText("output.txt", output)

            Console.WriteLine("Version information written to output.txt")
        Catch ex As Exception
            Console.WriteLine($"Error: {ex.Message}")
        End Try
    End Sub

    Public Function Runcmd(ByVal strCMD As String) As String
        Dim p As New Process
        With p.StartInfo
            .FileName = "cmd.exe"
            .Arguments = "/c " + strCMD
            .UseShellExecute = False
            .RedirectStandardInput = True
            .RedirectStandardOutput = True
            .RedirectStandardError = True
            .CreateNoWindow = True

        End With
        p.Start()
        Dim result As String = p.StandardOutput.ReadToEnd()
        p.Close()
        Return result

    End Function


    '这是程序启动的方法

    Public Sub StartLocalProcess(ByVal exeName As String)
        Try
            Dim process As New Process()
            process.StartInfo.FileName = exeName
            process.StartInfo.UseShellExecute = False
            process.StartInfo.CreateNoWindow = True

            process.Start()
            process.WaitForExit()

            Dim exitCode As Integer = process.ExitCode
            If exitCode <> 0 Then
                Console.WriteLine($"Error: {exitCode}")
            Else
                Console.WriteLine($"Process exited with code {exitCode}")
            End If
        Catch ex As Exception
            Console.WriteLine($"Error: unable to start process: {ex.Message}")
        End Try
    End Sub


    Private Sub CopyWithProgress(ByVal ParamArray filenames As String())
        ' Display the ProgressBar control.
        ProgressBar1.Visible = True
        ' Set Minimum to 1 to represent the first file being copied.
        ProgressBar1.Minimum = 0
        ' Set Maximum to the total number of files to copy.
        ProgressBar1.Maximum = 100
        ' Set the initial value of the ProgressBar.
        ProgressBar1.Value = 1
        ' Set the Step property to a value of 1 to represent each file being copied.
        ProgressBar1.Step = 1

        ' Loop through all files to copy.
        Dim x As Integer
        For x = 1 To 100

            ProgressBar1.PerformStep()

        Next x
        'runCmd("echo 输出文件成功！！！>'%homepath%\desktop\检测报告.txt'")

        MessageBox.Show("安装成功！！！")
    End Sub



    ' Windows消息常量
    Const HWND_BROADCAST As Integer = &HFFFF
    Const WM_FONTCHANGE As Integer = &H1D
    Const SMTO_NORMAL As Integer = &H0

    ' Windows API函数
    Declare Function SendMessageTimeout Lib "user32.dll" Alias "SendMessageTimeoutW" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByVal fuFlags As SendMessageTimeoutFlags, ByVal uTimeout As Integer, ByRef lpdwResult As IntPtr) As IntPtr

    ' SendMessageTimeout标志位枚举
    Enum SendMessageTimeoutFlags As Integer
        SMTO_NORMAL = 0
        SMTO_BLOCK = 1
        SMTO_ABORTIFHUNG = 2
        SMTO_NOTIMEOUTIFNOTHUNG = 8
    End Enum

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles inpage.LinkClicked, inoffice.LinkClicked, inwps.LinkClicked, inchrome.LinkClicked, inIE.LinkClicked

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles pengshen.Click, Label1.Click, Label3.Click, Label2.Click, Label4.Click

    End Sub
End Class
