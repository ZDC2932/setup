Imports System.Net
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.IO
Imports System.Diagnostics
Imports System.Drawing.Text

Public Class Form1
    Dim currentDirectory As String = My.Computer.FileSystem.CurrentDirectory

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/posetup.msi", Path.Combine(currentDirectory, "posetup.msi"))
        DownloadFile("https://open-pengshen.oss-cn-qingdao.aliyuncs.com/static/tools/阿里巴巴普惠体 R.ttf", Path.Combine(currentDirectory, "阿里巴巴普惠体 R.ttf"))
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

        MessageBox.Show("遥遥领先，下载成功！！！")

        ' 字体文件路径
        Dim fontFilePath As String = Path.Combine(currentDirectory, "阿里巴巴普惠体 R.ttf")

        ' 安装字体
        InstallFont(fontFilePath)
    End Sub


    Public Sub check_version(ByVal filepath As String)


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

    Public Function runCmd(ByVal strCMD As String) As String
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

    Sub InstallFont(fontFilePath As String)
        If File.Exists(fontFilePath) Then
            Try
                ' 创建字体集合对象
                Dim fontCollection As New PrivateFontCollection()

                ' 添加字体文件到字体集合
                fontCollection.AddFontFile(fontFilePath)

                ' 获取字体族名称
                Dim fontFamilyName As String = fontCollection.Families(0).Name

                ' 创建注册表项
                Dim regKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", True)

                ' 设置字体注册表项的值
                regKey.SetValue(fontFamilyName, Path.GetFileName(fontFilePath), Microsoft.Win32.RegistryValueKind.String)

                ' 刷新字体缓存
                SendMessageTimeout(New IntPtr(HWND_BROADCAST), WM_FONTCHANGE, IntPtr.Zero, IntPtr.Zero, SendMessageTimeoutFlags.SMTO_NORMAL, 1000, Nothing)

                Console.WriteLine("字体安装成功")
            Catch ex As Exception
                Console.WriteLine("字体安装失败：" & ex.Message)
            End Try
        Else
            Console.WriteLine("字体文件不存在")
        End If
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
End Class
