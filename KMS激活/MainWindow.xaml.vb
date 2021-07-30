Imports System.ComponentModel
Imports System.IO
Class MainWindow
	Private 密钥列表 As String()
	ReadOnly 设置目录 As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
	ReadOnly 设置路径 As String = Path.Combine(设置目录, "KMS激活设置.bin")
	Sub 读入设置()
		Dim 读入器 As BinaryReader
		Try
			读入器 = New BinaryReader(File.OpenRead(设置路径))
		Catch ex As Exception
			写出设置()
			Return
		End Try
		Try
			Width = 读入器.ReadUInt16
			Height = 读入器.ReadUInt16
			KMS服务器地址.Text = 读入器.ReadString
			Dim 版本选项 As SByte = 读入器.ReadSByte
			If 版本选项 < 0 Then
				版本或密钥.Text = 读入器.ReadString
			Else
				版本或密钥.SelectedIndex = 版本选项
			End If
		Catch
			读入器.Close()
			写出设置()
			Return
		End Try
		读入器.Close()
	End Sub
	Sub 写出设置()
		Call Directory.CreateDirectory(设置目录)
		Using 写出器 As New BinaryWriter(File.OpenWrite(设置路径))
			写出器.Write(CUShort(Width))
			写出器.Write(CUShort(Height))
			写出器.Write(KMS服务器地址.Text)
			写出器.Write(CSByte(版本或密钥.SelectedIndex))
			写出器.Write(版本或密钥.Text)
		End Using
	End Sub
	Sub New()

		' This call is required by the designer.
		InitializeComponent()

		' Add any initialization after the InitializeComponent() call.
		Dim 文件读取器 As New BinaryReader(File.OpenRead(Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess.MainModule.FileName), "密钥库.bin")))
		Dim 选项个数 As Byte = 文件读取器.ReadByte
		Dim 操作系统版本(选项个数 - 1) As String
		For a As Byte = 0 To 选项个数 - 1
			操作系统版本(a) = 文件读取器.ReadString
		Next
		版本或密钥.ItemsSource = 操作系统版本
		ReDim 密钥列表(选项个数 - 1)
		For a As Byte = 0 To 选项个数 - 1
			密钥列表(a) = 文件读取器.ReadString
		Next
		文件读取器.Close()
	End Sub

	Private Sub 激活Windows_Click(sender As Object, e As RoutedEventArgs) Handles 激活Windows.Click
		Static 启动信息 As New ProcessStartInfo With {.CreateNoWindow = True, .FileName = "powershell.exe", .UseShellExecute = False, .WindowStyle = ProcessWindowStyle.Hidden, .RedirectStandardOutput = True, .RedirectStandardError = True}
		Dim 服务器地址 As String = KMS服务器地址.Text
		Dim 密钥 = If(版本或密钥.SelectedIndex < 0, 版本或密钥.Text, 密钥列表(版本或密钥.SelectedIndex))
		'防注入检查
		If 服务器地址.Contains("'") OrElse 服务器地址.Contains(vbCrLf) Then
			Windows激活结果.Text = "服务器地址含有非法字符"
			Return
		End If
		If 密钥.Contains("'") OrElse 密钥.Contains(vbCrLf) Then
			Windows激活结果.Text = "密钥含有非法字符"
			Return
		End If
		启动信息.Arguments = "Start-Process powershell -ArgumentList @('slmgr /skms " & 服务器地址 & ";slmgr /ipk " & 密钥 & ";slmgr /ato') -Verb Runas"
		Windows激活结果.Text = ""
		Call Process.Start(启动信息)
	End Sub

	Private Sub 激活Office_Click(sender As Object, e As RoutedEventArgs) Handles 激活Office.Click
		Static 启动信息 As New ProcessStartInfo With {.CreateNoWindow = True, .FileName = "powershell.exe", .UseShellExecute = False, .WindowStyle = ProcessWindowStyle.Hidden, .RedirectStandardOutput = True, .RedirectStandardError = True}
		Dim 服务器地址 As String = KMS服务器地址.Text
		'防注入检查
		If 服务器地址.Contains("'") OrElse 服务器地址.Contains(vbCrLf) Then
			Office激活结果.Text = "服务器地址含有非法字符"
			Return
		End If
		启动信息.Arguments = "Start-Process powershell -ArgumentList @('cd ''C:\Program Files\Microsoft Office\Office16'';cscript OSPP.VBS /sethst:" & 服务器地址 & ";cscript OSPP.VBS /act') -Verb Runas"
		Office激活结果.Text = ""
		Call Process.Start(启动信息)
	End Sub

	Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
		写出设置()
	End Sub

	Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
		读入设置()
	End Sub
End Class
