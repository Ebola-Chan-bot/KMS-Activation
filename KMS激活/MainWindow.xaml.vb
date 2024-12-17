Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Class MainWindow
	Private 密钥列表 As String()
	ReadOnly 设置目录 As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
	ReadOnly 设置路径 As String = Path.Combine(设置目录, "KMS激活设置.bin")
	ReadOnly 启动信息 As New ProcessStartInfo With {.CreateNoWindow = True, .FileName = "powershell.exe", .UseShellExecute = False, .WindowStyle = ProcessWindowStyle.Hidden, .RedirectStandardOutput = True, .RedirectStandardError = True}
	ReadOnly 目录选择对话框 As New Microsoft.Win32.OpenFileDialog With {.Filter = "OSPP.VBS|OSPP.VBS"}

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
			OSPP位置.Text = 读入器.ReadString
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
			写出器.Write(OSPP位置.Text)
			写出器.Write(CSByte(版本或密钥.SelectedIndex))
			写出器.Write(版本或密钥.Text)
		End Using
	End Sub

	Private Sub 激活Windows_Click(sender As Object, e As RoutedEventArgs) Handles 激活Windows.Click
		Dim 服务器地址 As String = KMS服务器地址.Text
		Dim 密钥 = If(版本或密钥.SelectedIndex < 0, 版本或密钥.Text, 密钥列表(版本或密钥.SelectedIndex))
		'防注入检查
		If 服务器地址.Contains("'"c) OrElse 服务器地址.Contains(vbCrLf) Then
			Windows激活结果.Text = "服务器地址含有非法字符"
			Return
		End If
		If 密钥.Contains("'"c) OrElse 密钥.Contains(vbCrLf) Then
			Windows激活结果.Text = "密钥含有非法字符"
			Return
		End If
		启动信息.Arguments = "Start-Process powershell -ArgumentList @('slmgr /skms " & 服务器地址 & ";slmgr /ipk " & 密钥 & ";slmgr /ato') -Verb Runas"
		Windows激活结果.Text = ""
		Call Process.Start(启动信息)
	End Sub

	Private Sub 激活Office_Click(sender As Object, e As RoutedEventArgs) Handles 激活Office.Click
		Dim 服务器地址 As String = KMS服务器地址.Text
		'防注入检查
		If 服务器地址.Contains("'"c) OrElse 服务器地址.Contains(vbCrLf) Then
			Office激活结果.Text = "服务器地址含有非法字符"
			Return
		End If
		If File.Exists(Path.Combine(OSPP位置.Text, "OSPP.VBS")) Then
			启动信息.Arguments = "Start-Process powershell -ArgumentList @('cd ''" & OSPP位置.Text & "'';cscript OSPP.VBS /sethst:" & 服务器地址 & ";cscript OSPP.VBS /act;pause') -Verb Runas"
			Office激活结果.Text = ""
		Else
			Office激活结果.Text = "没有找到OSPP，请检查你的Office版本和安装位置"
		End If
		Call Process.Start(启动信息)
	End Sub

	Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
		写出设置()
	End Sub

	<DllImport("WinBrand.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
	Public Shared Function BrandingFormatString(ByVal format As String) As String
	End Function

	Private Async Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
		Dim 文本列表 = Await File.ReadAllLinesAsync(Path.Combine(Path.GetDirectoryName(Environment.ProcessPath), "密钥库.txt"))
		Dim 选项上限 As Byte = 文本列表.Length - 1
		Dim 操作系统版本(选项上限) As String
		ReDim 密钥列表(选项上限)
		Dim 两项 As String()
		For a = 0 To 选项上限
			两项 = 文本列表(a).Split(vbTab)
			操作系统版本(a) = 两项(0)
			密钥列表(a) = 两项(1)
		Next
		版本或密钥.ItemsSource = 操作系统版本
		Dim 当前版本 As String = BrandingFormatString("%WINDOWS_LONG%")
		For a = 0 To 选项上限
			If 操作系统版本(a) = 当前版本 Then
				版本或密钥.SelectedIndex = a
				Exit For
			End If
		Next
		读入设置()
	End Sub

	Private Sub 设置OSPP位置_Click(sender As Object, e As RoutedEventArgs) Handles 设置OSPP位置.Click
		If 目录选择对话框.ShowDialog Then
			OSPP位置.Text = Path.GetDirectoryName(目录选择对话框.FileName)
		End If
	End Sub
End Class
