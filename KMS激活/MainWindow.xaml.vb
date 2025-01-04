Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
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
			If 读入器.ReadBoolean Then
				版本或密钥.Text = 读入器.ReadString
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
			写出器.Write(版本或密钥.SelectedIndex < 0)
			写出器.Write(版本或密钥.Text) '用户自定义密钥
		End Using
	End Sub

	<DllImport("WinBrand.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
	Public Shared Function BrandingFormatString(ByVal format As String) As String
	End Function

	Shared ReadOnly 当前版本 As String = BrandingFormatString("%WINDOWS_LONG%")

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
		Process.Start("cscript", "//B //Nologo C:\Windows\System32\slmgr.vbs /skms " & 服务器地址)
		If 当前版本.EndsWith("Evaluation") Then
			Process.Start("Dism", $"/online /Set-Edition:ServerStandard /ProductKey:{密钥} /AcceptEula")
		Else
			Process.Start("cscript", "//B //Nologo C:\Windows\System32\slmgr.vbs /ipk " & 密钥)
		End If
		Process.Start("cscript", "//B //Nologo C:\Windows\System32\slmgr.vbs /ato")
		Windows激活结果.Text = ""
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

	Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
		Dim 版本密钥 As IEnumerable(Of String()) = From 行 In New StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream("KMS激活.密钥库.txt")).ReadToEnd().Split(vbCrLf) Select 行.Split(vbTab)
		Dim 所有版本 As IEnumerable(Of String) = From 两项 In 版本密钥 Select 两项(0)
		版本或密钥.ItemsSource = 所有版本
		密钥列表 = (From 两项 In 版本密钥 Select 两项(1)).ToArray
		Dim 匹配度 = (From 两项 In 版本密钥 Select Aggregate 元组 In 两项(0).Zip(当前版本) Where 元组.First <> " " Take While 元组.First = 元组.Second Into Count()).ToArray
		版本或密钥.SelectedIndex = (From Index In 匹配度.Index Where Index.Item = 匹配度.Max Select Index.Index).First
		读入设置()
	End Sub

	Private Sub 设置OSPP位置_Click(sender As Object, e As RoutedEventArgs) Handles 设置OSPP位置.Click
		If 目录选择对话框.ShowDialog Then
			OSPP位置.Text = Path.GetDirectoryName(目录选择对话框.FileName)
		End If
	End Sub
End Class