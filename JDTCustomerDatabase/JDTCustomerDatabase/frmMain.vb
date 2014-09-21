Imports System.IO
Public Class frmMain
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadCSV()
        ListBox1.SelectedIndex = 0
    End Sub
    Private Function loadCSV()
        ListBox1.Items.Clear()
        ListBox1.Items.Add("<New Customer>")
        If Not System.IO.Directory.Exists("C:\JDexterTech\") Then
            System.IO.Directory.CreateDirectory("C:\JDexterTech\")
        End If
        If Not System.IO.File.Exists("C:\JDexterTech\Customer.csv") Then
            System.IO.File.Create("C:\JDexterTech\Customer.csv")
            Return 1
            Exit Function
        End If
        Dim readCSV As New System.IO.StreamReader("C:\JDexterTech\Customer.csv")
        Dim lineCSV As String
        Try
            Do While readCSV.Peek <> -1
                lineCSV = readCSV.ReadLine
                Dim splt = Split(lineCSV, ",")
                ListBox1.Items.Add(splt(0) & " " & splt(1) & " " & splt(2))
            Loop
            readCSV.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        ListBox1.SelectedIndex = 0
        Return 0
    End Function
    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        ListBox1.Height = Me.Height - 50
    End Sub
    Private Sub txtPhone_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone.GotFocus
        If txtPhone.Text = "" Then
            txtPhone.Text = "615"
            txtPhone.Select(4, 4)
        End If
    End Sub
    Private Sub txtPhone_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone.LostFocus
        If Not txtPhone.Text.Contains("-") AndAlso Double.Parse(txtPhone.Text) = 615 Then
            txtPhone.Text = ""
            Exit Sub
        End If
        If Not txtPhone.Text.Contains("-") And txtPhone.Text.Length > 6 Then
            txtPhone.Text = txtPhone.Text.Substring(0, 3) & "-" & txtPhone.Text.Substring(3, 3) & "-" & txtPhone.Text.Substring(6, 4)
        End If
    End Sub
    Private Sub txtPhone_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPhone.TextChanged
        If (txtPhone.Text.Length > 10 And txtPhone.Text.Contains("615")) And Not txtPhone.Text.Contains("-") Then
            txtPhone.Text = txtPhone.Text.Replace("615", "")
            txtPhone.Select(8, 7)
        End If
    End Sub
    Private Sub txtPhone2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone2.GotFocus
        If txtPhone2.Text = "" Then
            txtPhone2.Text = "615"
            txtPhone2.Select(4, 4)
        End If
    End Sub
    Private Sub txtPhone2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone2.LostFocus
        If Not txtPhone2.Text.Contains("-") AndAlso Double.Parse(txtPhone2.Text) = 615 Then
            txtPhone2.Text = ""
            Exit Sub
        End If
        If Not txtPhone2.Text.Contains("-") And txtPhone2.Text.Length > 6 Then
            txtPhone2.Text = txtPhone2.Text.Substring(0, 3) & "-" & txtPhone2.Text.Substring(3, 3) & "-" & txtPhone2.Text.Substring(6, 4)
        End If
    End Sub
    Private Sub txtPhone2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPhone2.TextChanged
        If (txtPhone2.Text.Length > 10 And txtPhone2.Text.Contains("615")) And Not txtPhone2.Text.Contains("-") Then
            txtPhone2.Text = txtPhone2.Text.Replace("615", "")
            txtPhone2.Select(8, 7)
        End If
    End Sub
    Private Sub btnFlip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFlip.Click
        Dim tmp As String = txtPhone.Text
        Dim tmpe As String = txtEx.Text
        txtPhone.Text = txtPhone2.Text
        txtEx.Text = txtEx2.Text
        txtPhone2.Text = tmp
        txtEx2.Text = tmpe
    End Sub
    Private Sub txtFirst_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFirst.TextChanged
        If txtFirst.Text.Length = 1 Then
            txtFirst.Text = txtFirst.Text.ToUpper()
            txtFirst.Select(2, 2)
        End If
    End Sub
    Private Sub txtMid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMid.TextChanged
        If txtMid.Text.Length = 1 Then
            txtMid.Text = txtMid.Text.ToUpper()
            txtMid.Select(2, 2)
        End If
    End Sub
    Private Sub txtLast_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLast.TextChanged
        If txtLast.Text.Length = 1 Then
            txtLast.Text = txtLast.Text.ToUpper()
            txtLast.Select(2, 2)
        End If
    End Sub
    Private Sub txtCity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCity.TextChanged
        If txtCity.Text.Length = 1 Then
            txtCity.Text = txtCity.Text.ToUpper()
            txtCity.Select(2, 2)
        End If
    End Sub
    Private Sub txtState_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtState.GotFocus
        If txtCity.AutoCompleteCustomSource.Contains(txtCity.Text) Then
            txtState.Text = "TN"
        End If
    End Sub
    Private Sub txtZip_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtZip.GotFocus
        If txtCity.Text = "Smyrna" Then
            txtZip.Text = "37167"
        ElseIf txtCity.Text = "La Vergne" Then
            txtZip.Text = "37086"
        End If
    End Sub
    Private Sub chkLaptop_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLaptop.CheckedChanged
        If chkLaptop.Checked = True Then
            chkPSU.Checked = True
            chkPSU.Enabled = True
            chkBag.Enabled = True
        Else
            chkPSU.Enabled = False
            chkBag.Enabled = False
            chkPSU.CheckState = CheckState.Indeterminate
        End If
    End Sub
    Private Sub txtOem_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOem.TextChanged
        If txtOem.Text.Length = 1 Then
            txtOem.Text = txtOem.Text.ToUpper()
            txtOem.Select(2, 2)
        End If
    End Sub
    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedIndex = 0 Then
            clearText()
        Else
            If Not System.IO.Directory.Exists("C:\JDexterTech\") Then
                System.IO.Directory.CreateDirectory("C:\JDexterTech\")
            End If
            If Not System.IO.File.Exists("C:\JDexterTech\Customer.csv") Then
                System.IO.File.Create("C:\JDexterTech\Customer.csv")
            End If
            Dim readCSV As New System.IO.StreamReader("C:\JDexterTech\Customer.csv")
            Dim lineCSV As String
            Try
                Do While readCSV.Peek <> -1
                    lineCSV = readCSV.ReadLine
                    Dim splt = Split(lineCSV, ",")
                    If ListBox1.SelectedItem.ToString = splt(0) & " " & splt(1) & " " & splt(2) Then
                        txtFirst.Text = splt(0)
                        txtMid.Text = splt(1)
                        txtLast.Text = splt(2)
                        txtPhone.Text = splt(3)
                        txtPhone2.Text = splt(4)
                        txtEx.Text = splt(5)
                        txtEx2.Text = splt(6)
                        txtAddress.Text = splt(7)
                        txtCity.Text = splt(8)
                        txtState.Text = splt(9)
                        txtZip.Text = splt(10)
                        txtOem.Text = splt(11)
                        txtDetail.Text = splt(12)
                        chkLaptop.Checked = splt(13)
                        chkPSU.Checked = splt(14)
                        chkBag.Checked = splt(15)
                        txtExtra.Text = splt(16)
                        txtPass.Text = splt(17)
                        txtProblem.Text = splt(18)
                        txtDam.Text = splt(19)
                        txtRef.Text = splt(20)
                        txtEmail.text = splt(21)
                    End If
                Loop
                readCSV.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
        txtFirst.Focus()
    End Sub
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        If MessageBox.Show("Clear shown customer data?", "Clear", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
            clearText()
        End If
    End Sub
    Private Function clearText()
        txtFirst.Text = ""
        txtMid.Text = ""
        txtLast.Text = ""
        txtPhone.Text = ""
        txtPhone2.Text = ""
        txtEx.Text = ""
        txtEx2.Text = ""
        txtAddress.Text = ""
        txtCity.Text = ""
        txtState.Text = ""
        txtZip.Text = ""
        txtOem.Text = ""
        txtDetail.Text = ""
        txtExtra.Text = ""
        txtPass.Text = ""
        txtProblem.Text = ""
        txtDam.Text = ""
        txtRef.Text = ""
        txtEmail.Text = ""
        chkLaptop.Checked = False
        Return 0
    End Function
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        txtFirst.Text = txtFirst.Text.Replace(",", ";")
        txtMid.Text = txtMid.Text.Replace(",", ";")
        txtLast.Text = txtLast.Text.Replace(",", ";")
        txtPhone.Text = txtPhone.Text.Replace(",", ";")
        txtPhone2.Text = txtPhone2.Text.Replace(",", ";")
        txtEx.Text = txtEx.Text.Replace(",", ";")
        txtEx2.Text = txtEx2.Text.Replace(",", ";")
        txtAddress.Text = txtAddress.Text.Replace(",", ";")
        txtCity.Text = txtCity.Text.Replace(",", ";")
        txtState.Text = txtState.Text.Replace(",", ";")
        txtZip.Text = txtZip.Text.Replace(",", ";")
        txtOem.Text = txtOem.Text.Replace(",", ";")
        txtDetail.Text = txtDetail.Text.Replace(",", ";")
        txtExtra.Text = txtExtra.Text.Replace(",", ";")
        txtPass.Text = txtPass.Text.Replace(",", ";")
        txtProblem.Text = txtProblem.Text.Replace(",", ";")
        txtDam.Text = txtDam.Text.Replace(",", ";")
        txtRef.Text = txtRef.Text.Replace(",", ";")
        txtProblem.Text = txtProblem.Text.Replace(vbCrLf, "  ")
        txtExtra.Text = txtExtra.Text.Replace(vbCrLf, "  ")
        txtEmail.Text = txtEmail.Text.Replace(vbCrLf, "  ")
        Try
            If ListBox1.SelectedIndex = 0 Then
                Dim writeCSV As New System.IO.StreamWriter("C:\JDexterTech\Customer.csv", True)
                writeCSV.WriteLine(txtFirst.Text & "," & _
                txtMid.Text & "," & _
                txtLast.Text & "," & _
                txtPhone.Text & "," & _
                txtPhone2.Text & "," & _
                txtEx.Text & "," & _
                txtEx2.Text & "," & _
                txtAddress.Text & "," & _
                txtCity.Text & "," & _
                txtState.Text & "," & _
                txtZip.Text & "," & _
                txtOem.Text & "," & _
                txtDetail.Text & "," & _
                chkLaptop.Checked.ToString & "," & _
                chkPSU.Checked.ToString & "," & _
                chkBag.Checked.ToString & "," & _
                txtExtra.Text & "," & _
                txtPass.Text & "," & _
                txtProblem.Text & "," & _
                txtDam.Text & "," & _
                txtRef.Text & "," & _
                txtEmail.Text)
                writeCSV.Close()

            Else
                Dim lineNum As Integer = 0
                Dim readCSV As New System.IO.StreamReader("C:\JDexterTech\Customer.csv")
                Dim lineCSV As String
                Do While readCSV.Peek <> -1
                    lineCSV = readCSV.ReadLine
                    Dim splt = Split(lineCSV, ",")
                    If ListBox1.SelectedItem.ToString = splt(0) & " " & splt(1) & " " & splt(2) Then
                        Exit Do
                    End If
                    lineNum = lineNum + 1
                Loop
                readCSV.Close()
                deleteLine(lineNum)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        loadCSV()
    End Sub
    Private Function deleteLine(ByVal l As Integer, Optional ByVal remove As Boolean = False)
        Dim readCSV As New System.IO.StreamReader("C:\JDexterTech\Customer.csv")
        Dim wholeCSV As String = readCSV.ReadToEnd
        readCSV.Close()
        Dim piec = Split(wholeCSV, vbCrLf)
        Dim writeCSV As New System.IO.StreamWriter("C:\JDexterTech\Customer.csv", False)
        For ix As Integer = 0 To piec.Length - 1
            If ix = l Then
                If Not remove Then
                    writeCSV.WriteLine(txtFirst.Text & "," & _
                    txtMid.Text & "," & _
                    txtLast.Text & "," & _
                    txtPhone.Text & "," & _
                    txtPhone2.Text & "," & _
                    txtEx.Text & "," & _
                    txtEx2.Text & "," & _
                    txtAddress.Text & "," & _
                    txtCity.Text & "," & _
                    txtState.Text & "," & _
                    txtZip.Text & "," & _
                    txtOem.Text & "," & _
                    txtDetail.Text & "," & _
                    chkLaptop.Checked.ToString & "," & _
                    chkPSU.Checked.ToString & "," & _
                    chkBag.Checked.ToString & "," & _
                    txtExtra.Text & "," & _
                    txtPass.Text & "," & _
                    txtProblem.Text & "," & _
                    txtDam.Text & "," & _
                    txtRef.Text & "," & _
                    txtEmail.Text)
                End If
            Else
                If Not ix = piec.Length - 1 Then
                    writeCSV.WriteLine(piec(ix))
                Else
                    writeCSV.Write(piec(ix))
                End If

            End If
        Next
        writeCSV.Close()
        Return 0
    End Function
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim MyPrintObject As New TextPrint(txtFirst.Text & " " & txtMid.Text & " " & txtLast.Text & " " & vbCrLf & _
            "Primary: " & txtPhone.Text & " " & txtEx.Text & vbCrLf & _
            "Secondary: " & txtPhone2.Text & " " & txtEx2.Text & vbCrLf & _
            "Email: " & txtEmail.Text & vbCrLf & _
            txtAddress.Text & vbCrLf & _
            txtCity.Text & "," & txtState.Text & " " & txtZip.Text & vbCrLf & _
            txtOem.Text & " " & txtDetail.Text & vbCrLf & vbCrLf & _
            "Laptop: " & chkLaptop.Checked.ToString & vbCrLf & _
            "PSU: " & chkPSU.Checked.ToString & vbCrLf & _
            "Bag: " & chkBag.Checked.ToString & vbCrLf & _
            txtExtra.Text & vbCrLf & vbCrLf & _
            "Password: " & txtPass.Text & vbCrLf & vbCrLf & _
            "Problem: " & txtProblem.Text & vbCrLf & _
            "Damage: " & txtDam.Text & vbCrLf & _
            "Referred From: " & txtRef.Text)
        MyPrintObject.Font = New Font("Tahoma", 14)
        MyPrintObject.Print()
    End Sub
    Private Sub DumpToNotepadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DumpToNotepadToolStripMenuItem.Click
        If Not System.IO.Directory.Exists("C:\JDexterTech\TXTs\") Then
            System.IO.Directory.CreateDirectory("C:\JDexterTech\TXTs\")
        End If
        Try
            Dim ts As String = "C:\JDexterTech\TXTs\" & txtFirst.Text & " " & txtMid.Text & " " & txtLast.Text & " - " & DateTime.Now.ToString.Replace(":", ";").Replace("/", "-") & ".txt"
            Dim writeTXT As New System.IO.StreamWriter(ts)
            writeTXT.Write(txtFirst.Text & " " & txtMid.Text & " " & txtLast.Text & " " & vbCrLf & _
            "Primary: " & txtPhone.Text & " " & txtEx.Text & vbCrLf & _
            "Secondary: " & txtPhone2.Text & " " & txtEx2.Text & vbCrLf & _
            "Email: " & txtEmail.Text & vbCrLf & _
            txtAddress.Text & vbCrLf & _
            txtCity.Text & "," & txtState.Text & " " & txtZip.Text & vbCrLf & _
            txtOem.Text & " " & txtDetail.Text & vbCrLf & vbCrLf & _
            "Laptop: " & chkLaptop.Checked.ToString & vbCrLf & _
            "PSU: " & chkPSU.Checked.ToString & vbCrLf & _
            "Bag: " & chkBag.Checked.ToString & vbCrLf & _
            txtExtra.Text & vbCrLf & vbCrLf & _
            "Password: " & txtPass.Text & vbCrLf & vbCrLf & _
            "Problem: " & txtProblem.Text & vbCrLf & _
            "Damage: " & txtDam.Text & vbCrLf & _
            "Referred From: " & txtRef.Text)
            writeTXT.Close()
            Shell("notepad " & ts, AppWinStyle.NormalFocus, True)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub DumpToQuickbooksToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DumpToQuickbooksToolStripMenuItem.Click
        Dim hWnd As Long
        Try
            Dim p As Process() = Process.GetProcessesByName("QBW32")
            hWnd = CType(p(0).MainWindowHandle, Integer)
            ShowWindow(hWnd, 9)
            SetForegroundWindow(hWnd)
            SendKeys.SendWait("^i")
            SendKeys.SendWait("%{DOWN}")
            SendKeys.SendWait("{DOWN}")
            SendKeys.SendWait("{ENTER}")
            SendKeys.SendWait(txtFirst.Text + " " + txtMid.Text + " " + txtLast.Text)
            SendKeys.SendWait("{TAB 5}")
            SendKeys.SendWait(txtFirst.Text)
            SendKeys.SendWait("{TAB}")
            SendKeys.SendWait(txtMid.Text)
            SendKeys.SendWait("{TAB}")
            SendKeys.SendWait(txtLast.Text)
            SendKeys.SendWait("{TAB}")
            SendKeys.SendWait("{ENTER}")
            SendKeys.SendWait(txtAddress.Text & "{ENTER}" & txtCity.Text & "," & txtState.Text & " " & txtZip.Text)
            SendKeys.SendWait("{TAB 3}")
            SendKeys.SendWait(txtPhone.Text)
            SendKeys.SendWait("{TAB 2}")
            SendKeys.SendWait(txtPhone2.Text)
            SendKeys.SendWait("{TAB 4}")
            SendKeys.SendWait("{ENTER 2}")
            SendKeys.SendWait("{TAB 7}")
            SendKeys.SendWait("{ENTER}")
            SendKeys.SendWait("{TAB 14}")
            SendKeys.SendWait("1")
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        If MessageBox.Show("Delete customer?", "Delete", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
            deleteLine(ListBox1.SelectedIndex - 1, True)
            loadCSV()
        End If
    End Sub
End Class
Public Class TextPrint
    ' Inherits all the functionality of a PrintDocument
    Inherits Printing.PrintDocument
    ' Private variables to hold default font and text
    Private fntPrintFont As Font
    Private strText As String
    Public Sub New(ByVal Text As String)
        ' Sets the file stream
        MyBase.New()
        strText = Text
    End Sub
    Public Property Text() As String
        Get
            Return strText
        End Get
        Set(ByVal Value As String)
            strText = Value
        End Set
    End Property
    Protected Overrides Sub OnBeginPrint(ByVal ev _
                                As Printing.PrintEventArgs)
        ' Run base code
        MyBase.OnBeginPrint(ev)
        ' Sets the default font
        If fntPrintFont Is Nothing Then
            fntPrintFont = New Font("Times New Roman", 12)
        End If
    End Sub
    Public Property Font() As Font
        ' Allows the user to override the default font
        Get
            Return fntPrintFont
        End Get
        Set(ByVal Value As Font)
            fntPrintFont = Value
        End Set
    End Property
    Protected Overrides Sub OnPrintPage(ByVal ev _
       As Printing.PrintPageEventArgs)
        ' Provides the print logic for our document

        ' Run base code
        MyBase.OnPrintPage(ev)
        ' Variables
        Static intCurrentChar As Integer
        Dim intPrintAreaHeight, intPrintAreaWidth, _
            intMarginLeft, intMarginTop As Integer
        ' Set printing area boundaries and margin coordinates
        With MyBase.DefaultPageSettings
            intPrintAreaHeight = .PaperSize.Height - _
                               .Margins.Top - .Margins.Bottom
            intPrintAreaWidth = .PaperSize.Width - _
                              .Margins.Left - .Margins.Right
            intMarginLeft = .Margins.Left 'X
            intMarginTop = .Margins.Top   'Y
        End With
        ' If Landscape set, swap printing height/width
        If MyBase.DefaultPageSettings.Landscape Then
            Dim intTemp As Integer
            intTemp = intPrintAreaHeight
            intPrintAreaHeight = intPrintAreaWidth
            intPrintAreaWidth = intTemp
        End If
        ' Calculate total number of lines
        Dim intLineCount As Int32 = _
                CInt(intPrintAreaHeight / Font.Height)
        ' Initialise rectangle printing area
        Dim rectPrintingArea As New RectangleF(intMarginLeft, _
            intMarginTop, intPrintAreaWidth, intPrintAreaHeight)
        ' Initialise StringFormat class, for text layout
        Dim objSF As New StringFormat(StringFormatFlags.LineLimit)
        ' Figure out how many lines will fit into rectangle
        Dim intLinesFilled, intCharsFitted As Int32
        ev.Graphics.MeasureString(Mid(strText, _
                    UpgradeZeros(intCurrentChar)), Font, _
                    New SizeF(intPrintAreaWidth, _
                    intPrintAreaHeight), objSF, _
                    intCharsFitted, intLinesFilled)
        ' Print the text to the page
        ev.Graphics.DrawString(Mid(strText, _
            UpgradeZeros(intCurrentChar)), Font, _
            Brushes.Black, rectPrintingArea, objSF)
        ' Increase current char count
        intCurrentChar += intCharsFitted
        ' Check whether we need to print more
        If intCurrentChar < strText.Length Then
            ev.HasMorePages = True
        Else
            ev.HasMorePages = False
            intCurrentChar = 0
        End If
    End Sub
    Public Function UpgradeZeros(ByVal Input As Integer) As Integer
        ' Upgrades all zeros to ones
        ' - used as opposed to defunct IIF or messy If statements
        If Input = 0 Then
            Return 1
        Else
            Return Input
        End If
    End Function
End Class

