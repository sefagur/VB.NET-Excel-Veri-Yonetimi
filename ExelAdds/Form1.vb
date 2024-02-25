Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.Drawing

Public Class Form1
    Dim excelFilePath As String = ""

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Dim selectedRowIndex As Integer = ListBox1.SelectedIndex
        If selectedRowIndex >= 0 Then
            ' Seçilen satırı silmek için onay.
            Dim result As DialogResult = MessageBox.Show("Seçilen Satırı Silmek İstiyor Musunuz?", "Onay", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                ' ListBox'tan seçili satırı sil.
                ListBox1.Items.RemoveAt(selectedRowIndex)

                ' Excel dosyasındaki seçilen satırı sil.
                Dim excelApp As Excel.Application = Nothing
                Dim excelWorkbook As Excel.Workbook = Nothing
                Dim excelWorksheet As Excel.Worksheet = Nothing

                Try
                    excelApp = New Excel.Application()
                    excelWorkbook = excelApp.Workbooks.Open(excelFilePath)
                    excelWorksheet = excelWorkbook.Sheets(1)

                    Dim lastRow As Integer = excelWorksheet.Cells(excelWorksheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
                    Dim rowToDelete As Integer = selectedRowIndex + 1
                    If rowToDelete >= 1 AndAlso rowToDelete <= lastRow Then
                        ' Excel satırını sil.
                        excelWorksheet.Rows(rowToDelete).Delete()
                        excelWorkbook.Save()

                        ' Başarılı mesajı.
                        MessageBox.Show("Satır Başarıyla Silindi")

                        ' Excel dosyasını kapatın.
                        excelWorkbook.Close()
                        excelApp.Quit()
                    End If

                Catch ex As Exception
                    ' Hata mesajı.
                    MessageBox.Show("Hata Oluştu")
                Finally
                    ' Excel nesnelerini serbest bırakın.
                    ReleaseObject(excelWorksheet)
                    ReleaseObject(excelWorkbook)
                    ReleaseObject(excelApp)
                End Try
            End If
        End If
    End Sub

    ' Excel nesnelerini serbest bırak.
    Private Sub ReleaseObject(obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        ' Excel dosyasını seçmek.
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "Excel Dosyaları|*.xlsx;*.xlsm;*.xlsb;*.xls;*.xlt;*.xltx;*.xltm;*.csv|Excel Dosyası (.xlsx)|*.xlsx|Excel Makro Etkin Dosyası (.xlsm)|*.xlsm|Excel Bileşik Dosyası (.xlsb)|*.xlsb|Excel 97-2003 Çalışma Sayfası (.xls)|*.xls|Excel Şablonu (.xlt)|*.xlt|Excel Şablonu (.xltx)|*.xltx|Excel Şablonu Makro Etkin (.xltm)|*.xltm|CSV Dosyası (.csv)|*.csv"
        openFileDialog1.Title = "Excel Dosyası Seçin"
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            excelFilePath = openFileDialog1.FileName
            Label5.Text = "Açılan Dosya : " & Path.GetFileName(openFileDialog1.FileName)

            ' Excel dosyasını açın ve verileri ListBox'a listele.
            ListBox1.Items.Clear()
            Dim excelApp As Excel.Application = Nothing
            Dim excelWorkbook As Excel.Workbook = Nothing
            Dim excelWorksheet As Excel.Worksheet = Nothing

            Try
                Panel5.Visible = True

                excelApp = New Excel.Application()
                excelWorkbook = excelApp.Workbooks.Open(excelFilePath)
                excelWorksheet = excelWorkbook.Sheets(1)


                If CheckBox1.Checked = True Then
                    ' Sütun genişliklerini ayarla.
                    excelWorksheet.Columns("A:C").AutoFit()

                    ' Sayfayı genişlet.
                    excelWorksheet.UsedRange.Columns.AutoFit()
                    excelWorksheet.UsedRange.Rows.AutoFit()

                    Me.WindowState = FormWindowState.Minimized
                Else

                End If

                Dim lastRow As Integer = excelWorksheet.Cells(excelWorksheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
                For i As Integer = 1 To lastRow
                    Dim tcKimlik As Object = excelWorksheet.Cells(i, 1).Value
                    Dim adSoyad As Object = excelWorksheet.Cells(i, 2).Value
                    Dim notlar As Object = excelWorksheet.Cells(i, 3).Value

                    ' Satırları ListBox'a ekle.
                    ListBox1.Items.Add(tcKimlik & " - " & adSoyad & " - " & notlar)
                Next


                ' Excel dosyasını kapat.
                excelWorkbook.Close()
                excelApp.Quit()

                Panel3.Visible = False

                ' Formu Büyült
                Me.WindowState = FormWindowState.Maximized

                ' ListBox'u en altına kaydır
                ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                ListBox1.TopIndex = ListBox1.Items.Count - 1

            Catch ex As Exception
                ' Hata mesajı göster.
                MessageBox.Show("Hata oluştu")
            Finally
                ' Excel nesnelerini serbest bırak.
                ReleaseObject(excelWorksheet)
                ReleaseObject(excelWorkbook)
                ReleaseObject(excelApp)
            End Try
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If String.IsNullOrEmpty(excelFilePath) Then
            MessageBox.Show("Lütfen Önce Excel Dosyasını Seçin")
            Return
        End If

        ' TextBoxlardan alınan bilgileri exele ekle.
        Dim tcKimlik As String = TextBox1.Text ' TC Kimlik
        Dim adSoyad As String = TextBox2.Text ' Ad Soyad
        Dim notlar As String = ComboBox1.Text ' Notlar

        If String.IsNullOrEmpty(tcKimlik) OrElse String.IsNullOrEmpty(adSoyad) Then
            MessageBox.Show("Lütfen Tüm Alanları Doldurun")
            Return
        End If

        ' Excel uygulamasını başlatın ve yeni bir çalışma kitabı aç.
        Dim excelApp As Excel.Application = Nothing
        Dim excelWorkbook As Excel.Workbook = Nothing
        Dim excelWorksheet As Excel.Worksheet = Nothing

        Try
            Dim birlesikMetin As String = tcKimlik & " - " & adSoyad & " - " & notlar
            ListBox1.Items.Add(birlesikMetin)

            excelApp = New Excel.Application()
            excelWorkbook = excelApp.Workbooks.Open(excelFilePath)
            excelWorksheet = excelWorkbook.Sheets(1)


            ' Boş satırın konumunu bul.
            Dim emptyRow As Integer = excelWorksheet.Cells(excelWorksheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1

            ' TextBoxlardan alınan bilgileri exele ekle.
            excelWorksheet.Cells(emptyRow, 1).Value = tcKimlik ' TC Kimlik
            excelWorksheet.Cells(emptyRow, 2).Value = adSoyad ' Ad Soyad
            excelWorksheet.Cells(emptyRow, 3).Value = notlar ' Notlar

            ' Değişiklikleri kaydet ve Excel dosyasını kapat.
            excelWorkbook.Save()
            excelWorkbook.Close()
            excelApp.Quit()

            ' ListBox'u en altına kaydır
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
            ListBox1.TopIndex = ListBox1.Items.Count - 1

            ' Başarılı mesajı göster.
            MessageBox.Show("Veriler Excel Dosyasına Başarıyla Eklendi")
            TextBox1.Clear()
            TextBox2.Clear()
            ComboBox1.Text = ""
            TextBox1.Focus()

        Catch ex As Exception
            ' Hata mesajı göster.
            MessageBox.Show("Hata Oluştu")
        Finally
            ' Excel nesnelerini serbest bırak.
            ReleaseObject(excelWorksheet)
            ReleaseObject(excelWorkbook)
            ReleaseObject(excelApp)

        End Try
    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        PictureBox1.BackgroundImage = My.Resources.Kutucuk
    End Sub

    Private Sub PictureBox1_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox1.MouseLeave
        PictureBox1.BackgroundImage = Nothing
    End Sub

    Private Sub PictureBox2_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox2.MouseMove
        PictureBox2.BackgroundImage = My.Resources.Kutucuk
    End Sub

    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox2.MouseLeave
        PictureBox2.BackgroundImage = Nothing
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim searchText As String = TextBox3.Text

        ' ListBox'taki öğeleri kontrol et ve aranan metinle eşleşen öğeyi seçin.
        For i As Integer = 0 To ListBox1.Items.Count - 1
            Dim itemText As String = ListBox1.Items(i).ToString()
            If itemText.Contains(searchText) Then
                ListBox1.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        System.Diagnostics.Process.Start("https://msgsoftware.blogspot.com/")
    End Sub

    Private Sub PictureBox3_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox3.MouseMove
        PictureBox3.Image = My.Resources.msg1
    End Sub

    Private Sub PictureBox3_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox3.MouseLeave
        PictureBox3.Image = My.Resources.msg
    End Sub


    Private Sub ListBox2_DoubleClick(sender As Object, e As EventArgs) Handles ListBox2.DoubleClick
        Try
            Dim result As DialogResult = MessageBox.Show("Seçilen Notu Silmek İstiyor Musunuz?", "Onay", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                If ListBox2.SelectedIndex <> -1 Then ' Bir öğe seçildi mi kontrol et
                    Dim selectedItem As String = ListBox2.SelectedItem.ToString() ' Seçilen öğeyi al
                    ListBox2.Items.Remove(selectedItem) ' ListBox'tan öğeyi kaldır
                    ComboBox1.Items.Remove(selectedItem) ' ComboBox'tan öğeyi kaldır
                    My.Settings.ListBoxItems.Remove(selectedItem) ' My.Settings'ten öğeyi kaldır
                    My.Settings.Save() ' Değişiklikleri kaydet

                    MessageBox.Show("Not Başarıyla Silindi")
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Hata oluştu")
        End Try

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' My.Settings'ten kayıtlı verileri çek
        If My.Settings.ListBoxItems Is Nothing Then
            My.Settings.ListBoxItems = New System.Collections.Specialized.StringCollection()
            My.Settings.Save() ' Boş bir StringCollection oluşturup kaydet
        End If

        For Each item In My.Settings.ListBoxItems
            ListBox2.Items.Add(item)
            ComboBox1.Items.Add(item)
        Next
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Try
            Dim newItem As String = TextBox4.Text.Trim() ' TextBox'tan gelen yeni öğeyi al ve baştaki ve sondaki boşlukları kaldır


            If String.IsNullOrEmpty(newItem) Then
                MessageBox.Show("Lütfen Tüm Alanları Doldurun.")
                Return
            End If

            If Not String.IsNullOrEmpty(newItem) Then ' Yeni öğe boş değilse devam et
                If Not ListBox2.Items.Contains(newItem) Then ' Yeni öğe zaten ListBox'ta yoksa devam et
                    ListBox2.Items.Add(newItem) ' ListBox'a yeni öğe ekle
                    ComboBox1.Items.Add(newItem) ' ComboBox'a yeni öğe ekle
                    My.Settings.ListBoxItems.Add(newItem) ' My.Settings'e yeni öğe ekle
                    My.Settings.Save() ' Değişiklikleri kaydet
                End If
            End If

            TextBox4.Clear() ' TextBox'ı temizle

            ' ListBox'u en altına kaydırın
            ListBox2.SelectedIndex = ListBox2.Items.Count - 1
            ListBox2.TopIndex = ListBox2.Items.Count - 1

            MessageBox.Show("Not Başarıyla Eklendi")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        Panel4.Visible = True

        ' ListBox'u en altına kaydırın
        ListBox2.SelectedIndex = ListBox2.Items.Count - 1
        ListBox2.TopIndex = ListBox2.Items.Count - 1
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Panel4.Visible = False
    End Sub

    Private Sub PictureBox4_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox4.MouseMove
        PictureBox4.BackgroundImage = My.Resources.Kutucuk
    End Sub

    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox4.MouseLeave
        PictureBox4.BackgroundImage = Nothing
    End Sub

    Private Sub PictureBox5_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox5.MouseMove
        PictureBox5.BackgroundImage = My.Resources.Kutucuk
    End Sub

    Private Sub PictureBox5_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox5.MouseLeave
        PictureBox5.BackgroundImage = Nothing
    End Sub

    Private Sub PictureBox6_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox6.MouseMove
        PictureBox6.BackgroundImage = My.Resources.Kutucuk
    End Sub

    Private Sub PictureBox6_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox6.MouseLeave
        PictureBox6.BackgroundImage = Nothing
    End Sub
End Class
