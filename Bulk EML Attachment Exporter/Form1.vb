
Imports LumiSoft.Net.Mime

Public Class Form1
    Dim files As New List(Of String)

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dialog = OpenFileDialog1.ShowDialog
        If dialog = DialogResult.OK Then
            files.AddRange(OpenFileDialog1.FileNames)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim message As Mime
        For Each file In files
            message = Mime.Parse(file)
            Dim info As New IO.FileInfo(file)
            Dim savedir As IO.DirectoryInfo
            If Not IO.Directory.Exists(info.DirectoryName + "\" + file) Then
                savedir = IO.Directory.CreateDirectory(info.DirectoryName + "\" + info.Name.Substring(0, info.Name.LastIndexOf(".eml")))
            Else
                savedir = New IO.DirectoryInfo(info.DirectoryName + "\" + file)
            End If
            For Each att As MimeEntity In message.Attachments
                IO.File.WriteAllBytes(savedir.FullName & "\" & If(att.ContentDisposition_FileName, att.ContentType_Name), att.Data)
            Next
        Next
    End Sub
End Class
