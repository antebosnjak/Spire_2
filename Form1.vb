
Imports System.IO

Imports Spire.Pdf

Imports Spire.Pdf.Graphics

Imports System.Drawing

Imports Spire.Pdf.Conversion






Public Class Form1

    Private files() As System.IO.FileInfo


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim broj_datoteka As Integer
        Dim broj_jpg As Integer
        Dim datoteka As String

        Dim datoteka2 As String

        Dim datoteka3 As String
        Dim postotak As Integer


        datoteka = ""
        datoteka2 = ""
        datoteka3 = ""


        Dim dirinfo As New System.IO.DirectoryInfo("C:\Users\GeoModel\Desktop\Povratnice\")
        files = dirinfo.GetFiles("*.jpg", IO.SearchOption.TopDirectoryOnly)






        broj_jpg = 0
        For Each file In files
            broj_jpg = broj_jpg + 1

        Next


        Dim files2(broj_jpg - 1) As String


        ProgressBar1.Minimum = 0
        'ProgressBar1.Step = 1
        ProgressBar1.Maximum = broj_jpg + 2





        broj_datoteka = 0
        For Each file In files
            broj_datoteka = broj_datoteka + 1





            Dim doc As New PdfDocument()
            Dim section As PdfSection = doc.Sections.Add()

            Dim page As PdfPageBase = doc.Pages.Add(PdfPageSize.A4, New PdfMargins(0))



            'Load a tif image from system
            Dim image As PdfImage = PdfImage.FromFile("C:\Users\GeoModel\Desktop\Povratnice\" + file.Name)
            'Set image display location and size in PDF



            Dim widthFitRate As Single = image.PhysicalDimension.Width / page.Canvas.ClientSize.Width

            Dim heightFitRate As Single = image.PhysicalDimension.Height / page.Canvas.ClientSize.Height



            Dim fitRate As Single = Math.Max(widthFitRate, heightFitRate)


            Dim fitWidth As Single = image.PhysicalDimension.Width / fitRate

            Dim fitHeight As Single = image.PhysicalDimension.Height / fitRate



            page.Canvas.DrawImage(image, 0, 0, fitWidth, fitHeight)


            doc.SaveToFile("C:\Users\GeoModel\Desktop\Povratnice\" + Trim(Str(broj_datoteka)) + ".pdf")

            doc.Close()


            datoteka2 = "C:\Users\GeoModel\Desktop\Povratnice\" + Trim(Str(broj_datoteka)) + ".pdf"



            files2(broj_datoteka - 1) = datoteka2




            '


            ' MsgBox(files2(broj_datoteka - 1))




            '{ "Sample1.pdf", "Sample2.pdf", "Sample3.pdf" }


            postotak = ((broj_datoteka / (broj_jpg + 2))) * 100

            ProgressBar1.Value = broj_datoteka
            Label1.Text = postotak


        Next











        Dim outputfile As String = "C:\Users\GeoModel\Desktop\Povratnice\" + "spojeno" + ".pdf"
        Dim doc2 As PdfDocumentBase = PdfDocument.MergeFiles(files2)
        doc2.Save(outputfile, FileFormat.PDF)




        postotak = (((broj_datoteka + 1) / (broj_jpg + 2))) * 100

        ProgressBar1.Value = broj_datoteka + 1
        Label1.Text = postotak






        System.Diagnostics.Process.Start("C:\Users\GeoModel\Desktop\Povratnice\" + "spojeno" + ".pdf")

        ' MsgBox("")












    End Sub







    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
