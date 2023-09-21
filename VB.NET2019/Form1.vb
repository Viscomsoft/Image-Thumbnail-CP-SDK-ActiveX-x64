Public Class Form1

    Dim oldBackColor As Color
    Dim oldBorderColor As Color
    Dim oldShadowColor As Color
    Dim oldClipColor As Color

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        cbofilter.Items.Add("JPG|BMP|GIF|PNG")
        cbofilter.Items.Add("JPG|BMP|GIF|PNG|PSD|TIF")
        cbofilter.Items.Add("PDF|TIF")
        cbofilter.Items.Add("All Supported Image formats")
        cbofilter.SelectedIndex = 3

        cboExportImageType.Items.Add("BMP")
        cboExportImageType.Items.Add("JPEG")
        cboExportImageType.Items.Add("TIF")
        cboExportImageType.Items.Add("PNG")
        cboExportImageType.Items.Add("GIF")
        cboExportImageType.SelectedIndex = 0




        oldBackColor = AxImageThumbnailCP1.BackgroundColor
        oldBorderColor = AxImageThumbnailCP1.ClipBorderColor
        oldShadowColor = AxImageThumbnailCP1.ClipShadowColor
        oldClipColor = AxImageThumbnailCP1.ClipColor

        cboscrollbar.Items.Add("Vertical")
        cboscrollbar.Items.Add("Horizontal")
        cboscrollbar.SelectedIndex = 0

        cbodisplaymode.Items.Add("Fast Speed, poor Quality")
        cbodisplaymode.Items.Add("Normal Speed, Normal Quality")
        cbodisplaymode.Items.Add("Slow Speed, High Quality")
        cbodisplaymode.Items.Add("Very Slow Speed, Best Quality")
        cbodisplaymode.SelectedIndex = 1

    End Sub

    Private Sub cboscrollbar_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboscrollbar.SelectedIndexChanged

        AxImageThumbnailCP1.ScrollBarType = cboscrollbar.SelectedIndex

        If cboscrollbar.SelectedIndex = 0 Then
            AxImageThumbnailCP1.Height = 328
        Else

            AxImageThumbnailCP1.Height = 150


        End If
        AxImageThumbnailCP1.Refresh()


    End Sub

    Private Sub cbodisplaymode_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbodisplaymode.SelectedIndexChanged
        AxImageThumbnailCP1.DisplayMode = cbodisplaymode.SelectedIndex
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim strpath As String

        strpath = AxImageThumbnailCP1.BrowseFolder("select folder")

        Select Case cbofilter.SelectedIndex

            Case 0
                Me.AxImageThumbnailCP1.AddClipsFromFolder(strpath, "jpg|bmp|gif|png")
            Case 1
                Me.AxImageThumbnailCP1.AddClipsFromFolder(strpath, "jpg|bmp|gif|png|psd|tif")
            Case 2
                Me.AxImageThumbnailCP1.AddClipsFromFolder(strpath, "pdf|tif")
            Case 3
                Me.AxImageThumbnailCP1.AddClipsFromFolder(strpath, "psd|pdf|tif|pcx|png|gif|tga|wbmp|wmf|j2k|jp2|j2c|pgx|ras|pmn|ico|bmp|jpg|png")
        End Select
    End Sub

    Private Sub btnadd_Click(sender As System.Object, e As System.EventArgs) Handles btnadd.Click
        OpenFileDialog1.Filter = "All Files (*.*)|*.*|PDF (*.pdf)|*.pdf|PhotoShop (*.psd)|*.psd|JPEG 2000 (*.j2k)|*.j2k;*.j2c|JPEG (*.jpg)|*.jpg|PCX (*.pcx)|*.pcx|WMF (*.wmf)|*.wmf|Wireless Bitmap (*.wbmp)|*.wbmp|Bitmap (*.bmp)|*.bmp|TIF (*.tif)|*.tif|TGA (*.tga)|*.tga|Gif (*.gif)|*.gif |PGX (*.pgx)|*.pgx|RAS (*.ras)|*.ras|PNM (*.pnm)|*.pnm|PNG (*.png)|*.png|Icon (*.ico)|*.ico||"

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            AxImageThumbnailCP1.UnSelectAllClip()
            AxImageThumbnailCP1.AddClip(OpenFileDialog1.FileName, "")
            AxImageThumbnailCP1.SelectClip(AxImageThumbnailCP1.ClipCount - 1)
            AxImageThumbnailCP1.VScroll(7)
            AxImageThumbnailCP1.Focus()
        End If
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        Dim strPDFProperty As String
        Dim iCurIndex As Integer

        iCurIndex = AxImageThumbnailCP1.GetCurSel

        strPDFProperty = "PDF Resolution:" + Str(Me.AxImageThumbnailCP1.GetClipPDFWidth(iCurIndex)) + "x" + Str(Me.AxImageThumbnailCP1.GetClipPDFHeight(iCurIndex)) + Chr(13) + Chr(10)
        strPDFProperty = strPDFProperty + "PDF Title:" + Me.AxImageThumbnailCP1.GetClipPDFTitle(iCurIndex) + Chr(13) + Chr(10)
        strPDFProperty = strPDFProperty + "PDF Subject:" + Me.AxImageThumbnailCP1.GetClipPDFSubject(iCurIndex) + Chr(13) + Chr(10)
        strPDFProperty = strPDFProperty + "PDF Author:" + Me.AxImageThumbnailCP1.GetClipPDFAuthor(iCurIndex) + Chr(13) + Chr(10)
        strPDFProperty = strPDFProperty + "PDF Creation Date:" + Me.AxImageThumbnailCP1.GetClipPDFCreationDate(iCurIndex) + Chr(13) + Chr(10)
        strPDFProperty = strPDFProperty + "PDF Modify Date:" + Me.AxImageThumbnailCP1.GetClipPDFModifyDate(iCurIndex) + Chr(13) + Chr(10)
        strPDFProperty = strPDFProperty + "PDF Version No:" + Str(Me.AxImageThumbnailCP1.GetClipPDFVersionNo(iCurIndex)) + Chr(13) + Chr(10)
        strPDFProperty = strPDFProperty + "PDF Keyword:" + Me.AxImageThumbnailCP1.GetClipPDFKeyword(iCurIndex) + Chr(13) + Chr(10)
        strPDFProperty = strPDFProperty + "Total Page:" + Str(Me.AxImageThumbnailCP1.GetClipPDFTotalPage(iCurIndex)) + Chr(13) + Chr(10)

        MessageBox.Show(strPDFProperty)
    End Sub

    Private Sub btnbgcolor_Click(sender As System.Object, e As System.EventArgs) Handles btnbgcolor.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            AxImageThumbnailCP1.BackgroundColor = ColorDialog1.Color
            AxImageThumbnailCP1.Focus()
        End If
    End Sub

    Private Sub btnclipcolor_Click(sender As System.Object, e As System.EventArgs) Handles btnclipcolor.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            AxImageThumbnailCP1.ClipColor = ColorDialog1.Color
            AxImageThumbnailCP1.Focus()
        End If
    End Sub

    Private Sub btnhighlightcolor_Click(sender As System.Object, e As System.EventArgs) Handles btnhighlightcolor.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            AxImageThumbnailCP1.ClipHighlightColor = ColorDialog1.Color
            AxImageThumbnailCP1.Focus()
        End If

    End Sub

    Private Sub btnbordercolor_Click(sender As System.Object, e As System.EventArgs) Handles btnbordercolor.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            AxImageThumbnailCP1.ClipBorderColor = ColorDialog1.Color
            AxImageThumbnailCP1.Focus()
        End If

    End Sub

    Private Sub btnshadowcolor_Click(sender As System.Object, e As System.EventArgs) Handles btnshadowcolor.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            AxImageThumbnailCP1.ClipShadowColor = ColorDialog1.Color
            AxImageThumbnailCP1.Focus()
        End If
    End Sub

    Private Sub btntextcolor_Click(sender As System.Object, e As System.EventArgs) Handles btntextcolor.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            AxImageThumbnailCP1.ClipTextColor = ColorDialog1.Color
            AxImageThumbnailCP1.Focus()
        End If

    End Sub

    Private Sub btnupdatetext_Click(sender As System.Object, e As System.EventArgs) Handles btnupdatetext.Click
        Dim iCount As Integer
        Dim i As Integer

        iCount = AxImageThumbnailCP1.ClipCount


        For i = 0 To iCount
            AxImageThumbnailCP1.SetClipInfo(i, "", " Demo " + Convert.ToString(i))
        Next

        AxImageThumbnailCP1.Focus()
    End Sub

    Private Sub btnacdsee_Click(sender As System.Object, e As System.EventArgs) Handles btnacdsee.Click
        Dim Color1 As Color = Color.FromArgb(128, 128, 128)
        Dim iCount As Integer
        Dim i As Integer

        AxImageThumbnailCP1.BackgroundColor = Color1
        AxImageThumbnailCP1.ClipBorderColor = Color1
        AxImageThumbnailCP1.ClipShadowColor = Color1
        AxImageThumbnailCP1.ClipColor = Color1
        AxImageThumbnailCP1.ClipFontSize = 10
        AxImageThumbnailCP1.ClipFontName = "Arial"
        AxImageThumbnailCP1.ClipFontTopPos = 95
        AxImageThumbnailCP1.ClipTopMargin = 0
        AxImageThumbnailCP1.ClipBottomMargin = 15
        AxImageThumbnailCP1.ClipWidth = 110
        AxImageThumbnailCP1.ClipHeight = 120

        iCount = AxImageThumbnailCP1.ClipCount

        For i = 0 To iCount
            AxImageThumbnailCP1.SetClipInfo(i, "", " Demo " + Convert.ToString(i))
        Next

        AxImageThumbnailCP1.Focus()
    End Sub

    Private Sub btnnormal_Click(sender As System.Object, e As System.EventArgs) Handles btnnormal.Click
        AxImageThumbnailCP1.BackgroundColor = oldBackColor
        AxImageThumbnailCP1.ClipBorderColor = oldBorderColor
        AxImageThumbnailCP1.ClipShadowColor = oldShadowColor
        AxImageThumbnailCP1.ClipColor = oldClipColor
    End Sub

    Private Sub Button14_Click(sender As System.Object, e As System.EventArgs) Handles Button14.Click
        Dim iCount As Integer
        Dim i As Integer

        iCount = AxImageThumbnailCP1.ClipCount

        For i = 0 To iCount
            AxImageThumbnailCP1.DeleteClipByIndex(0)
        Next

        AxImageThumbnailCP1.Focus()
    End Sub

    Private Sub Button15_Click(sender As System.Object, e As System.EventArgs) Handles Button15.Click
        AxImageThumbnailCP1.DeleteSelectedClip()
        AxImageThumbnailCP1.Focus()
    End Sub

    Private Sub chkfontbold_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkfontbold.CheckedChanged
        If chkfontbold.Checked Then
            AxImageThumbnailCP1.ClipFontBold = True
        Else
            AxImageThumbnailCP1.ClipFontBold = False

        End If
    End Sub

    Private Sub chkfontItalic_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkfontItalic.CheckedChanged

        If chkfontItalic.Checked Then
            AxImageThumbnailCP1.ClipFontItalic = True
        Else
            AxImageThumbnailCP1.ClipFontItalic = False

        End If
    End Sub

    Private Sub chkfontunderline_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkfontunderline.CheckedChanged
        If chkfontunderline.Checked Then
            AxImageThumbnailCP1.ClipFontUnderline = True
        Else
            AxImageThumbnailCP1.ClipFontUnderline = False
        End If
    End Sub

    Private Sub btnmargin_Click(sender As System.Object, e As System.EventArgs) Handles btnmargin.Click
        AxImageThumbnailCP1.ClipTopMargin = txttopmargin.Text
        AxImageThumbnailCP1.ClipBottomMargin = txtbottommargin.Text
    End Sub

    Private Sub btnfontsize_Click(sender As System.Object, e As System.EventArgs) Handles btnfontsize.Click
        AxImageThumbnailCP1.ClipFontSize = txtfontsize.Text
    End Sub

    Private Sub btnfonttop_Click(sender As System.Object, e As System.EventArgs) Handles btnfonttop.Click
        AxImageThumbnailCP1.ClipFontTopPos = txtFontTop.Text
    End Sub

    Private Sub Button13_Click(sender As System.Object, e As System.EventArgs) Handles Button13.Click
        AxImageThumbnailCP1.ClipWidth = txtClipWidth.Text
        AxImageThumbnailCP1.ClipHeight = txtClipHeight.Text
    End Sub

    Private Sub Button16_Click(sender As System.Object, e As System.EventArgs) Handles Button16.Click
        Dim i As Integer

        For i = 0 To AxImageThumbnailCP1.ClipCount

            If AxImageThumbnailCP1.GetSelect(i) Then
                MessageBox.Show(AxImageThumbnailCP1.GetClipFilePath(i))
            End If
        Next
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        AxImageThumbnailCP1.UnSelectAllClip()
        AxImageThumbnailCP1.SelectClip(0)
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        AxImageThumbnailCP1.UnSelectAllClip()
        AxImageThumbnailCP1.SelectClip(1)
        AxImageThumbnailCP1.SelectClip(2)
        AxImageThumbnailCP1.SelectClip(3)
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        AxImageThumbnailCP1.UnSelectClip(1)
        AxImageThumbnailCP1.UnSelectClip(2)
        AxImageThumbnailCP1.UnSelectClip(3)
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Dim iFirstClip As Integer
        Dim i As Integer
        Dim iTotalVisibleClip As Integer
        Dim iTotalClip As Integer

        iFirstClip = AxImageThumbnailCP1.CurrentVisibleFirstClipIndex

        iTotalVisibleClip = AxImageThumbnailCP1.CurrentVisibleClipCount

        iTotalClip = iFirstClip + iTotalVisibleClip - 1

        For i = iFirstClip To iTotalClip
            MessageBox.Show(AxImageThumbnailCP1.GetClipFilePath(i))
        Next
    End Sub

    Private Sub AxImageThumbnailCP1_ClickEvent(sender As System.Object, e As AxIMAGETHUMBNAILCPLib._DImageThumbnailCPEvents_ClickEvent) Handles AxImageThumbnailCP1.ClickEvent
        MyThubnailClick(e.iClipIndex, e.strFilePath, e.iPageNo)
    End Sub

    Private Sub MyThubnailClick(ByVal iClipIndex As Short, ByVal strFilePath As String, ByVal iPageNo As Short)


        txtCurrentIndex.Text = iClipIndex

        txtFileName.Text = strFilePath


        txtPageNo.Text = iPageNo


        AxImageThumbnailCP1.Focus()
    End Sub

    Private Sub button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button8.Click
        Dim iCurIndex As Integer

        iCurIndex = AxImageThumbnailCP1.GetCurSel()

        If (cboExportImageType.SelectedIndex = 0) Then
            SaveFileDialog1.Filter = "BMP File (*.bmp)|*.bmp"
        ElseIf (cboExportImageType.SelectedIndex = 1) Then
            SaveFileDialog1.Filter = "JPEG File (*.jpg)|*.jpg"
        ElseIf (cboExportImageType.SelectedIndex = 2) Then
            SaveFileDialog1.Filter = "TIF File (*.tif)|*.tif"
        ElseIf (cboExportImageType.SelectedIndex = 3) Then
            SaveFileDialog1.Filter = "PNG File (*.png)|*.png"
        ElseIf (cboExportImageType.SelectedIndex = 4) Then
            SaveFileDialog1.Filter = "GIF File (*.gif)|*.gif"

        End If

        

        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then

            AxImageThumbnailCP1.ExportImageClipByIndex(iCurIndex, SaveFileDialog1.FileName)

        End If

    End Sub
End Class
