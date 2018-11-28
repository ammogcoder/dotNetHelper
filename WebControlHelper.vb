Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Reflection
Imports System.Drawing
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Data
Imports System.Data.Linq
Imports System.Web.UI

Namespace WebControls
    ''' <summary>
    ''' ASP.NET Web Helper class for adding more methods to the web controls
    ''' </summary>
    ''' <remarks></remarks>
    Public Module WebControlHelper
        ''' <summary>
        ''' Reads a supported picture file from file upload to byte array
        ''' </summary>
        ''' <param name="FU">Target control</param>
        ''' <param name="picdata">Destination byte array</param>
        ''' <param name="ImgSize">Required imgsize</param>
        ''' <param name="MaxFileSize">Maximum allowed size of file in byte</param>
        ''' <param name="ReduceSize">Indicates if file should be reduced with specified imgSize</param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub GetPicture(ByVal FU As FileUpload, ByRef picdata As Byte(), ByRef ImgSize As Size, Optional ByVal MaxFileSize As Long = 0, Optional ReduceSize As Boolean = False)
            Dim data() As Byte
            Dim MIMEType As String = ""
            Dim extension As String = ""
            Dim filename As String = ""
            If FU.PostedFile Is Nothing OrElse String.IsNullOrEmpty(FU.PostedFile.FileName) OrElse FU.PostedFile.InputStream Is Nothing Then
                data = New Byte() {0} 'Return New Byte() {0}
                Throw New Exception("Pls Select A Valid Picture File")
            Else
                extension = Path.GetExtension(FU.PostedFile.FileName).ToLower()
                Select Case extension
                    Case ".gif"
                        MIMEType = "image/gif"
                    Case ".jpg", ".jpeg", ".jpe"
                        MIMEType = "image/jpeg"
                    Case ".png"
                        MIMEType = "image/png"
                    Case Else
                        Throw New Exception("Only GIF, JPG, and PNG files can be uploaded to the picture album.")
                        data = New Byte() {0} 'Return New Byte() {0}
                        Throw New Exception("Pls Select A Valid Picture File")
                        'Exit Function
                End Select
                Dim imageBytes(FU.PostedFile.InputStream.Length) As Byte
                FU.PostedFile.InputStream.Read(imageBytes, 0, imageBytes.Length)
                data = imageBytes

                'check the size
                If MaxFileSize > 0 Then If data.Length > MaxFileSize Then Throw New Exception("File size is more than " & (MaxFileSize / 1024) & "KB")

                Dim mem As New MemoryStream(data)
                Dim img As System.Drawing.Image = System.Drawing.Image.FromStream(mem)
                ImgSize = New Size(img.Width, img.Height)

                Dim reducedimagedata As Byte()
                Select Case ReduceSize
                    Case False
                        Dim dummyCallBack As System.Drawing.Image.GetThumbnailImageAbort
                        dummyCallBack = New  _
                          System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)

                        Dim thumbNailImg As System.Drawing.Image
                        thumbNailImg = img.GetThumbnailImage(img.Width, img.Height, _
                                                                 dummyCallBack, IntPtr.Zero)
                        reducedimagedata = saveImageToarray(thumbNailImg, getimageformat(MIMEType))
                        picdata = reducedimagedata
                    Case True
                        Dim dummyCallBack As System.Drawing.Image.GetThumbnailImageAbort
                        dummyCallBack = New  _
                          System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)

                        Dim thumbNailImg As System.Drawing.Image
                        thumbNailImg = img.GetThumbnailImage(115, 135, _
                                                             dummyCallBack, IntPtr.Zero)
                        reducedimagedata = saveImageToarray(thumbNailImg, getimageformat(MIMEType))

                        picdata = reducedimagedata
                End Select

            End If

        End Sub
        Private Function saveImageToarray(ByVal img As System.Drawing.Image, ByVal format As Imaging.ImageFormat) As Byte()
            Dim ms As New MemoryStream()
            img.Save(ms, format)
            Return ms.ToArray()
        End Function
        Private Function getimageformat(ByVal format As String) As Imaging.ImageFormat
            Select Case format
                Case "image/gif"
                    Return Imaging.ImageFormat.Gif
                Case "image/jpeg"
                    Return Imaging.ImageFormat.Jpeg
                Case "image/png"
                    Return Imaging.ImageFormat.Png
                Case Else
                    Return Nothing
            End Select
        End Function
        Private Function ThumbnailCallback() As Boolean
            Return False
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub CheckAll(source As CheckBoxList, Value As Boolean)
            Try
                For i As Integer = 0 To source.Items.Count - 1
                    source.Items(i).Selected = Value
                Next
            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function HasItemSelected(source As CheckBoxList) As Boolean
            Try
                For i As Integer = 0 To source.Items.Count - 1
                    If source.Items(i).Selected Then
                        Return True : Exit For
                    End If
                Next
                Return False
            Catch ex As Exception
            End Try
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <param name="Values"></param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub SelectValues(source As CheckBoxList, Values() As String)
            Try
                For j As Integer = 0 To Values.Length - 1
                    For i As Integer = 0 To source.Items.Count - 1
                        If source.Items(i).Value.Trim = Values(j).Trim Then
                            source.Items(i).Selected = True : Exit For
                        End If
                    Next
                Next
            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function GetCheckedValues(source As CheckBoxList)
            Try
                Dim val As String = ""
                For i As Integer = 0 To source.Items.Count - 1
                    If source.Items(i).Selected Then
                        val &= "|" & source.Items(i).Value
                    End If
                Next
                Dim out() As String
                If val.Length > 0 Then out = val.Substring(1).Split("|") Else out = New String() {}
                Return out
            Catch ex As Exception
            End Try
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <param name="itm"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function HasItem(source As CheckBoxList, itm As String) As Boolean
            Try
                For i As Integer = 0 To source.Items.Count - 1
                    If source.Items(i).Value.ToLower.Trim = itm Then
                        Return True : Exit For
                    End If
                Next
                Return False
            Catch ex As Exception
            End Try
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function HasItem(source As DropDownList) As Boolean
            Try
                Return source IsNot Nothing AndAlso source.Items.Count > 0
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Chk"></param>
        ''' <param name="dt"></param>
        ''' <param name="Value"></param>
        ''' <param name="Display"></param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub FormatForChkList(ByVal Chk As CheckBoxList, ByVal dt As DataTable, ByVal Value As String, ByVal Display As String)
            Try
                Chk.DataSource = dt
                Chk.DataValueField = Value
                Chk.DataTextField = Display
                Chk.DataBind()
            Catch ex As Exception
                Throw ex
            End Try
        End Sub


        ''' <summary>
        ''' Returns a flag of empty object in table, gridview, etc
        ''' </summary>
        ''' <param name="dg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function HasRows(ByVal dg As Object) As Boolean
            Try
                If dg.GetType Is GetType(DataTable) Or dg.GetType Is GetType(GridView) Then
                    If dg IsNot Nothing AndAlso dg.rows.count > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Throw New Exception("Invalid Object")
                End If
            Catch ex As Exception
                Throw ex
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Lst"></param>
        ''' <param name="dt"></param>
        ''' <param name="Value"></param>
        ''' <param name="Display"></param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub FormatForList(ByVal Lst As ListBox, ByVal dt As DataTable, ByVal Value As String, ByVal Display As String)
            Try
                Lst.DataSource = dt
                Lst.DataValueField = Value
                Lst.DataTextField = Display
                Lst.DataBind()
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="ddl"></param>
        ''' <param name="dt"></param>
        ''' <param name="Value"></param>
        ''' <param name="Display"></param>
        ''' <param name="DefaultDisplay"></param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub FormatForCombo(ByVal ddl As DropDownList, ByVal dt As DataTable, ByVal Value As String, ByVal Display As String, ByVal DefaultDisplay As String)
            Try
                ddl.DataSource = dt
                ddl.DataValueField = Value
                ddl.DataTextField = Display
                ddl.DataBind()
                If DefaultDisplay.Trim <> "" Then
                    ddl.Items.Insert(0, DefaultDisplay)
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="URL"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function QueryUrl(URL) As String
            Try
                Dim request As HttpWebRequest = CType(WebRequest.Create(URL), HttpWebRequest)
                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim sr As New StreamReader(response.GetResponseStream)
                Dim out As String = sr.ReadToEnd
                response.Close()
                sr.Close()
                Return out
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="ctrl"></param>
        ''' <remarks></remarks>
        Public Sub ClearControls(ByVal ctrl As Control)
            Try
                Dim ctl As Control
                For Each ctl In ctrl.Controls
                    If ctl.GetType Is GetType(TextBox) Then
                        CType(ctl, TextBox).Text = ""
                    End If 
                    If ctl.GetType Is GetType(HiddenField) Then
                        CType(ctl, HiddenField).Value = ""
                    End If
                    If ctl.GetType Is GetType(DropDownList) Then
                        If CType(ctl, DropDownList).Items.Count > 0 Then
                            CType(ctl, DropDownList).SelectedIndex = 0
                        End If
                    End If

                    If ctl.Controls.Count > 0 Then
                        ClearControls(ctl)
                    End If
                Next
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' Proper case the content of a textbox in a control or standing alone
        ''' </summary>
        ''' <param name="ctrl">Target control</param>
        ''' <remarks></remarks>
        Public Sub DoProperCase(ByVal ctrl As Control)
            Try 
                Dim ctl As Control
                If ctrl.HasControls Then
                    For Each ctl In ctrl.Controls

                        If ctl.GetType Is GetType(TextBox) Then
                            If Not (CType(ctl, TextBox).TextMode = TextBoxMode.Password) Then
                                CType(ctl, TextBox).Text = StrConv(CType(ctl, TextBox).Text, VbStrConv.ProperCase)
                            End If
                        End If

                        If ctl.Controls.Count > 0 Then
                            DoProperCase(ctl)
                        End If
                    Next
                Else
                    If ctrl.GetType Is GetType(TextBox) Then 
                        If Not (CType(ctrl, TextBox).TextMode = TextBoxMode.Password) Then
                            CType(ctrl, TextBox).Text = StrConv(CType(ctrl, TextBox).Text, VbStrConv.ProperCase)
                        End If
                    End If
                End If

            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' Hides all columms of a gridview
        ''' </summary>
        ''' <param name="dgvlist">Target Control</param>
        ''' <remarks></remarks> 
        <Extension()> Sub BlindColumns(ByVal dgvlist As GridView)
            For i As Integer = 0 To dgvlist.Columns.Count - 1
                dgvlist.Columns(i).Visible = False
            Next
        End Sub

        ''' <summary>
        ''' Gee the login url from any location
        ''' </summary>
        ''' <param name="relativeUrl"></param>
        ''' <param name="query"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLogoutUrl(ByVal relativeUrl As String, ByVal query As String) As String
            Dim root As String = HttpContext.Current.ApplicationInstance.Context.Request.Url.GetLeftPart _
                                (UriPartial.Authority)
            Dim pathx As String = VirtualPathUtility.ToAbsolute(relativeUrl)
            'Dim pathx As String = VirtualPathUtility.ToAbsolute("~/Default.aspx")
            pathx = pathx.Remove(pathx.IndexOf("/"), 1)
            Dim str As String = String.Format("{0}/{1}", root, pathx) & query
            Return str
        End Function
    End Module

End Namespace
