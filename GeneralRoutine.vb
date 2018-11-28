Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Reflection
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System.Net.NetworkInformation
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports System.Drawing.Text
Imports Microsoft.VisualBasic 
Imports System.Math

Namespace Core
    ''' <summary>
    ''' Core Helper class for general useful routines on various objects e.g strings, array etc.
    ''' </summary>
    ''' <remarks></remarks>
    Public Module CoreHelper


        ''' <summary>
        ''' Check if a string has number and character
        ''' </summary>
        ''' <param name="str"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsAlphaNumeric(ByVal str As String) As Boolean
            Try
                Dim hasNumber As Boolean = False
                Dim hasAlpha As Boolean = False

                For Each c As Char In str

                    If Char.IsNumber(c) Then
                        hasNumber = True
                    ElseIf Char.IsLetter(c) Then
                        hasAlpha = True
                    End If

                    If hasAlpha And hasNumber Then
                        Return True
                    End If
                Next

                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Formats string as propercase
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToProper(source As String) As String
            Try
                Return StrConv(source, VbStrConv.ProperCase)
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Reads through the string array and combine all contents as string separated with space or other separator
        ''' </summary>
        ''' <param name="aString"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToStringConcatenate(ByVal aString() As String, Optional ByVal Separator As String = " ") As String
            Dim out As String = ""
            For Each a In aString
                out &= Separator & a
            Next
            If out.Trim.Length > 1 Then Return out.Substring(1) Else If out.Trim.Trim.Length = 1 Then Return out
        End Function

        ''' <summary>
        ''' Returns a shuffled array of string 
        ''' </summary>
        ''' <param name="texts"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ShuffleString(ByVal texts() As String) As String()
            Dim random As New Random()
            For t As Integer = 0 To texts.Length - 1
                Dim tmp As String = texts(t)
                Dim r As Integer = random.Next(t, texts.Length)
                texts(t) = texts(r)
                texts(r) = tmp
            Next
            Return texts
        End Function

        ''' <summary>
        ''' Returns a shuffled array of T
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="array"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function Shuffle(Of T)(ByVal array As T())
            Dim random As New Random()

            For i As Integer = 0 To array.Length - 1
                Dim idx As Integer = random.[Next](i, array.Length)

                'swap elements
                Dim tmp As T = array(i)
                array(i) = array(idx)
                array(idx) = tmp
            Next
            Return array
        End Function

        ''' <summary>
        ''' Reads through the string array and combine all contents as html List
        ''' </summary>
        ''' <param name="aString"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToHtmlList(ByVal aString() As String) As String
            Dim out As String = "<ul>"
            For Each a In aString
                out &= "<li>" & a & "</li>"
            Next
            Return out & "</ul>"
        End Function

        ''' <summary>
        ''' Checks if a remote computer is reachable
        ''' </summary>
        ''' <param name="ComputerID">Qualified identity of the computer</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function PingComputer(ComputerID As String) As Boolean
            Try
                Dim objPIng As New Ping
                Dim pngReply As PingReply = objPIng.Send(ComputerID.Trim(), 3000)
                If (pngReply.Status = IPStatus.Success) Then
                    Return True
                End If
                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Checks if a web URL is reachable
        ''' </summary>
        ''' <param name="WebURL">Web URL to be checked</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function PingWebURL(WebURL As String) As Boolean
            Try
                Dim Http_URL As String = HttpUtility.UrlDecode(WebURL)

                Dim lObjReq As HttpWebRequest = CType(HttpWebRequest.Create(Http_URL), HttpWebRequest)

                lObjReq.Timeout = 36000
                lObjReq.UseDefaultCredentials = True

                If (lObjReq.GetResponse().ContentLength > 0) Then Return True
                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Strips off html tags from a string and return plain string
        ''' </summary>
        ''' <param name="source">String to be processed</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function StripHtml(source As String) As String
            Dim output As String

            'get rid of HTML tags
            output = Regex.Replace(source, "<[^>]*>", String.Empty)

            'get rid of multiple blank lines
            output = Regex.Replace(output, "^\s*$\n", String.Empty, RegexOptions.Multiline)

            Return output
        End Function

        ''' <summary>
        ''' Converts an array of string to datatable
        ''' </summary>
        ''' <param name="source">Target array</param>
        ''' <param name="ColumnName">Optional column name for the datatable</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToDatatable(source() As String, Optional ByVal ColumnName As String = "Item") As DataTable
            Try
                Dim dt As New DataTable : dt.Columns.Add(ColumnName)
                For Each s In source
                    dt.Rows.Add(s)
                Next
                Return dt
            Catch ex As Exception
                Return Nothing
            End Try
        End Function


#Region "DateTime Extension"
        ''' <summary>
        '''  Gets the last day of a specified datetime
        ''' </summary>
        ''' <param name="MonthDate">Target datetime</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> _
        Public Function GetLastDayOfMonth(ByVal MonthDate As DateTime) As Date
            Dim FirstDayInMonth = New DateTime(MonthDate.Year, MonthDate.Month, 1)
            Return FirstDayInMonth.AddMonths(1).AddDays(-1)
        End Function

#End Region

        ''' <summary>
        ''' Compares if two array of byte are the same
        ''' </summary>
        ''' <param name="Array1">First array</param>
        ''' <param name="Array2">Second array</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CompareByteArrays(ByVal Array1() As Byte, ByVal Array2() As Byte) As Boolean
            If Array1.Length <> Array2.Length Then
                Return False
            End If

            For i As Integer = 0 To Array1.Length - 1
                If Array1(i) <> Array2(i) Then
                    Return False
                End If
            Next

            Return True
        End Function

        ''' <summary>
        ''' Generates Random character of specified length (length must no be more than 32)
        ''' </summary>
        ''' <param name="maxSize"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUniqueKeySlim(maxSize As Integer) As String
            If maxSize > 32 Then
                Throw New Exception("Length has exceeded limit")
            Else
                Return Guid.NewGuid().ToString("N").Replace("-", String.Empty).Substring(0, maxSize)
            End If
        End Function

        ''' <summary>
        ''' Generates Random alpha-numeric characters of specified length
        ''' </summary>
        ''' <param name="maxSize"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUniqueKey(maxSize As Integer) As String
            Dim chars As Char() = New Char(61) {}
            chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890".ToCharArray()
            Dim data As Byte() = New Byte(0) {}
            Dim crypto As New RNGCryptoServiceProvider()
            crypto.GetNonZeroBytes(data)
            data = New Byte(maxSize - 1) {}
            crypto.GetNonZeroBytes(data)
            Dim result As New StringBuilder(maxSize)
            For Each b As Byte In data
                result.Append(chars(b Mod (chars.Length)))
            Next
            Return result.ToString()
        End Function

        ''' <summary>
        ''' Writes a file to disk with hidden attribute
        ''' </summary>
        ''' <param name="FileName">File name to be written</param>
        ''' <param name="data">Binary to be written</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function WriteBinaryAsFileHidden(ByVal FileName As String, ByVal data As Byte()) As Boolean
            Dim FilePath As String = FileName
            Try
                ' new file, create an empty file
                File.Create(FilePath).Close()
                Dim f As New FileInfo(FilePath)
                f.Attributes = FileAttributes.Hidden

                ' open a file stream and write the buffer.  Don't open with FileMode.Append because the transfer may wish to start a different point
                Using fs As New FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read)
                    fs.Seek(0, SeekOrigin.Begin)
                    fs.Write(data, 0, data.Length)
                End Using
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Writes a binary to disk
        ''' </summary>
        ''' <param name="FileName">File name to be written</param>
        ''' <param name="data">Binary to be written</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function WriteBinaryAsFile(ByVal FileName As String, ByVal data As Byte()) As Boolean
            Dim FilePath As String = FileName
            Try
                ' new file, create an empty file
                File.Create(FilePath).Close()

                ' open a file stream and write the buffer.  Don't open with FileMode.Append because the transfer may wish to start a different point
                Using fs As New FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read)
                    fs.Seek(0, SeekOrigin.Begin)
                    fs.Write(data, 0, data.Length)
                End Using
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Module

    ''' <summary>
    ''' Converts a string to bitmap for captcha rendering
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SimpleCaptcha
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sImageText"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CreateBitmapImage(ByVal sImageText As String) As Bitmap
            ' For generating random numbers.
            Dim random As Random
            random = New Random()

            Dim objBmpImage As New Bitmap(1, 1)
            Dim intWidth As Integer = 0

            Dim intHeight As Integer = 0
            ' Create the Font object for the image text drawing.
            Dim objFont As New Font("Jokerman", 22, System.Drawing.FontStyle.Italic + FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel)

            ' Create a graphics object to measure the text's width and height.
            Dim objGraphics As Graphics = Graphics.FromImage(objBmpImage)

            ' This is where the bitmap size is determined.
            intWidth = CInt(objGraphics.MeasureString(sImageText, objFont).Width)

            intHeight = CInt(objGraphics.MeasureString(sImageText, objFont).Height)

            ' Create the bmpImage again with the correct size for the text and font.
            objBmpImage = New Bitmap(objBmpImage, New Size(intWidth, intHeight))

            ' Add the colors to the new bitmap.
            objGraphics = Graphics.FromImage(objBmpImage)

            objGraphics.Clear(Color.White)

            objGraphics.SmoothingMode = SmoothingMode.AntiAlias

            objGraphics.TextRenderingHint = TextRenderingHint.AntiAlias

            objGraphics.DrawString(sImageText, objFont, New SolidBrush(Color.FromArgb(102, 102, 102)), 0, 0)
            Dim HatchBrush As New HatchBrush(HatchStyle.LightUpwardDiagonal, Color.Black, Color.DarkGray)
            'add some noise 
            Dim m As Integer = Math.Max(intWidth, intHeight)
            For i As Integer = 0 To ((intWidth * intHeight) / 30)
                Dim x As Integer = random.Next(intWidth)
                Dim y As Integer = random.Next(intHeight)
                Dim w As Integer = random.Next(m / 50)
                Dim h As Integer = random.Next(m / 50)
                objGraphics.FillEllipse(HatchBrush, x, y, w, h)
            Next

            objGraphics.Flush()

            Return (objBmpImage)

        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sImageText"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CreateBitmapImageWithoutNoise(ByVal sImageText As String) As Bitmap

            Dim objBmpImage As New Bitmap(1, 1)
            Dim intWidth As Integer = 0

            Dim intHeight As Integer = 0
            ' Create the Font object for the image text drawing.
            Dim objFont As New Font("Arial", 17, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel)

            ' Create a graphics object to measure the text's width and height.
            Dim objGraphics As Graphics = Graphics.FromImage(objBmpImage)

            ' This is where the bitmap size is determined.
            intWidth = CInt(objGraphics.MeasureString(sImageText, objFont).Width)

            intHeight = CInt(objGraphics.MeasureString(sImageText, objFont).Height)

            ' Create the bmpImage again with the correct size for the text and font.
            objBmpImage = New Bitmap(objBmpImage, New Size(intWidth, intHeight))

            ' Add the colors to the new bitmap.
            objGraphics = Graphics.FromImage(objBmpImage)

            objGraphics.Clear(Color.White)

            objGraphics.SmoothingMode = SmoothingMode.AntiAlias

            objGraphics.TextRenderingHint = TextRenderingHint.AntiAlias

            objGraphics.DrawString(sImageText, objFont, New SolidBrush(Color.FromArgb(102, 102, 102)), 0, 0)

            objGraphics.Flush()

            Return (objBmpImage)

        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="txt"></param>
        ''' <param name="font_family_name"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MakeCaptchaImge(ByVal txt As String, ByVal font_family_name As String) As Bitmap
            ' Make the bitmap and associated Graphics object.
            Dim wid As Integer, hgt As Integer

            'derive the width and height
            Dim objBmpImage As New Bitmap(1, 1)
            ' Create the Font object for the image text drawing.
            Dim objFont As New Font(font_family_name, 17, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel)

            ' Create a graphics object to measure the text's width and height.
            Dim objGraphics As Graphics = Graphics.FromImage(objBmpImage)

            ' This is where the bitmap size is determined.
            wid = CInt(objGraphics.MeasureString(txt, objFont).Width)
            hgt = CInt(objGraphics.MeasureString(txt, objFont).Height)

            Dim bm As New Bitmap(wid, hgt)
            Dim gr As Graphics = Graphics.FromImage(bm)
            gr.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

            Dim rectf As New RectangleF(0, 0, wid, hgt)
            Dim br As Brush
            br = New HatchBrush(HatchStyle.SmallConfetti, _
                Color.LightGray, Color.White)
            gr.FillRectangle(br, rectf)

            Dim text_size As SizeF
            Dim the_font As Font
            Dim font_size As Single = hgt + 1
            Do
                font_size -= 1
                the_font = New Font(font_family_name, font_size, _
                    FontStyle.Bold, GraphicsUnit.Pixel)
                text_size = gr.MeasureString(txt, the_font)
            Loop While (text_size.Width > wid) OrElse _
                (text_size.Height > hgt)

            ' Center the text.
            Dim string_format As New StringFormat
            string_format.Alignment = StringAlignment.Center
            string_format.LineAlignment = StringAlignment.Center

            ' Convert the text into a path.
            Dim graphics_path As New GraphicsPath
            graphics_path.AddString(txt, the_font.FontFamily, _
                1, the_font.Size, rectf, _
                string_format)

            ' Make random warping parameters.
            Dim rnd As New Random
            Dim pts() As PointF = { _
                New PointF(CSng(rnd.Next(wid) / 4), _
                    CSng(rnd.Next(hgt) / 4)), _
                New PointF(wid - CSng(rnd.Next(wid) / 4), _
                    CSng(rnd.Next(hgt) / 4)), _
                New PointF(CSng(rnd.Next(wid) / 4), hgt - _
                    CSng(rnd.Next(hgt) / 4)), _
                New PointF(wid - CSng(rnd.Next(wid) / 4), hgt - _
                    CSng(rnd.Next(hgt) / 4)) _
            }
            Dim mat As New Matrix
            graphics_path.Warp(pts, rectf, mat, _
                WarpMode.Perspective, 0)

            ' Draw the text.
            br = New HatchBrush(HatchStyle.LargeConfetti, _
                Color.LightGray, Color.DarkGray)
            gr.FillPath(br, graphics_path)

            ' Mess things up a bit.
            Dim max_dimension As Integer = Max(wid, hgt)
            For i As Integer = 0 To CInt(wid * hgt / 30)
                Dim X As Integer = rnd.Next(wid)
                Dim Y As Integer = rnd.Next(hgt)
                Dim W As Integer = CInt(rnd.Next(max_dimension) / _
                    50)
                Dim H As Integer = CInt(rnd.Next(max_dimension) / _
                    50)
                gr.FillEllipse(br, X, Y, W, H)
            Next i
            For i As Integer = 1 To 5
                Dim x1 As Integer = rnd.Next(wid)
                Dim y1 As Integer = rnd.Next(hgt)
                Dim x2 As Integer = rnd.Next(wid)
                Dim y2 As Integer = rnd.Next(hgt)
                gr.DrawLine(Pens.DarkGray, x1, y1, x2, y2)
            Next i
            For i As Integer = 1 To 5
                Dim x1 As Integer = rnd.Next(wid)
                Dim y1 As Integer = rnd.Next(hgt)
                Dim x2 As Integer = rnd.Next(wid)
                Dim y2 As Integer = rnd.Next(hgt)
                gr.DrawLine(Pens.LightGray, x1, y1, x2, y2)
            Next i

            graphics_path.Dispose()
            br.Dispose()
            the_font.Dispose()
            gr.Dispose()

            Return bm
        End Function
    End Class

End Namespace