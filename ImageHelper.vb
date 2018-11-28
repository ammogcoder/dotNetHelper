Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Reflection
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing

Namespace Images
    ''' <summary>
    ''' Image Helper class for conversion and manipulation of image objects
    ''' </summary>
    ''' <remarks></remarks>
    Public Module ImageHelper
        ''' <summary>
        ''' Converts a picture file to a bitmap
        ''' </summary>
        ''' <param name="FileName">Name of the file to be converted</param>
        ''' <param name="Imgsize">Optional dimension for resizing the bitmap if default is not required</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ConvertFileToBitmap(ByVal FileName As String, Optional Imgsize As Size = Nothing) As Bitmap
            Dim myPhotoStream As New FileStream(FileName, FileMode.Open, FileAccess.Read)
            Dim myPhotoBuffer(myPhotoStream.Length) As Byte
            myPhotoStream.Read(myPhotoBuffer, 0, Convert.ToInt32(myPhotoStream.Length))
            Dim myPhoto As Bitmap = New Bitmap(myPhotoStream)
            Dim PhotoToDisplay As Bitmap
            If Imgsize = Nothing Then
                PhotoToDisplay = New Bitmap(myPhoto)
            Else
                PhotoToDisplay = New Bitmap(myPhoto, Imgsize.Width, Imgsize.Height)
            End If
            myPhotoStream.Dispose()
            Return PhotoToDisplay
        End Function

        ''' <summary>
        ''' Writes image binary to file
        ''' </summary>
        ''' <param name="ImagePath">Full file name for image</param>
        ''' <param name="Images">Image binary</param>
        ''' <param name="ImgWidth">The width to be used for image</param>
        ''' <param name="imgHeight">The height to be used for image</param>
        ''' <param name="ImageFormats">The image format required</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function WriteBinarytoFile(ByVal ImagePath As String, ByVal Images As Byte(),
                                          ImgWidth As Int16, imgHeight As Int16,
                                            ByVal ImageFormats As Imaging.ImageFormat) As Boolean
            Try
                Dim PhotoByte() As Byte = CType(Images, Byte())
                Dim PhotoStream As New MemoryStream(PhotoByte, True)
                PhotoStream.Write(PhotoByte, 0, PhotoByte.Length)
                Dim im As Image = Image.FromStream(PhotoStream)
                Dim dummyCallBack As System.Drawing.Image.GetThumbnailImageAbort
                dummyCallBack = New System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
                Dim ReducedImage As Drawing.Image
                ReducedImage = im.GetThumbnailImage(ImgWidth, imgHeight, dummyCallBack, IntPtr.Zero)
                PhotoStream.Dispose()
                ReducedImage.Save(ImagePath, ImageFormats)
                Return True
            Catch ex As Exception
                Throw ex
            End Try
            Return False
        End Function

        Private Function ThumbnailCallback() As Boolean
            Return False
        End Function

        ''' <summary>
        ''' Converts image binary to bitmap
        ''' </summary>
        ''' <param name="Photo">Array of byte to be converted</param>
        ''' <param name="BitmapSize">Optional Size of the bitmap required</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ConvertByteToBitmap(ByVal Photo() As Byte, Optional BitmapSize As Size = Nothing) As Bitmap
            Try
                Dim StaffPhotoStream As New MemoryStream(Photo, True)
                StaffPhotoStream.Write(Photo, 0, Photo.Length)
                Dim CandidatesImage As New Bitmap(StaffPhotoStream)
                If BitmapSize <> Nothing Then CandidatesImage = New Bitmap(CandidatesImage, BitmapSize)
                Return CandidatesImage
            Catch ex As Exception
                Throw ex
            End Try
        End Function

    End Module

End Namespace