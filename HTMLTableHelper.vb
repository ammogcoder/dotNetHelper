Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Reflection
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls

Namespace HTML

    ''' <summary>
    ''' A Class for server manipulation of html table to provide easy creation and output as string
    ''' </summary>
    ''' <remarks></remarks>
    Public Class HTMLTableHelper
        Inherits System.Web.UI.Page

        'Table String Variable
        Dim RetTable As String = ""
        ''' <summary>
        ''' Instantiate a new instance of the HTML Table with attributes
        ''' </summary>
        ''' <param name="TWidth"></param>
        ''' <param name="TAlign"></param>
        ''' <param name="TBorder"></param>
        ''' <param name="TCellPadding"></param>
        ''' <param name="TCellSpacing"></param>
        ''' <param name="TTitle"></param>
        ''' <param name="TBgColor"></param>
        ''' <param name="TClass"></param>
        ''' <param name="TStyle"></param>
        ''' <param name="TID"></param>
        ''' <remarks></remarks>
        Sub New(Optional ByVal TWidth As String = "", Optional ByVal TAlign As String = "", _
                Optional ByVal TBorder As String = "", Optional ByVal TCellPadding As String = "", _
                Optional ByVal TCellSpacing As String = "", Optional ByVal TTitle As String = "", _
                Optional ByVal TBgColor As String = "", Optional ByVal TClass As String = "", _
                Optional ByVal TStyle As String = "", _
                Optional ByVal TID As String = "")

            RetTable = "<table "
            If TWidth <> "" Then
                RetTable &= " Width=" & TWidth
            End If
            If TAlign <> "" Then
                RetTable &= " Align=" & TAlign
            End If
            If TBorder <> "" Then
                RetTable &= " Border=" & TBorder
            End If
            If TCellPadding <> "" Then
                RetTable &= " CellPadding=" & TCellPadding
            End If
            If TCellSpacing <> "" Then
                RetTable &= " CellSpacing=" & TCellSpacing
            End If
            If TTitle <> "" Then
                RetTable &= " Title=" & TTitle
            End If
            If TBgColor <> "" Then
                RetTable &= " BgColor=" & TBgColor
            End If
            If TClass <> "" Then
                RetTable &= " Class=" & TClass
            End If
            If TStyle <> "" Then
                RetTable &= " Style=" & TStyle
            End If
            If TID <> "" Then
                RetTable &= " ID='" & TID & "'"
            End If
            RetTable &= ">"
        End Sub

        ''' <summary>
        ''' Add special tag to the derived output
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <remarks></remarks>
        Public Sub AddSpecial(ByVal Str As String)
            RetTable &= Str
        End Sub

        ''' <summary>
        ''' Add a table row to the table
        ''' </summary>
        ''' <param name="RWidth"></param>
        ''' <param name="RTitle"></param>
        ''' <param name="RBgColor"></param>
        ''' <param name="RClass"></param>
        ''' <param name="RStyle"></param>
        ''' <param name="RValign"></param>
        ''' <param name="ROnClickAttr"></param>
        ''' <remarks></remarks>
        Public Sub AddRow(Optional ByVal RWidth As String = "", Optional ByVal RTitle As String = "", _
                          Optional ByVal RBgColor As String = "", Optional ByVal RClass As String = "", _
                          Optional ByVal RStyle As String = "", Optional ByVal RValign As String = "", _
                      Optional ByVal ROnClickAttr As String = "", Optional ByVal RInlineAttr As String = "")
            'check if a row is still opened
            If RetTable.ToLower.EndsWith("</td>") Or RetTable.ToLower.EndsWith("</th>") Then
                RetTable &= "</tr>"
            End If
            'add the new row
            RetTable &= "<tr"
            If RWidth <> "" Then
                RetTable &= " Width=" & RWidth
            End If
            If RTitle <> "" Then
                RetTable &= " Title=" & RTitle
            End If
            If RBgColor <> "" Then
                RetTable &= " BgColor=" & RBgColor
            End If
            If RClass <> "" Then
                RetTable &= " Class=" & RClass
            End If
            If RStyle <> "" Then
                RetTable &= " Style=" & RStyle
            End If
            If RValign <> "" Then
                RetTable &= " Valign=" & RValign
            End If
            If ROnClickAttr <> "" Then
                RetTable &= " style='pointer:cursor' onclick='" & ROnClickAttr & "'"
            End If
            If RInlineAttr <> "" Then
                RetTable &= " " & RInlineAttr & " "
            End If
            RetTable &= ">"
        End Sub

        ''' <summary>
        ''' Add a cell to the current row in the table
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <param name="AsHeader"></param>
        ''' <param name="CWidth"></param>
        ''' <param name="CHeight"></param>
        ''' <param name="CAlign"></param>
        ''' <param name="CTitle"></param>
        ''' <param name="CBgColor"></param>
        ''' <param name="CClass"></param>
        ''' <param name="CStyle"></param>
        ''' <param name="CValign"></param>
        ''' <param name="CRowSpan"></param>
        ''' <param name="CColSpan"></param>
        ''' <param name="COnClickAttr"></param>
        ''' <remarks></remarks>
        Public Sub AddCell(ByVal Value As String, Optional ByVal AsHeader As Boolean = False, _
                           Optional ByVal CWidth As String = "", _
                           Optional ByVal CHeight As String = "", _
                           Optional ByVal CAlign As String = "", Optional ByVal CTitle As String = "", _
                           Optional ByVal CBgColor As String = "", _
                          Optional ByVal CClass As String = "", Optional ByVal CStyle As String = "", _
                          Optional ByVal CValign As String = "", _
                           Optional ByVal CRowSpan As String = "", _
                           Optional ByVal CColSpan As String = "", Optional ByVal COnClickAttr As String = "")

            If AsHeader Then
                RetTable &= "<th "
            Else
                RetTable &= "<td "
            End If

            If CWidth <> "" Then
                RetTable &= " Width=" & CWidth
            End If
            If CHeight <> "" Then
                RetTable &= " Height=" & CHeight
            End If
            If CAlign <> "" Then
                RetTable &= " Align=" & CAlign
            End If
            If CValign <> "" Then
                RetTable &= " VAlign=" & CValign
            End If
            If CTitle <> "" Then
                RetTable &= " Title=" & CTitle
            End If
            If CRowSpan <> "" Then
                RetTable &= " RowSpan=" & CRowSpan
            End If
            If CColSpan <> "" Then
                RetTable &= " ColSpan=" & CColSpan
            End If
            If CBgColor <> "" Then
                RetTable &= " BgColor=" & CBgColor
            End If
            If CClass <> "" Then
                RetTable &= " Class=" & CClass
            End If
            If CStyle <> "" Then
                RetTable &= " Style=" & CStyle
            End If
            If COnClickAttr <> "" Then
                RetTable &= "  onclick=" & COnClickAttr
            End If
            RetTable &= ">" & Value

            If AsHeader Then
                RetTable &= "</th>"
            Else
                RetTable &= "</td>"
            End If

        End Sub

        ''' <summary>
        ''' Replace all HTML space in a string with HTML space key string
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function HTMLEncodeForSpace(ByVal Value As String) As String
            Return Value.Replace(" ", "&nbsp;")
        End Function

        ''' <summary>
        ''' Perform HTML encode on a string
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function HTMLEncode(ByVal Value As String) As String
            Return Server.HtmlEncode(Value)
        End Function

        ''' <summary>
        ''' Return the derived table as string
        ''' </summary>
        ''' <param name="SpecialEnd"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTable(Optional ByVal SpecialEnd As String = "") As String
            If Not RetTable.ToLower.EndsWith("</tr>") And SpecialEnd = "" Then
                RetTable &= "</tr></table>"
            ElseIf SpecialEnd <> "" Then
                RetTable &= "</tr>" & SpecialEnd & "</table>"
            Else
                RetTable &= "</table>"
            End If
            'Return "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'><html xmlns=http://www.w3.org/1999/xhtml><head></head><body>" & RetTable & "<table><tr><td><a onclick=popupDiv('n'); href=>jku</a></td></tr></table></body></html>"
            Return RetTable '& "<table>"
        End Function

    End Class

End Namespace
