Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Reflection  
Imports System.Drawing
Imports System.Web
Imports System.Windows.Forms
Imports System.Windows.Forms.Control
Imports System.Data
Namespace WinForms

    ''' <summary>
    ''' Winform Helper class for extending the methods available for common windows form controls
    ''' </summary>
    ''' <remarks></remarks>
    Public Module WinFormHelper
        ''' <summary>
        ''' Loads a user control to specific panel inside a form
        ''' </summary>
        ''' <param name="f">Parent form</param>
        ''' <param name="targetPanel">Target panle</param>
        ''' <param name="sourceControl">User control to be loaded</param>
        ''' <remarks></remarks>
        <Extension()> _
        Public Sub LoadControl(f As Form, ByRef targetPanel As Panel, sourceControl As UserControl)
            Try
                targetPanel.Controls.Clear()
                sourceControl.Width = targetPanel.Width - 1
                sourceControl.Height = targetPanel.Height
                targetPanel.Controls.Add(sourceControl)
                sourceControl.Focus()
            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' Binds datatable as items on a combobox
        ''' </summary>
        ''' <param name="Cmb">Target combobox</param>
        ''' <param name="dts">Source datatable</param>
        ''' <param name="SelValue">Column name for value member</param>
        ''' <param name="DisValue">Column name for display value</param>
        ''' <param name="DisPlayStr">Default string for display e.g Select..</param>
        ''' <param name="IsGuidCol">Optional indicator if value member is of type guid</param>
        ''' <remarks></remarks>
        <Extension()> Public Sub FormatForCombo(ByVal Cmb As ComboBox, ByVal dts As DataTable, _
                                   ByVal SelValue As String, _
                                   ByVal DisValue As String, _
                                   ByVal DisPlayStr As String, Optional ByVal IsGuidCol As Boolean = True)
            Try
                Dim dt As New DataTable
                dt = dts.Clone()

                Dim dtRow As DataRow : dtRow = dt.NewRow
                If IsGuidCol Then
                    dtRow(SelValue) = Guid.NewGuid
                Else
                    dtRow(SelValue) = 0
                End If

                dtRow(DisValue) = DisPlayStr
                dt.Rows.Add(dtRow)
                dt.Merge(dts, True)
                Cmb.DataSource = dt
                Cmb.ValueMember = SelValue
                Cmb.DisplayMember = DisValue
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        ''' <summary>
        ''' Binds collection object to a combo box
        ''' </summary>
        ''' <param name="cmb">Target combobox</param>
        ''' <param name="dt">Source collection object e.g. datatable</param>
        ''' <param name="ValueStr">Name of value member</param>
        ''' <param name="DisplayStr">Name of display member</param>
        ''' <remarks></remarks>
        <Extension()> Public Sub BindCombo(ByVal cmb As ComboBox, ByVal dt As Object, ByVal ValueStr As String, ByVal DisplayStr As String)
            Try
                cmb.DataSource = dt
                cmb.DisplayMember = DisplayStr
                cmb.ValueMember = ValueStr
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        ''' <summary>
        ''' Binds list of object to a combobox
        ''' </summary>
        ''' <param name="cmb">Target combobox</param>
        ''' <param name="Items">Source list for binding</param>
        ''' <remarks></remarks>
        <Extension()> Public Sub BindCombo(ByVal cmb As ComboBox, ByVal Items As List(Of Object))
            Try
                cmb.Items.Clear()
                For i As Integer = 0 To Items.Count - 1
                    cmb.Items.Add(Items(i))
                Next
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' Hide a form with a fade in effect
        ''' </summary>
        ''' <param name="f">target form</param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub FadeHide(ByVal f As Form)
            f.Opacity = 1
            For i As Decimal = 1.0 To 0.0001 Step -0.001
                f.Opacity = i
            Next
            f.Hide()
            f.Opacity = 1
        End Sub

        ''' <summary>
        ''' Show a form with a fade in effect
        ''' </summary>
        ''' <param name="f">Target form</param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub FadeShow(ByVal f As Form)
            f.Opacity = 0.001
            f.Show()
            For i As Decimal = 0.001 To 1 Step 0.001
                f.Opacity = i
            Next
            f.Opacity = 1
        End Sub

        ''' <summary>
        ''' Check of an object contains no content or valid item selection has not been performed
        ''' </summary>
        ''' <param name="CtrlObj">Target control object which include Textbox, Picturebox, Combobox</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function IsEmpty(ByVal CtrlObj As Object) As Boolean
            Try
                If CtrlObj.GetType Is GetType(TextBox) Then 'textbox
                    If CType(CtrlObj, TextBox).Text.Trim = "" Then
                        Return True
                    Else
                        Return False
                    End If
                ElseIf CtrlObj.GetType Is GetType(ComboBox) Then 'combobox
                    Dim cmb As ComboBox = CType(CtrlObj, ComboBox)
                    If cmb.SelectedItem Is Nothing OrElse cmb.Text.ToString.Trim = "" Or cmb.Text.ToString.Trim.ToLower.Contains("select-") Or cmb.Text.ToString.Trim.ToLower.Contains("select.") Then
                        Return True
                    Else
                        Return False
                    End If
                ElseIf CtrlObj.GetType Is GetType(PictureBox) Then 'combobox
                    If CType(CtrlObj, PictureBox).Image Is Nothing Then
                        Return True
                    Else
                        Return False
                    End If

                Else
                    Throw New Exception("Valid Target Not Found")
                End If
            Catch ex As Exception
                Throw ex
                Return True
            End Try
        End Function

        ''' <summary>
        ''' Gets desired value ID from displayed items on a checklistbox by splitting items
        ''' </summary>
        ''' <param name="chk">Target checklistbox</param>
        ''' <param name="ItmSeparator">The char for splitting each item</param>
        ''' <param name="Separator">The char for combining the values together as string</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function GetIDAsString(ByVal chk As CheckedListBox, Optional ByVal ItmSeparator As String = "-", Optional ByVal Separator As String = "|") As String
            Try
                Dim Val As String = ""
                For i As Integer = 0 To chk.Items.Count - 1
                    If chk.GetItemChecked(i) Then
                        Val &= Separator & Trim(chk.Items(i).ToString.Split(ItmSeparator)(chk.Items(i).ToString.Split(ItmSeparator).Length - 1))
                    End If
                Next
                If Val.Trim.Length > 1 Then Return Val.Substring(1) Else If Val.Trim.Trim.Length = 1 Then Return Val
            Catch ex As Exception
                Throw ex
            End Try
            Return ""
        End Function

        ''' <summary>
        ''' Check if at least an item is selected on a checkedlistbox
        ''' </summary>
        ''' <param name="chk">Target control</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ItemSelected(ByVal chk As CheckedListBox) As Boolean
            Try
                For i As Integer = 0 To chk.Items.Count - 1
                    If chk.GetItemChecked(i) Then
                        Return True
                    End If
                Next
                Return False
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Check if checkedlistbox contains a an item
        ''' </summary>
        ''' <param name="chk">Target control</param>
        ''' <param name="ItemName">Item to be checked</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function HasItem(ByVal chk As CheckedListBox, ByVal ItemName As String) As Boolean
            Try
                For i As Integer = 0 To chk.Items.Count - 1
                    If chk.Items(i).ToString.ToLower = ItemName.ToLower Then
                        Return True
                    End If
                Next
                Return False
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Binds a datatable to checkedlistbox
        ''' </summary>
        ''' <param name="ChkList">Target control</param>
        ''' <param name="Display">The column name of the datatable to be used</param>
        ''' <param name="dt">Source datatable</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function BindChkList(ByRef ChkList As CheckedListBox, _
                             ByVal Display As String, _
                             ByRef dt As DataTable) As Boolean
            Try
                ChkList.Items.Clear()
                For Each rw As DataRow In dt.Rows
                    ChkList.Items.Add(rw.Item(Display), False)
                Next
                Return True
            Catch ex As Exception
                Throw ex
            End Try
            Return False
        End Function

        ''' <summary>
        ''' Toggle selection state of the checkboxes in checkedListbox
        ''' </summary>
        ''' <param name="ChkList">Target control</param>
        ''' <param name="Value">Boolean value to be passed for selection</param>
        ''' <remarks></remarks>
        <Extension()> Public Sub ToggleSelect(ByRef ChkList As CheckedListBox, Optional ByVal Value As Boolean = True)
            For i As Integer = 0 To ChkList.Items.Count - 1
                ChkList.SetItemChecked(i, Value)
            Next
        End Sub

        ''' <summary>
        ''' Binds dataview to a checkedlistbox
        ''' </summary>
        ''' <param name="ChkList">Target Control</param>
        ''' <param name="dv">Source dataview</param>
        ''' <param name="Display">Required display name from the dataview</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function BindChkList(ByRef ChkList As CheckedListBox, _
                             ByRef dv As DataView, _
                             ByVal Display As String) As Boolean
            Try
                ChkList.Items.Clear()
                Dim dt As DataTable = CreateTableFromDataview(dv)
                If dt IsNot Nothing Then
                    For Each rw As DataRow In dt.Rows
                        ChkList.Items.Add(rw.Item(Display), False)
                    Next
                End If

                Return True
            Catch ex As Exception
                Throw ex
            End Try
            Return False
        End Function

        ''' <summary>
        ''' Gets the number of selected items in a checkedlistbox
        ''' </summary>
        ''' <param name="chk">Target control</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function GetSelectedItemCount(ByVal chk As CheckedListBox) As Integer
            Try
                Dim j As Integer = 0
                For i As Integer = 0 To chk.Items.Count - 1
                    If chk.GetItemChecked(i) Then
                        j += 1
                    End If
                Next
                Return j
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Converts a dataview to a datatable
        ''' </summary>
        ''' <param name="obDataView">Target control</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function CreateTableFromDataview(ByVal obDataView As DataView) As DataTable

            If obDataView Is Nothing Then
                Return Nothing
            End If

            Dim obNewDt As DataTable = obDataView.Table.Clone()
            Dim idx As Integer = 0
            Dim strColNames As String() = New String(obNewDt.Columns.Count - 1) {}

            For Each col As DataColumn In obNewDt.Columns

                'strColNames(System.Math.Max(System.Threading.Interlocked.Increment(idx), idx - 1)) = col.ColumnName
                strColNames(idx) = col.ColumnName
                idx += 1
            Next


            Dim viewEnumerator As IEnumerator = obDataView.GetEnumerator()
            While viewEnumerator.MoveNext()
                Dim drv As DataRowView = DirectCast(viewEnumerator.Current, DataRowView)
                Dim dr As DataRow = obNewDt.NewRow()
                Try
                    For Each strName As String In strColNames
                        dr(strName) = drv(strName)
                    Next
                Catch ex As Exception
                    Trace.WriteLine(ex.Message)
                End Try
                obNewDt.Rows.Add(dr)
            End While


            Return obNewDt
        End Function

        ''' <summary>
        ''' Binds a datatable to a listbox
        ''' </summary>
        ''' <param name="List">Target control</param>
        ''' <param name="Display">Display member required</param>
        ''' <param name="dt">Source datatable</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function BindList(ByRef List As ListBox, _
                                 ByVal Display As String, _
                                 ByRef dt As DataTable) As Boolean
            Try
                List.Items.Clear()
                For Each rw As DataRow In dt.Rows
                    List.Items.Add(rw.Item(Display))
                Next
                Return True
            Catch ex As Exception
                Throw ex
            End Try
            Return False
        End Function

        ''' <summary>
        ''' Binds a dataview to a listbox
        ''' </summary>
        ''' <param name="List">Target control</param>
        ''' <param name="Display">Required display name from dataview</param>
        ''' <param name="dv">Source dataview</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function BindList(ByRef List As ListBox, _
                                 ByVal Display As String, _
                                 ByRef dv As DataView) As Boolean
            Try
                List.Items.Clear()
                Dim dt As DataTable = CreateTableFromDataview(dv)
                If dt IsNot Nothing Then
                    For Each rw As DataRow In dt.Rows
                        List.Items.Add(rw.Item(Display))
                    Next
                End If

            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Binds a datatable to a listbox
        ''' </summary>
        ''' <param name="LstBox">Target control</param>
        ''' <param name="Value">Value member to be used</param>
        ''' <param name="Display">Display member used</param>
        ''' <param name="dt">Source datatable</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function BindList(ByRef LstBox As ListBox, _
                             ByVal Value As String, _
                             ByVal Display As String, _
                             ByRef dt As DataTable) As Boolean
            Try
                If dt.Rows.Count > 0 Then
                    LstBox.DataSource = dt
                    LstBox.DisplayMember = Display
                    LstBox.ValueMember = Value
                    Return True
                End If
            Catch ex As Exception
            End Try
            Return False
        End Function

        ''' <summary>
        ''' Removes all the columns of DataGridView
        ''' </summary>
        ''' <param name="dgvList">Target control</param>
        ''' <remarks></remarks>
        <Extension()> Public Sub ClearGrid(ByRef dgvList As DataGridView)
            dgvList.DataSource = Nothing
            'remove all columns
            For i As Integer = dgvList.ColumnCount - 1 To 0 Step -1
                dgvList.Columns.RemoveAt(i)
            Next
        End Sub

        ''' <summary>
        ''' Hides all the columns in datagridview
        ''' </summary>
        ''' <param name="dgvList">Target control</param>
        ''' <remarks></remarks>
        <Extension()> Public Sub HideColumn(ByRef dgvList As DataGridView)
            'remove all columns
            For i As Integer = dgvList.ColumnCount - 1 To 0 Step -1
                dgvList.Columns(i).Visible = False
            Next
        End Sub

        ''' <summary>
        ''' Display MsgBox with application title as the default title
        ''' </summary>
        ''' <param name="Prompt">Content of msgbox</param>
        ''' <param name="MsgBoxStyles">Optional msgboxstyles required. Default is OkOnly</param>
        ''' <remarks></remarks>
        Public Sub TMsgBox(ByVal Prompt As String, Optional ByVal MsgBoxStyles As Microsoft.VisualBasic.MsgBoxStyle = MsgBoxStyle.OkOnly)
            MsgBox(Prompt, MsgBoxStyles, My.Application.Info.Title)
        End Sub
    End Module
End Namespace
