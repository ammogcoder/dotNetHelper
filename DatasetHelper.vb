Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.Reflection
Imports System.Data
Imports System.Data.Linq
Imports System.Data.SqlClient
Imports System.Linq

Namespace Datasets
    ''' <summary>
    ''' Data Helper class for extending the methods for dataset and datatables
    ''' </summary>
    ''' <remarks></remarks>
    Public Module DatasetHelper
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function HasRow(source As DataTable) As Boolean
            Try
                Return source IsNot Nothing AndAlso source.Rows.Count > 0
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="Column"></param>
        ''' <param name="Separator"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToStringConcatenate(ByVal dt As DataTable, Column As Object, Optional ByVal Separator As String = " ") As String
            Dim out As String = ""
            For Each rw As DataRow In dt.Rows
                out &= Separator & rw(Column).ToString
            Next
            If out.Trim.Length > 1 Then Return out.Substring(1) Else If out.Trim.Trim.Length = 1 Then Return out
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <param name="row"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function DatarowsToDatatable(source As DataTable, ByVal row() As DataRow) As DataTable
            Try
                Dim dt As New System.Data.DataTable
                dt = source.Clone

                For Each drow In row
                    dt.ImportRow(drow)
                Next
                Return dt
            Catch ex As Exception
                Return Nothing
            End Try
        End Function
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <param name="Value"></param>
        ''' <param name="ColumnIndex"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function HasItem(source As DataTable, ByVal Value As String, Optional ByVal ColumnIndex As Object = Nothing) As Boolean
            Try
                If ColumnIndex = Nothing Then ColumnIndex = 0 'set to first column
                For Each drow As DataRow In source.Rows
                    If drow.Item(ColumnIndex).ToString.Trim.ToLower = Value.Trim.ToLower Then Return True
                Next
                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="source"></param>
        ''' <param name="Value"></param>
        ''' <param name="ColumnIndex"></param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub RemoveSpecificRow(source As DataTable, ByVal Value As String, Optional ByVal ColumnIndex As Int16 = 0)
            Try
                For i = source.Rows.Count - 1 To 0 Step -1
                    If source.Rows(i).Item(ColumnIndex).ToString.Trim.ToLower = Value.Trim.ToLower Then source.Rows.RemoveAt(i)
                Next
            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="ds"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function GetFirstRowInFirstTable(ByVal ds As DataSet) As DataRow
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
                Return ds.Tables(0).Rows(0)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="ds"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function GetFirstTable(ByVal ds As DataSet) As DataTable
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                Return ds.Tables(0)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="i_dSourceTable"></param>
        ''' <param name="i_sGroupByColumn"></param>
        ''' <param name="i_sAggregateColumn"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function GroupBy(ByVal i_dSourceTable As DataTable, ByVal i_sGroupByColumn As String, ByVal i_sAggregateColumn As String) As DataTable

            Dim dv As New DataView(i_dSourceTable)

            'getting distinct values for group column
            Dim dtGroup As DataTable = dv.ToTable(True, New String() {i_sGroupByColumn})

            'adding column for the row count
            dtGroup.Columns.Add("Count", GetType(Integer))

            'looping thru distinct values for the group, counting
            For Each dr As DataRow In dtGroup.Rows
                dr("Count") = i_dSourceTable.Compute("Count(" & i_sAggregateColumn & ")", _
                                                     i_sGroupByColumn & " = '" & dr(i_sGroupByColumn) & "'")
            Next

            'returning grouped/counted result
            Return dtGroup
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="i_dSourceTable"></param>
        ''' <param name="i_sAggregateColumn"></param>
        ''' <param name="i_sGroupByColumn"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()> Public Function GroupBySum(ByVal i_dSourceTable As DataTable, ByVal i_sAggregateColumn As String, _
                             ByVal i_sGroupByColumn As String) As DataTable

            Dim dv As New DataView(i_dSourceTable)

            'getting distinct values for group column
            Dim dtGroup As DataTable = dv.ToTable(True, New String() {i_sGroupByColumn})

            'adding column for the row count
            dtGroup.Columns.Add("Count", GetType(Integer))

            'looping thru distinct values for the group, counting
            For Each dr As DataRow In dtGroup.Rows
                dr("Count") = i_dSourceTable.Compute("Count(" & i_sAggregateColumn & ")", _
                                                     i_sGroupByColumn & " = '" & dr(i_sGroupByColumn) & "'")
            Next

            'returning grouped/counted result
            Return dtGroup
        End Function
       
        ''' <summary>
        ''' Converts a datatable to IEnumerable of T by providing a converter from a delegate routine
        ''' </summary>
        ''' <typeparam name="T">Destination type of model </typeparam>
        ''' <param name="dt">Source datatable</param>
        ''' <param name="converter">Converter routine which must run as a delegate</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Extension()>
        Public Function ToEnumerableModel(Of T)(dt As DataTable, converter As Converter(Of DataRow, T)) As IEnumerable(Of T)
            Return (From row In dt.AsEnumerable
                    Select converter(row)).ToList()
        End Function

#Region "Table Handler"
        ''' <summary>
        ''' Perform distinct select on a datatavble
        ''' </summary>
        ''' <param name="SourceTable"></param>
        ''' <param name="FieldNames"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SelectDistinct(ByVal SourceTable As DataTable, ByVal ParamArray FieldNames() As String) As DataTable
            Dim lastValues() As Object
            Dim newTable As DataTable

            If FieldNames Is Nothing OrElse FieldNames.Length = 0 Then
                Throw New ArgumentNullException("FieldNames")
            End If

            lastValues = New Object(FieldNames.Length - 1) {}
            newTable = New DataTable

            For Each field As String In FieldNames
                newTable.Columns.Add(field, SourceTable.Columns(field).DataType)
            Next

            For Each Row As DataRow In SourceTable.Select("", String.Join(", ", FieldNames))
                If Not fieldValuesAreEqual(lastValues, Row, FieldNames) Then
                    newTable.Rows.Add(createRowClone(Row, newTable.NewRow(), FieldNames))

                    setLastValues(lastValues, Row, FieldNames)
                End If
            Next
            Return newTable
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="lastValues"></param>
        ''' <param name="currentRow"></param>
        ''' <param name="fieldNames"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function fieldValuesAreEqual(ByVal lastValues() As Object, ByVal currentRow As DataRow, ByVal fieldNames() As String) As Boolean
            Dim areEqual As Boolean = True

            For i As Integer = 0 To fieldNames.Length - 1
                If lastValues(i) Is Nothing OrElse Not lastValues(i).Equals(currentRow(fieldNames(i))) Then
                    areEqual = False
                    Exit For
                End If
            Next

            Return areEqual
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="newRow"></param>
        ''' <param name="fieldNames"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function createRowClone(ByVal sourceRow As DataRow, ByVal newRow As DataRow, ByVal fieldNames() As String) As DataRow
            For Each field As String In fieldNames
                newRow(field) = sourceRow(field)
            Next

            Return newRow
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="lastValues"></param>
        ''' <param name="sourceRow"></param>
        ''' <param name="fieldNames"></param>
        ''' <remarks></remarks>
        Private Sub setLastValues(ByVal lastValues() As Object, ByVal sourceRow As DataRow, ByVal fieldNames() As String)
            For i As Integer = 0 To fieldNames.Length - 1
                lastValues(i) = sourceRow(fieldNames(i))
            Next
        End Sub
#End Region

#Region "Merge, Update Dataset For Report"
        ''' <summary>
        ''' Merge the content of two tables
        ''' </summary>
        ''' <param name="dt1"></param>
        ''' <param name="dt2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MergeDatatable(ByVal dt1 As DataTable, ByVal dt2 As DataTable) As DataSet
            Try
                Dim ds As New DataSet
                ds.Tables.Add(dt1.Copy)
                ds.Tables.Add(dt2.Copy)

                Return ds
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Add a table to a dataset
        ''' </summary>
        ''' <param name="ds"></param>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AddDatatableToDataset(ByVal ds As DataSet, ByVal dt As DataTable) As DataSet
            Try
                ds.Tables.Add(dt.Copy)
                Return ds
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Name all the tables in a dataset 
        ''' </summary>
        ''' <param name="arrTablesName"></param>
        ''' <param name="Dobj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function RenameDatatables(ByVal arrTablesName() As String, ByVal Dobj As Object) As Object
            Dim Dtype = Dobj.GetType

            Try
                If Dtype Is GetType(DataSet) Then

                    'check to see the table count matches array count
                    If arrTablesName.Length <> CType(Dobj, DataSet).Tables.Count Then
                        Throw New Exception("Table Count Must Match The Array Count")
                    End If

                    For i As Integer = 0 To arrTablesName.Length - 1
                        CType(Dobj, DataSet).Tables(i).TableName = arrTablesName(i)
                    Next

                ElseIf Dtype Is GetType(DataTable) Then
                    'check to see the table count matches array count
                    If arrTablesName.Length > 1 Then
                        Throw New Exception("Only One Name Required For Table")
                    End If

                    CType(Dobj, DataTable).TableName = arrTablesName(0)

                Else
                    Throw New Exception("Only Set or Table Is Allowed")
                End If

                Return Dobj
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' Name all the tables in a dataset and write them to file with option
        ''' </summary>
        ''' <param name="arrTablesName"></param>
        ''' <param name="Dobj"></param>
        ''' <param name="XmlName"></param>
        ''' <param name="WriteXmlToFile"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function RenameDatatables(ByVal arrTablesName() As String, ByVal Dobj As Object, ByVal XmlName As String, ByVal WriteXmlToFile As Boolean) As Object
            Try
                Dim RenameObj = RenameDatatables(arrTablesName, Dobj)
                If WriteXmlToFile Then
                    RenameObj.WriteXmlSchema(XmlName & ".xsd")
                End If
                Return RenameObj

            Catch ex As Exception
                Throw ex
            End Try

        End Function

        ''' <summary>
        ''' Name all the tables in a dataset and write them to file
        ''' </summary>
        ''' <param name="arrTablesName"></param>
        ''' <param name="Dobj"></param>
        ''' <param name="XmlDataName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function RenameDatatables(ByVal arrTablesName() As String, ByVal Dobj As Object, ByVal XmlDataName As String) As Object
            Try
                Dim RenameObj = RenameDatatables(arrTablesName, Dobj)
                RenameObj.WriteXml(XmlDataName & ".xsd")
                Return RenameObj

            Catch ex As Exception
                Throw ex
            End Try

        End Function
        ''' <summary>
        ''' Name a dataset and write the schema to file
        ''' </summary>
        ''' <param name="Ds"></param>
        ''' <param name="DatasetName"></param>
        ''' <param name="XmlSchemaName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function RenameDatasetWriteSchema(ByVal Ds As DataSet, ByVal DatasetName As String, ByVal XmlSchemaName As String) As DataSet
            Try
                Ds.DataSetName = DatasetName
                Ds.WriteXmlSchema(XmlSchemaName & ".xsd")
                Return Ds

            Catch ex As Exception
                Throw ex
            End Try

        End Function
        ''' <summary>
        ''' Writes a dataset schema to a file
        ''' </summary>
        ''' <param name="Ds"></param>
        ''' <param name="XmlSchemaName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function WriteXmlSchema(ByVal Ds As DataSet, ByVal XmlSchemaName As String) As Boolean
            Try
                Ds.WriteXmlSchema(XmlSchemaName & ".xsd")
                Return True

            Catch ex As Exception
                Throw ex
            End Try

        End Function
#End Region

        ''' <summary>
        ''' 'Search entries of a table based on some keywords
        ''' </summary>
        ''' <param name="DtList"></param>
        ''' <param name="txtstr"></param>
        ''' <param name="searchcol"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SearchExactFromDatatable(ByVal DtList As DataTable, ByVal txtstr As String, ByVal searchcol As String) As DataTable
            Try
                Dim filter As String = searchcol & " ='" & txtstr & "'"
                Dim rows() As DataRow
                rows = DtList.Select(filter)
                Dim dtlocal As New DataTable
                dtlocal = DtList.Clone
                For Each rw As DataRow In rows
                    dtlocal.ImportRow(rw)
                Next
                Return dtlocal

            Catch ex As Exception

            End Try
        End Function

        ''' <summary>
        ''' Perform column summation on a table
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="ColumnName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ComputeColumn(ByVal dt As DataTable, ByVal ColumnName As String) As Decimal
            Try
                Dim ComputeVal As Decimal
                For Each rw As DataRow In dt.Rows
                    ComputeVal += rw.Item(ColumnName)
                Next
                Return ComputeVal
            Catch ex As Exception

            End Try
        End Function

        ''' <summary>
        ''' Search entries of a table based on LIKE filter
        ''' </summary>
        ''' <param name="DtList"></param>
        ''' <param name="txtstr"></param>
        ''' <param name="searchcol"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SearchLikeFromDatatable(ByVal DtList As DataTable, ByVal txtstr As String, ByVal searchcol As String) As DataTable
            Try
                Dim filter As String = searchcol & " like '%" & txtstr & "%'"
                Dim rows() As DataRow
                rows = DtList.Select(filter)
                Dim dtlocal As New DataTable
                dtlocal = DtList.Clone
                For Each rw As DataRow In rows
                    dtlocal.ImportRow(rw)
                Next
                Return dtlocal

            Catch ex As Exception

            End Try
        End Function

        ''' <summary>
        ''' Search entries of a table based on some filter
        ''' </summary>
        ''' <param name="DtList"></param>
        ''' <param name="Filter"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SearchByFilterFromDatatable(ByVal DtList As DataTable, ByVal Filter As String) As DataTable
            Try
                Dim rows() As DataRow
                rows = DtList.Select(Filter)
                Dim dtlocal As New DataTable
                dtlocal = DtList.Clone
                For Each rw As DataRow In rows
                    dtlocal.ImportRow(rw)
                Next
                Return dtlocal

            Catch ex As Exception

            End Try
        End Function

    End Module

End Namespace
