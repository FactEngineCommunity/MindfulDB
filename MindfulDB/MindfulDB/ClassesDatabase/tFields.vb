Imports ADODB
Imports System.Runtime.InteropServices


Namespace Database

    Public Class Fields
        Implements ADODB.Fields

        Private _listFields As New List(Of ADODB.Field)()

        Public ReadOnly Property Count() As Integer Implements ADODB.Fields.Count
            Get
                Return _listFields.Count
            End Get
        End Property

        Public Sub Refresh() Implements ADODB.Fields.Refresh
            Throw New NotImplementedException()
        End Sub

        Default Public ReadOnly Property Item(Index As Object) As ADODB.Field Implements ADODB.Fields.Item
            Get
                Dim i As Integer = Convert.ToInt32(Index)
                If i < 0 OrElse i >= _listFields.Count Then
                    Throw New IndexOutOfRangeException()
                End If
                Return _listFields(i)
            End Get
        End Property

        Private ReadOnly Property Fields20_Item(Index As Object) As Field Implements Fields20.Item
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Public Function _NewEnum() As IEnumerator
            Return _listFields.GetEnumerator()
        End Function

        Public Sub Append(asField As ADODB.Field)
            _listFields.Add(asField)
        End Sub

        Public Sub Delete(Index As Object) Implements ADODB.Fields.Delete
            Dim i As Integer = Convert.ToInt32(Index)
            If i < 0 OrElse i >= _listFields.Count Then
                Throw New IndexOutOfRangeException()
            End If
            _listFields.RemoveAt(i)
        End Sub

        Public Function ItemByName(FieldName As String) As ADODB.Field
            Dim field As ADODB.Field = Nothing
            For Each f As ADODB.Field In _listFields
                If f.Name = FieldName Then
                    field = f
                    Exit For
                End If
            Next
            If field Is Nothing Then
                Throw New COMException("Field not found")
            End If
            Return field
        End Function

        Private Function Fields20_GetEnumerator() As IEnumerator Implements Fields20.GetEnumerator
            Throw New NotImplementedException()
        End Function

        Private Sub Fields20_Refresh() Implements Fields20.Refresh
            Throw New NotImplementedException()
        End Sub

        Private Sub Fields20__Append(Name As String, Type As DataTypeEnum, Optional DefinedSize As Integer = 0, Optional Attrib As FieldAttributeEnum = FieldAttributeEnum.adFldUnspecified) Implements Fields20._Append
            Throw New NotImplementedException()
        End Sub

        Private Sub Fields20_Delete(Index As Object) Implements Fields20.Delete
            Throw New NotImplementedException()
        End Sub

        Private ReadOnly Property Fields20_Count As Integer Implements Fields20.Count
            Get
                Throw New NotImplementedException()
            End Get
        End Property



        Private Function Fields15_GetEnumerator() As IEnumerator Implements Fields15.GetEnumerator
            Throw New NotImplementedException()
        End Function

        Private Sub Fields15_Refresh() Implements Fields15.Refresh
            Throw New NotImplementedException()
        End Sub

        Private ReadOnly Property Fields15_Count As Integer Implements Fields15.Count
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private ReadOnly Property Fields15_Item(Index As Object) As Field Implements Fields15.Item
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private Function _Collection_GetEnumerator() As IEnumerator Implements _Collection.GetEnumerator
            Throw New NotImplementedException()
        End Function

        Private Sub _Collection_Refresh() Implements _Collection.Refresh
            Throw New NotImplementedException()
        End Sub

        Private ReadOnly Property _Collection_Count As Integer Implements _Collection.Count
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private Sub Fields__Append(Name As String, Type As DataTypeEnum, Optional DefinedSize As Integer = 0, Optional Attrib As FieldAttributeEnum = FieldAttributeEnum.adFldUnspecified) Implements ADODB.Fields._Append
            Throw New NotImplementedException()
        End Sub

        Private Sub Fields_Append(Name As String, Type As DataTypeEnum, Optional DefinedSize As Integer = 0, Optional Attrib As FieldAttributeEnum = FieldAttributeEnum.adFldUnspecified, Optional FieldValue As Object = Nothing) Implements ADODB.Fields.Append
            Throw New NotImplementedException()
        End Sub

        Private Sub Fields_Update() Implements ADODB.Fields.Update
            Throw New NotImplementedException()
        End Sub

        Private Sub Fields_Resync(Optional ResyncValues As ResyncEnum = ResyncEnum.adResyncAllValues) Implements ADODB.Fields.Resync
            Throw New NotImplementedException()
        End Sub

        Private Sub Fields_CancelUpdate() Implements ADODB.Fields.CancelUpdate
            Throw New NotImplementedException()
        End Sub

        Private Function Fields_GetEnumerator() As IEnumerator Implements ADODB.Fields.GetEnumerator
            Return _listFields.GetEnumerator()
        End Function

        Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _listFields.GetEnumerator()
        End Function
    End Class
End Namespace