Imports ADODB

Namespace Database
    Public Class FieldClass
        Implements ADODB.Field

        Private fieldName As String
        Private fieldType As ADODB.DataTypeEnum
        Private fieldValue As Object

        Public Sub New(ByVal name As String, ByVal type As ADODB.DataTypeEnum, ByVal value As Object)
            fieldName = name
            fieldType = type
            fieldValue = value
        End Sub

        Public Property Value As Object Implements Field.Value
            Get
                Return Me.fieldValue
            End Get
            Set(value As Object)
                Me.fieldValue = value
            End Set
        End Property

        Private Property Field20_Value As Object Implements Field20.Value
            Get
                Return Me.fieldValue
            End Get
            Set(value As Object)
                Me.fieldValue = value
            End Set
        End Property

        Public ReadOnly Property ActualSize() As Integer Implements ADODB.Field.ActualSize
            Get
                Return Len(fieldValue)
            End Get
        End Property

        Public Property DefinedSize() As Integer Implements ADODB.Field.DefinedSize
            Get
                Return 0
            End Get
            Set(ByVal value As Integer)
                ' Do nothing
            End Set
        End Property

        Public Property Name() As String Implements ADODB.Field.Name
            Get
                Return fieldName
            End Get
            Set(ByVal value As String)
                ' Do nothing
            End Set
        End Property

        Public Property Type() As ADODB.DataTypeEnum Implements ADODB.Field.Type
            Get
                Return fieldType
            End Get
            Set(ByVal value As ADODB.DataTypeEnum)
                ' Do nothing
            End Set
        End Property

        Public Property Precision() As Byte Implements ADODB.Field.Precision
            Get
                Return 0
            End Get
            Set(value As Byte)
                Me.Field20_Precision = value
            End Set
        End Property

        Public Property NumericScale() As Byte Implements ADODB.Field.NumericScale
            Get
                Return 0
            End Get
            Set(value As Byte)
                Me.Field20_NumericScale = value
            End Set
        End Property

        Public ReadOnly Property OriginalValue() As Object Implements ADODB.Field.OriginalValue
            Get
                Return fieldValue
            End Get
        End Property

        Public ReadOnly Property UnderlyingValue() As Object Implements ADODB.Field.UnderlyingValue
            Get
                Return fieldValue
            End Get
        End Property

        Public Property DataFormat() As Object Implements ADODB.Field.DataFormat
            Get
                Return Nothing
            End Get
            Set(value As Object)
                Me.Field20_DataFormat = value
            End Set
        End Property


        Public Sub AppendChunk(Data As Object) Implements Field.AppendChunk
            Throw New NotImplementedException()
        End Sub

        Public Function GetChunk(Length As Integer) As Object Implements Field.GetChunk
            Throw New NotImplementedException()
        End Function

        Public ReadOnly Property Properties As ADODB.Properties Implements Field.Properties
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private Property Field_Attributes As Integer Implements Field.Attributes
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Integer)
                Throw New NotImplementedException()
            End Set
        End Property

        Public ReadOnly Property Status As Integer Implements Field.Status
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private Sub Field20_AppendChunk(Data As Object) Implements Field20.AppendChunk
            Throw New NotImplementedException()
        End Sub

        Private Function Field20_GetChunk(Length As Integer) As Object Implements Field20.GetChunk
            Throw New NotImplementedException()
        End Function

        Private ReadOnly Property Field20_Properties As ADODB.Properties Implements Field20.Properties
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private ReadOnly Property Field20_ActualSize As Integer Implements Field20.ActualSize
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private Property Field20_Attributes As Integer Implements Field20.Attributes
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Integer)
                Throw New NotImplementedException()
            End Set
        End Property

        Private Property Field20_DefinedSize As Integer Implements Field20.DefinedSize
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Integer)
                Throw New NotImplementedException()
            End Set
        End Property

        Private ReadOnly Property Field20_Name As String Implements Field20.Name
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private Property Field20_Type As DataTypeEnum Implements Field20.Type
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As DataTypeEnum)
                Throw New NotImplementedException()
            End Set
        End Property

        Private Property Field20_Precision As Byte Implements Field20.Precision
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Byte)
                Throw New NotImplementedException()
            End Set
        End Property

        Private Property Field20_NumericScale As Byte Implements Field20.NumericScale
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Byte)
                Throw New NotImplementedException()
            End Set
        End Property

        Private ReadOnly Property Field20_OriginalValue As Object Implements Field20.OriginalValue
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private ReadOnly Property Field20_UnderlyingValue As Object Implements Field20.UnderlyingValue
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private Property Field20_DataFormat As Object Implements Field20.DataFormat
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Object)
                Throw New NotImplementedException()
            End Set
        End Property

        Private ReadOnly Property _ADO_Properties As ADODB.Properties Implements _ADO.Properties
            Get
                Throw New NotImplementedException()
            End Get
        End Property

    End Class

End Namespace