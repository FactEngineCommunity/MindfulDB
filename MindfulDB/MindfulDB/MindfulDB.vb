Imports System.Reflection
Imports FactEngineForServices

Namespace MindfulDB

    Public Class MindfulDB

        Public Property ConnectionString As String
            Get
                Return prApplication.WorkingModel.TargetDatabaseConnectionString
            End Get
            Set(value As String)
                prApplication.WorkingModel.TargetDatabaseConnectionString = value
            End Set
        End Property

        Public DatabaseType As pcenumDatabaseType = pcenumDatabaseType.SQLite

        Public Property Model As FBM.Model
            Get
                Return prApplication.WorkingModel
            End Get
            Set(value As FBM.Model)
                prApplication.WorkingModel = value
            End Set
        End Property

        Public GraphDefinition As Graph.GraphProvider = Nothing

        Public Property Connection As FactEngine.DatabaseConnection
            Get
                Return prApplication.WorkingModel.DatabaseConnection
            End Get
            Set(value As FactEngine.DatabaseConnection)
                prApplication.WorkingModel.DatabaseConnection = value
            End Set
        End Property

        ''' <summary>
        ''' Parameerless Constructor
        ''' </summary>
        Public Sub New()

            prApplication = New tApplication()
            prApplication.WorkingModel = New FBM.Model
        End Sub



        Public Function ConnectToDatabase() As Boolean

            Try
                If prApplication.WorkingModel.DatabaseConnection Is Nothing Then
                    Call prApplication.WorkingModel.connectToDatabase()
                Else
                    Try
                        If Not prApplication.WorkingModel.DatabaseConnection.Connected Or prApplication.WorkingModel.DatabaseManager.Connection Is Nothing Then
                            prApplication.WorkingModel.connectToDatabase()
                        End If
                    Catch ex As Exception
                        prApplication.ThrowErrorMessage("Oops. Check the database conection configuration for the Model you are trying to connect to.", pcenumErrorType.Warning, , False)
                        Return False
                    End Try
                End If

                Return True
            Catch ex As Exception
                Dim lsMessage As String
                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage &= vbCrLf & vbCrLf & ex.Message
                prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace)

                Return False
            End Try
        End Function

        Public Function RetrieveRelationalDataStructure() As Boolean

            Dim lsErrorMessage As String = ""

            Try
                Dim lrReverseEngineer As New ODBCDatabaseReverseEngineer(Me.Model, Me.Model.TargetDatabaseConnectionString, False, Nothing, False)

                'ToDo - Add the SQLs to get the LABELS for FKs and PGSRelationTables 
                Call lrReverseEngineer.ReverseEngineerDatabase(lsErrorMessage, False, False, False)


                Call Me.GetEdgeLabelsFromTablesForeignKeys()
                Call Me.GetEdgeLabelsForPGSRelationTables()

                'Get the Graph Definition of the Model.
                Me.GraphDefinition = New Graph.GraphProvider(prApplication.WorkingModel.RDS)

                Return True

            Catch ex As Exception
                Dim lsMessage As String
                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage &= vbCrLf & vbCrLf & ex.Message
                prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace)

                Return False
            End Try
        End Function


        Public Function ProcessQuery(ByVal asQuery As String, ByVal asQueryLanguage As String) As ORMQL.Recordset

            Dim lsMessage As String
            Dim lrRecordset As New ORMQL.Recordset

            Try
                Dim lsQuery As String = asQuery

                If prApplication.WorkingModel.DatabaseConnection Is Nothing Then
                    Call prApplication.WorkingModel.connectToDatabase()
                Else
                    Try
                        If Not prApplication.WorkingModel.DatabaseConnection.Connected Or prApplication.WorkingModel.DatabaseManager.Connection Is Nothing Then
                            prApplication.WorkingModel.connectToDatabase()
                        End If
                    Catch ex As Exception
                        lsMessage = "Oops. Check the database conection configuration for the Model you are trying to connect to."
                        prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Warning, , False)
                        lrRecordset.ErrorString = lsMessage
                        Return lrRecordset
                    End Try
                End If

                If asQueryLanguage = "Cypher" Then
                    'Debugger.Break()

                    Dim lrParser = New openCypherTranspiler.openCypherParser.OpenCypherParser(Nothing)
                    Dim lrQueryNode As openCypherTranspiler.openCypherParser.AST.QueryNode
                    lrQueryNode = lrParser.Parse(lsQuery)
                    Dim plan = openCypherTranspiler.LogicalPlanner.LogicalPlan.ProcessQueryTree(lrQueryNode, Me.GraphDefinition, Nothing)
                    Dim sqlRender = New openCypherTranspiler.SQLRenderer.SQLRenderer(Me.GraphDefinition, Nothing)
                    lsQuery = sqlRender.RenderPlan(plan)

                    'Debugger.Break()
                End If

                Dim lrFEQLProcessor As New FEQL.Processor(prApplication.WorkingModel)

                lrFEQLProcessor.DatabaseManager = prApplication.WorkingModel.DatabaseManager


                lrRecordset = lrFEQLProcessor.DatabaseManager.GO(lsQuery)

                Return lrRecordset

            Catch ex As Exception
                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage &= vbCrLf & vbCrLf & ex.Message
                prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace)

                lrRecordset.ErrorString = ex.Message
                Return lrRecordset

            End Try

        End Function


        Private Sub GetEdgeLabelsFromTablesForeignKeys()

            Try
                Dim lsSQLQuery As String = ""

#Region "Get the SQLQuery"
                lsSQLQuery = "WITH RECURSIVE fk_constraints AS (" & Environment.NewLine &
    "SELECT" & Environment.NewLine &
    "    referencing.name AS referencing_table," & Environment.NewLine &
    "    TRIM(SUBSTR(referencing.sql, INSTR(referencing.sql, 'FOREIGN KEY') + LENGTH('FOREIGN KEY') + 1, INSTR(referencing.sql, 'REFERENCES') - INSTR(referencing.sql, 'FOREIGN KEY') - LENGTH('FOREIGN KEY') - 1)) AS foreign_key_constraint," & Environment.NewLine &
    "    SUBSTR(referencing.sql, INSTR(referencing.sql, 'REFERENCES') + LENGTH('REFERENCES') + 1) AS remaining_sql" & Environment.NewLine &
    "FROM" & Environment.NewLine &
    "    sqlite_master AS referencing" & Environment.NewLine &
    "WHERE" & Environment.NewLine &
    "    referencing.type = 'table'" & Environment.NewLine &
    "    AND referencing.sql LIKE '%FOREIGN KEY%'" & Environment.NewLine &
    "" & Environment.NewLine &
    "UNION ALL" & Environment.NewLine &
    "" & Environment.NewLine &
    "SELECT" & Environment.NewLine &
    "    referencing_table," & Environment.NewLine &
    "    TRIM(SUBSTR(remaining_sql, INSTR(remaining_sql, 'FOREIGN KEY') + LENGTH('FOREIGN KEY') + 1, INSTR(remaining_sql, 'REFERENCES') - INSTR(remaining_sql, 'FOREIGN KEY') - LENGTH('FOREIGN KEY') - 1))," & Environment.NewLine &
    "    SUBSTR(remaining_sql, INSTR(remaining_sql, 'REFERENCES') + LENGTH('REFERENCES') + 1)" & Environment.NewLine &
    "FROM" & Environment.NewLine &
    "    fk_constraints" & Environment.NewLine &
    "WHERE" & Environment.NewLine &
    "    remaining_sql LIKE '%FOREIGN KEY%'" & Environment.NewLine &
    ")" & Environment.NewLine &
    "SELECT" & Environment.NewLine &
    "    referencing_table," & Environment.NewLine &
    "    foreign_key_constraint," & Environment.NewLine &
    "    TRIM(SUBSTR(remaining_sql, 0, INSTR(remaining_sql, '('))) AS referenced_table," & Environment.NewLine &
    "    CASE" & Environment.NewLine &
    "        WHEN remaining_sql LIKE '%/* { Label:""%' THEN" & Environment.NewLine &
    "            TRIM(" & Environment.NewLine &
    "                SUBSTR(" & Environment.NewLine &
    "                    remaining_sql," & Environment.NewLine &
    "                    INSTR(remaining_sql, 'Label:""') + LENGTH('Label:""')," & Environment.NewLine &
    "                    INSTR(SUBSTR(remaining_sql, INSTR(remaining_sql, 'Label:""') + LENGTH('Label:""')), '""') - 1" & Environment.NewLine &
    "                )" & Environment.NewLine &
    "            )" & Environment.NewLine &
    "        ELSE" & Environment.NewLine &
    "            NULL" & Environment.NewLine &
    "    END AS label" & Environment.NewLine &
    "FROM" & Environment.NewLine &
    "    fk_constraints" & Environment.NewLine &
    "WHERE" & Environment.NewLine &
    "    foreign_key_constraint NOT LIKE '%REFERENCES%';"
#End Region

                Dim lrRecordset = Me.Model.DatabaseConnection.GO(lsSQLQuery)

                While Not lrRecordset.EOF

                    Dim lrRelation = Me.Model.RDS.Relation.Find(Function(x) x.OriginTable.Name = lrRecordset("referencing_table").Data _
                                                                And x.DestinationTable.Name = lrRecordset("referenced_table").Data.Replace("[", "").Replace("]", ""))
                    If lrRelation IsNot Nothing Then 'CodeSafe
                        lrRelation.SetLabel(lrRecordset("label").Data)
                    End If

                    lrRecordset.MoveNext()
                End While

            Catch ex As Exception
                Dim lsMessage As String
                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage &= vbCrLf & vbCrLf & ex.Message
                prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace)
            End Try
        End Sub

        Private Sub GetEdgeLabelsForPGSRelationTables()

            Try
                Dim lsSQLQuery As String = ""

#Region "Get the SQLQuery"
                lsSQLQuery &= "WITH RECURSIVE pk_constraints AS (" & Environment.NewLine &
    "SELECT" & Environment.NewLine &
    "    referencing.name AS referencing_table," & Environment.NewLine &
    "    TRIM(SUBSTR(referencing.sql, INSTR(referencing.sql, 'CONSTRAINT') + LENGTH('CONSTRAINT'), INSTR(referencing.sql, 'PRIMARY KEY') - INSTR(referencing.sql, 'CONSTRAINT') - LENGTH('CONSTRAINT'))) AS primary_key_constraint," & Environment.NewLine &
    "    SUBSTR(referencing.sql, INSTR(referencing.sql, 'PRIMARY KEY') + LENGTH('PRIMARY KEY') + 1) AS remaining_sql" & Environment.NewLine &
    "FROM" & Environment.NewLine &
    "    sqlite_master AS referencing" & Environment.NewLine &
    "WHERE" & Environment.NewLine &
    "    referencing.type = 'table'" & Environment.NewLine &
    "    AND referencing.sql LIKE '%PRIMARY KEY%'" & Environment.NewLine &
    "    AND referencing.sql LIKE '%/* { Label:""%'" & Environment.NewLine &
    Environment.NewLine &
    "UNION ALL" & Environment.NewLine &
    Environment.NewLine &
    "SELECT" & Environment.NewLine &
    "    referencing_table," & Environment.NewLine &
    "    TRIM(SUBSTR(remaining_sql, INSTR(remaining_sql, 'CONSTRAINT') + LENGTH('CONSTRAINT'), INSTR(remaining_sql, 'PRIMARY KEY') - INSTR(remaining_sql, 'CONSTRAINT') - LENGTH('CONSTRAINT')))," & Environment.NewLine &
    "    SUBSTR(remaining_sql, INSTR(remaining_sql, 'PRIMARY KEY') + LENGTH('PRIMARY KEY') + 1)" & Environment.NewLine &
    "FROM" & Environment.NewLine &
    "    pk_constraints" & Environment.NewLine &
    "WHERE" & Environment.NewLine &
    "    remaining_sql LIKE '%PRIMARY KEY%'" & Environment.NewLine &
    "    AND remaining_sql REGEXP '^(?!.*FOREIGN KEY.*{ Label).*{ Label'" & Environment.NewLine &
    ")" & Environment.NewLine &
    Environment.NewLine &
    "SELECT" & Environment.NewLine &
    "    referencing_table," & Environment.NewLine &
    "    primary_key_constraint," & Environment.NewLine &
    "    CASE" & Environment.NewLine &
    "        WHEN remaining_sql REGEXP '^(?!.*FOREIGN KEY.*{ Label).*{ Label' THEN" & Environment.NewLine &
    "            TRIM(" & Environment.NewLine &
    "                REPLACE(" & Environment.NewLine &
    "                    SUBSTR(" & Environment.NewLine &
    "                        remaining_sql," & Environment.NewLine &
    "                        INSTR(remaining_sql, '/* { Label:') + LENGTH('/* { Label:')," & Environment.NewLine &
    "                        INSTR(remaining_sql, '} */') - INSTR(remaining_sql, '/* { Label:') - LENGTH('/* { Label:')" & Environment.NewLine &
    "                    )," & Environment.NewLine &
    "                    '""',''" & Environment.NewLine &
    "                )" & Environment.NewLine &
    "            )" & Environment.NewLine &
    "        ELSE" & Environment.NewLine &
    "            NULL" & Environment.NewLine &
    "    END AS label," & Environment.NewLine &
    "    remaining_sql" & Environment.NewLine &
    "FROM" & Environment.NewLine &
    "    pk_constraints" & Environment.NewLine &
    "WHERE" & Environment.NewLine &
    "   remaining_sql REGEXP '^(?!.*FOREIGN KEY.*{ Label).*{ Label'"
#End Region

                Dim lrRecordset = Me.Model.DatabaseConnection.GO(lsSQLQuery)

                While Not lrRecordset.EOF

                    Dim lrTable = Me.Model.RDS.Table.Find(Function(x) x.Name = lrRecordset("referencing_table").Data)
                    If lrTable IsNot Nothing Then 'CodeSafe
                        lrTable.SetLabel(lrRecordset("label").Data)
                    End If

                    lrRecordset.MoveNext()
                End While

            Catch ex As Exception
                Dim lsMessage As String
                Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

                lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
                lsMessage &= vbCrLf & vbCrLf & ex.Message
                prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace)
            End Try

        End Sub

    End Class

End Namespace
