﻿Imports System.Reflection

Public Class ProgressObject

    Public IsError As Boolean
    Public Message As String
    Public SimpleAppend As Boolean = False 'As when adding a .
    Public Sub New()
    End Sub

    Public Sub New(ByVal abIsError As Boolean, ByVal asMessage As String, Optional ByVal abSimpleAppend As Boolean = False)
        Me.SimpleAppend = abSimpleAppend
        Me.IsError = abIsError
        Me.Message = asMessage
    End Sub

End Class

Public Enum pcenumReverseEngineeringErrorReportType
    Information
    [Error]
End Enum

Public Class ODBCDatabaseReverseEngineer

    Private Model As FBM.Model 'The model to which the database schema is to be mapped.

    Private ODBCConnection As System.Data.Odbc.OdbcConnection

    ''' <summary>
    ''' The RDS structure, from the database, is first mapped to this model.
    '''   NB Have to do this because creating the ModelElements in Me.Model will actually create the RDS structure for that Model.
    '''   The database structure is first loaded into TempModel, then the ModelElements created in Me.Model from TempModel.
    ''' </summary>
    Private TempModel As New FBM.Model

    ''' <summary>
    ''' True if the user elects to create a Page in the Model for each Table.
    ''' </summary>
    Private CreatePagePerTable As Boolean = False

    Private ProgressPercentage As Integer = 0

    Private [BackgroundWorker] As System.ComponentModel.BackgroundWorker = Nothing

    Private ShowExtraInformation As Boolean = True

    '=========================================================================================================================================================
    ''' <summary>
    ''' Constructor.
    ''' </summary>
    ''' <param name="arModel"></param>
    Public Sub New(ByRef arModel As FBM.Model,
                   ByVal asDatabaseConnectionString As String,
                   ByVal abCreatePagePerTable As Boolean,
                   Optional ByRef arBackgroundWorker As System.ComponentModel.BackgroundWorker = Nothing,
                   Optional ByVal abShowExtraInformation As Boolean = True)
        Me.Model = arModel
        Me.ODBCConnection = New System.Data.Odbc.OdbcConnection(asDatabaseConnectionString)
        Me.CreatePagePerTable = abCreatePagePerTable
        Me.BackgroundWorker = arBackgroundWorker
        Me.ShowExtraInformation = abShowExtraInformation
    End Sub

    ''' <summary>
    ''' Reverse engineers the Model from the database of that Model.
    ''' PRECONDITIONS: The Model has a linked database.
    ''' </summary>
    Public Function ReverseEngineerDatabase(ByRef asErrorMessage As String,
                                            Optional ByVal abReverseEngineeringKeepDatabaseColumnNames As Boolean = False,
                                            Optional ByVal abReverseEngineeringCreateEntityTypesForTablesWithNoColumns As Boolean = False,
                                            Optional ByVal abPopulateFBMModel As Boolean = False) As Boolean

        '=====================================================================================================================
        'PSEUDOCODE
        '
        ' * Connect to the ODBC database
        ' * Get the DataTypes for the database
        ' * Get the Tables...and put them into TempModel
        ' * Get the Columns for the TempModel Tables
        ' * Get the Indexes for the TempModel Tables
        ' * Get the Relations for the TempModel Tables
        ' * Order the TempModel Tables by their relations. Those that have no relations first.
        ' * Create EntityTypes for those Tables that have a Single Column PrimaryKey.
        ' * Create EntityTypes for those Tables that are not subtypes of other tables and have no Columns
        ' * For each Table in TempModel Tables (by their sorted order)
        '     * Create the ValueTypes for Columns that that are ValueTypes (even if they are referenced by a Relation)
        '         Value Types are, in this instance,:
        '           1. Not the ValueTypes created for ReferenceShemes of EntityTypes created for Tables with single Column PrimaryKeys.
        '              a) Including those Columns/ValueTypes that are referenced by a Column that references an EntityType for a Table with a single column PrimaryKey.
        '           2. Are Columns/ValueTypes where the Column has no Relation/reference to another table.
        '     * If the Table's ModelElement has not been created, create it.
        '         Can be an EntityType or an ObjectifiedFactType.
        '     * If the Table is an EntityType that has a Compound Reference Scheme, create the ReferenceScheme for EntityTypee/Table.
        '         NB The referenced ModelElements should have already been created.
        '     * If the Table is an ObjectifiedFactType then create the ObjectifiedFactTyhpe.
        '         NB The referenced ModelElements should have already been created.
        '   Loop
        ' * For each Table in the TempModel Tables (by their sorted order)
        '    * Create the ExternalUniquenessConstraints for Indexes of the Table
        '   Loop
        '=====================================================================================================================
        Try
            Dim lrTable As RDS.Table

            Me.TempModel.TargetDatabaseType = Me.Model.TargetDatabaseType
            Me.TempModel.TargetDatabaseConnectionString = Me.Model.TargetDatabaseConnectionString
            Me.TempModel.Port = Me.Model.Port
            Me.TempModel.Database = Me.Model.Database
            Me.TempModel.Server = Me.Model.Server

            Me.Model.connectToDatabase()
            Me.TempModel.connectToDatabase()

            'Load DatabaseTypes into memory.
            Call Me.Model.DatabaseConnection.getDatabaseTypes()
            Call Me.SetProgressBarValue(5, "Loaded Data Types.")

            'Call Me.GetDataTypes()

            Call Me.getTables()
            Call Me.SetProgressBarValue(10, "Loaded Tables.")

            lrTable = Me.TempModel.RDS.Table.Find(Function(x) x.Name = "location-hierarchy")

            Call Me.GetColumns()
            Call Me.SetProgressBarValue(20, "Loaded Columns for Tables.")

            Call Me.GetIndexes()
            Call Me.SetProgressBarValue(30, "Loaded Indexes for Tables.")

            Call Me.getRelations()
            Call Me.SetProgressBarValue(40, "Loaded Relations for Tables. ")

            Call Me.makeCapCamelCaseNames()

            If abPopulateFBMModel Then



                Call Me.TempModel.RDS.orderTablesByRelations()
                Call Me.SetProgressBarValue(45)

                '------------------------------------------------------------------------------
                'Create EntityTypes for each Table with a PrimaryKey with one Column.
                Call Me.createTablesForSingleColumnPKTables()
                Call Me.SetProgressBarValue(60, "Created Entit Types/Tables for Single Primary Key Column Tables.")

                'Create EntityTypes for those Tables that are not subtypes of other tables and have no Columns
                If abReverseEngineeringCreateEntityTypesForTablesWithNoColumns Then
                    Call Me.createEntityTypesForTablesThatAreNotSubtypesOfOtherTablesAndHaveNoColumns(".Id", "String")
                    Call Me.SetProgressBarValue(65, "Created Entity Types/Tables for those Tables that are not subtypes of other tables and have no Columns")
                End If

                'Create Subtype Relationships
                Call Me.createSubtypeRelationships()

                Call Me.SetProgressBarValue(70, "Creating other Value Types. ")

                For Each lrTable In Me.TempModel.RDS.Table

                    'Create ValueTypes (that haven't already been created by virtue of being the ReferenceModeValueType of Simple Reference Scheme EntityTypes.
                    Call Me.AppendProgress(".")
                    Call Me.createValueTypesByTable(lrTable)

#Region "Create ObjectifiedFactTypes"
                    Dim lsMessage As String = ""
                    Call Me.AppendProgress(".")
                    If Me.Model.GetModelObjectByName(lrTable.Name) Is Nothing Then
                        'The Table has no ModelElement, so create it.
                        If lrTable.getPrimaryKeyColumns.Count = 1 And lrTable.PrimarySupertype = "Entity" Then
                            'Is an EntityType, and should already be a ModelElement in the Model.
                            Call Me.ReportError($"Error: Creating Objectified Fact Types: {lrTable.Name} should already be a Model Element/Entity Type in the Model because it has a Single Column Primary Key.")
                        Else
                            'Create an ObjectifiedFactType
                            Dim lrFactType As FBM.FactType

                            Dim lrModelElement As FBM.ModelObject = Nothing
                            Dim larModelObject As New List(Of FBM.ModelObject)

                            For Each lrColumn In lrTable.getPrimaryKeyColumns
                                If lrColumn.hasOutboundRelation Then
                                    Dim lrRelation = lrColumn.Relation.Find(Function(x) x.OriginTable IsNot Nothing)
                                    Dim lrDestinationTable As RDS.Table = Me.Model.RDS.getTableByName(lrRelation.DestinationTable.Name)

                                    If lrDestinationTable Is Nothing Then
                                        Throw New Exception($"Creating Objectified Fact Types. For: {lrTable.Name}. Destination Table does not exist: " & lrRelation.DestinationTable.Name)
                                    End If

                                    lrModelElement = lrDestinationTable.FBMModelElement
                                    If lrModelElement Is Nothing Then
                                        lrModelElement = Me.Model.GetModelObjectByName(lrColumn.Name)
                                    End If
                                Else
                                    lrModelElement = Me.Model.GetModelObjectByName(lrColumn.Name)
                                End If

                                If lrModelElement.isReferenceModeValueType Then
                                    lrModelElement = Me.Model.getEntityTypeByReferenceModeValueType(lrModelElement)
                                End If
                                larModelObject.Add(lrModelElement)
                            Next

                            'FactTypes joining FactTypes only have one Role. See TimetableBookings FT in University model.
                            Dim larFTModelElement = (From ModelElement In larModelObject
                                                     Where ModelElement.GetType = GetType(FBM.FactType)
                                                     Select ModelElement Distinct).ToList

                            If larFTModelElement.Count > 0 Then
                                While larModelObject.Contains(larFTModelElement.First) And larModelObject.FindAll(Function(x) x Is larFTModelElement.First).Count > 1
                                    larModelObject.Remove(larFTModelElement.First)
                                End While
                            End If

                            If larModelObject.Count > 0 Then

                                lrFactType = Me.Model.CreateFactType(lrTable.Name, larModelObject, False, True, , , False, )
                                Me.Model.AddFactType(lrFactType)
                                lrFactType.Objectify(True) 'Add to model first, so LinkFactTypes have something to join to.

                                If lrTable.DerivationRule IsNot Nothing Then
                                    lrFactType.IsDerived = True
                                    lrFactType.DerivationText = lrTable.DerivationRule
                                End If

                                'Create the internalUniquenessConstraint.
                                Dim larRole As New List(Of FBM.Role)
                                For Each lrRole In lrFactType.RoleGroup
                                    larRole.Add(lrRole)
                                Next
                                Call lrFactType.CreateInternalUniquenessConstraint(larRole)

                                'Role Names
                                Dim lasUsedRoleName As New List(Of String)
                                Dim lsRoleName As String = ""
                                For Each lrRole In lrFactType.RoleGroup
                                    Dim larRolePlayed = From Table In Me.TempModel.RDS.Table
                                                        From RolePlayed In Table.RolesPlayed
                                                        Where LCase(Table.Name) = LCase(lrRole.JoinedORMObject.Id)
                                                        Where LCase(RolePlayed.RelationName) = LCase(lrFactType.Id)
                                                        Where Not lasUsedRoleName.Contains(RolePlayed.RoleName)
                                                        Select RolePlayed

                                    If larRolePlayed.Count > 0 Then

                                        lsRoleName = larRolePlayed.First.RoleName

                                        Dim larRolePlayed2 = From Table In Me.TempModel.RDS.Table
                                                             From RolePlayed In Table.RolesPlayed
                                                             Where LCase(Table.Name) = LCase(lrRole.JoinedORMObject.Id)
                                                             Where LCase(RolePlayed.RelationName) = LCase(lrFactType.Id)
                                                             Where Not lasUsedRoleName.Contains(RolePlayed.RoleName)
                                                             Select RolePlayed

                                        If larRolePlayed2.Count > 0 Then
                                            lsRoleName = larRolePlayed2.First.RoleName
                                        End If

                                        lrRole.Name = lsRoleName
                                        lasUsedRoleName.Add(lsRoleName)
                                    End If
                                Next

                                '-----------------------------------------------------------------------------------------------
                                'Create the FactTypeReading                            
                                If lrFactType.Arity <= 4 Then
                                    Dim larRoleGroupPermutations As IEnumerable(Of IEnumerable(Of FBM.Role))

                                    larRoleGroupPermutations = lrFactType.RoleGroup.Permute

                                    Dim liInd = 1
                                    For Each larRoleGroup In larRoleGroupPermutations
                                        If liInd <= 3 Then
                                            Dim lrSentence As New Language.Sentence("random sentence")
                                            Dim liInd2 = 1
                                            For Each lrRole In larRoleGroup.ToList
                                                If liInd2 < larRoleGroup.ToList.Count Then
                                                    lrSentence.PredicatePart.Add(New Language.PredicatePart("has"))
                                                Else
                                                    lrSentence.PredicatePart.Add(New Language.PredicatePart(""))
                                                End If
                                                liInd2 += 1
                                            Next
                                            Dim lrFactTypeReading As New FBM.FactTypeReading(lrFactType, larRoleGroup.ToList, lrSentence)
                                            lrFactType.FactTypeReading.Add(lrFactTypeReading)
                                        End If
                                        liInd += 1
                                    Next
                                End If


                                'Make sure the new Column Names (in Model rather than TempModel) are the same as what they are in the origin database.
                                'PSEUDOCODE
                                ' * For Each lrColumn in Model Table.GetPrimaryKeyColumns
                                '     * Find the Relation for that Column in the TempModel.
                                '     * If the Relation exists (i.e. is not a Column to a ValueType, but rather another ET or Objectified FT)
                                '         * Get the original OriginColumn Name and give it to lrColumn
                                '       End If
                                '   Next Column
                                Dim lrModelTable As RDS.Table = Me.Model.RDS.getTableByName(lrTable.Name)
                                Dim larCoveredColumn As New List(Of RDS.Column)
                                Dim lrColumn As RDS.Column
                                For Each lrTempColumn In lrTable.getPrimaryKeyColumns

                                    If lrTempColumn.hasOutboundRelation Then
                                        Dim lrTempColumPointsToTable As RDS.Table = lrTempColumn.Relation.Find(Function(x) x.OriginTable Is lrTempColumn.Table).DestinationTable
                                        Dim lrTempRelation As RDS.Relation = lrTempColumn.Relation.Find(Function(x) x.OriginTable Is lrTempColumn.Table)

                                        Dim larColumn = From Relation In Me.Model.RDS.Relation
                                                        From Column In Relation.OriginColumns
                                                        Where Relation.OriginTable.Name = lrTable.Name
                                                        Where Relation.OriginColumns(0).isPartOfPrimaryKey
                                                        Where Relation.DestinationTable.Name = lrTempColumPointsToTable.Name
                                                        Where Not larCoveredColumn.Contains(Column)
                                                        Select Column

                                        Try
                                            lrColumn = larColumn.First
                                            lrColumn.setName(lrTempColumn.Name)

                                            larCoveredColumn.Add(lrColumn)
                                        Catch
                                            lsMessage = "Error: Table: " & lrTable.Name & ". Trying to find Origin Column for releation from " & lrTable.Name & " to " & lrTempColumPointsToTable.Name & "."
                                            Call Me.ReportError(lsMessage)
                                        End Try


                                    End If
                                Next

                            Else
                                lsMessage = "Error: Creating Objectified Fact Types: For Table, " & lrTable.Name & "."
                                If lrTable.getPrimaryKeyColumns.Count > 0 Then
                                    lsMessage.AppendString(" Can't find Model Elements for the following: ")
                                    For Each lrColumn In lrTable.getPrimaryKeyColumns
                                        lsMessage.AppendLine(lrColumn.Name)
                                    Next
                                Else
                                    lsMessage.AppendString(" Table has no primary key columns")
                                End If

                                Call Me.ReportError(lsMessage)
                            End If
                        End If
                    End If
                    Call Me.AppendProgress(".")
#End Region
                Next

                '-----------------------------------------------------------------------------
                'Create Tables for all missing Tables.
                Call Me.SetProgressBarValue(85, "Creating Tables/Objectified Fact Types for remaining Tables/Relations.")
                '20211001-VM-Add code here for instance for Periodic-Event in TypeDB Social Network, where the PrimaryKey is
                'not for what it relates, but rather the Columns, Start-Date and End-Date.
                Call Me.createTablesForAllMissingTables()

                '-----------------------------------------------------------------------------
                'Create FactTypes for all other Relations.
                Call Me.createFactTypesForAllOtherRelations()

                '-----------------------------------------------------------------------------
                'Create FactTypes that are from a ModelElement straight to a ValueType.
                Dim lasNonReferenceModeValueTypeNames = From ValueType In Me.Model.ValueType
                                                        Where Not ValueType.isReferenceModeValueType
                                                        Select LCase(ValueType.Id)

                Call Me.SetProgressBarValue(90, "Creating Fact Types straight to Value Types.")
                For Each lrTable In Me.TempModel.RDS.Table

                    Dim larValueTypeColumns = From Column In lrTable.Column
                                              Where lasNonReferenceModeValueTypeNames.Contains(LCase(Column.Name))
                                              Where Not Column.isPartOfPrimaryKey
                                              Select Column

                    For Each lrColumn In larValueTypeColumns
                        Dim larModelElement As New List(Of FBM.ModelObject)
                        Dim lrModelElement1 As FBM.ModelObject = Nothing
                        Dim lrModelElement2 As FBM.ModelObject = Nothing

                        Dim lrModelTable As RDS.Table = Me.Model.RDS.getTableByName(lrTable.Name)
                        If lrModelTable IsNot Nothing Then
                            lrModelElement1 = lrModelTable.FBMModelElement

                            lrModelElement2 = Me.Model.GetModelObjectByName(lrColumn.Name)

                            If lrModelElement2 Is Nothing Then
                                Me.ReportError("Cannot find a Model Element in the Model with the name, " & lrColumn.Name & ", for addition to the " & lrModelTable.Name & " table.")
                                GoTo NextValueTypeColumn
                            End If

                            larModelElement.Add(lrModelElement1)
                            larModelElement.Add(lrModelElement2)

                            Dim lsFactTypeName As String = lrModelElement1.Id & "Has" & lrModelElement2.Id
                            lsFactTypeName = Me.Model.CreateUniqueFactTypeName(lsFactTypeName, 0)

                            Dim lrFactType As FBM.FactType = Me.Model.CreateFactType(lsFactTypeName, larModelElement, False, True, , , False, )
                            Me.Model.AddFactType(lrFactType)

                            'Create the internalUniquenessConstraint.
                            Dim larRole As New List(Of FBM.Role)
                            larRole.Add(lrFactType.RoleGroup(0))

                            Call lrFactType.CreateInternalUniquenessConstraint(larRole)

                            '-----------------------------------------------------------------------------------------------
                            'Create the FactTypeReadings
                            For liInd = 1 To 2
                                Dim lrSentence As New Language.Sentence("random sentence")
                                lrSentence.PredicatePart.Add(New Language.PredicatePart("has"))
                                lrSentence.PredicatePart.Add(New Language.PredicatePart(""))

                                Dim larRoleGroup As New List(Of FBM.Role)
                                If liInd = 1 Then
                                    larRoleGroup.Add(lrFactType.RoleGroup(0))
                                    larRoleGroup.Add(lrFactType.RoleGroup(1))
                                Else
                                    larRoleGroup.Add(lrFactType.RoleGroup(1))
                                    larRoleGroup.Add(lrFactType.RoleGroup(0))
                                End If
                                Dim lrFactTypeReading As New FBM.FactTypeReading(lrFactType, larRoleGroup, lrSentence)
                                lrFactType.FactTypeReading.Add(lrFactTypeReading)
                            Next
                            For Each lrFactTypeReading In lrFactType.FactTypeReading.ToArray
                                Call lrFactType.SetFactTypeReading(lrFactTypeReading, False)
                            Next
                        Else
                            'Throw warning
                        End If
NextValueTypeColumn:
                    Next
                    Call Me.AppendProgress(".")
                Next

                'Create External Uniqueness Constraints
#Region "Create External Uniqueness Constraints"
                For Each lrTempTable In Me.TempModel.RDS.Table

                    For Each lrUniqueNonPKIndex In lrTempTable.Index.FindAll(Function(x) Not x.IsPrimaryKey And x.Unique)

                        If lrUniqueNonPKIndex.Column.Count > 1 Then

                            Try
                                Dim larRole As New List(Of FBM.Role)

                                Dim lrColumn As RDS.Column
                                For Each lrTempColumn In lrUniqueNonPKIndex.Column

                                    lrTable = Me.Model.RDS.getTableByName(lrTempTable.Name)

                                    lrColumn = lrTable.Column.Find(Function(x) LCase(x.Name) = LCase(lrTempColumn.Name))
                                    larRole.Add(lrColumn.ActiveRole)
                                Next

                                'Add the RoleConstraint without any RoleConstraintRoles, so that Index is created when the RoleConstraintRoles are added.
                                Dim lrRoleConstraint As New FBM.RoleConstraint(Me.Model,
                                                                               pcenumRoleConstraintType.ExternalUniquenessConstraint,
                                                                               New List(Of FBM.Role),
                                                                               True,,
                                                                               True)

                                Dim liInd As Integer = 1
                                For Each lrRole In larRole
                                    Dim lrRoleConstraintRole As New FBM.RoleConstraintRole(lrRole, lrRoleConstraint,,, liInd, True)
                                    Call lrRoleConstraint.AddRoleConstraintRole(lrRoleConstraintRole)
                                    liInd += 1
                                Next
                            Catch ex As Exception
                                'Don't fail the reverse engineering, just report the error and move on.
                                Call Me.ReportError("Failed to make Unique Index, " & lrUniqueNonPKIndex.Name & ", on Table," & lrTempTable.Name & ".")
                            End Try
                        End If

                    Next
                Next

#End Region

            Else
                Me.Model.RDS = Me.TempModel.RDS
            End If

#Region "Wordnet-NamesSingular"
            ''Change names to Singular, rather than Plural
            'Call Me.SetProgressBarValue(95, "Changing plural model element names to singular.")
            'For Each lrTable In Me.Model.RDS.Table

            '    Dim loLanguageGeneric As New Language.LanguageGeneric(My.Settings.WordNetDictionaryEnglishPath)

            '    Dim lsNewName As String = ""
            '    lsNewName = loLanguageGeneric.GetNounOverviewForWord(lrTable.Name)
            '    If lsNewName IsNot Nothing And lsNewName <> lrTable.Name Then
            '        Call lrTable.FBMModelElement.setName(MakeCapCamelCase(lsNewName), False, True)
            '    End If
            '    Call Me.AppendProgress(".")
            'Next
#End Region

            'Revert Column Names to Column.DatabaseName if necessary
            If abReverseEngineeringKeepDatabaseColumnNames Then
                Call Me.SetProgressBarValue(97, "Setting Column/Attribute names to the case in the database.")
                For Each lrTable In Me.Model.RDS.Table
                    For Each lrColumn In lrTable.Column
                        If lrColumn.ActiveRole.JoinedORMObject.DBName <> "" Then
                            lrColumn.setName(lrColumn.ActiveRole.JoinedORMObject.DBName)
                        End If
                    Next
                Next
            End If

            'CodeSafe-IsDirty Everything
            Call Me.IsDirtyEverything()

            Call Me.SetProgressBarValue(100)

        Catch ex As Exception
            Me.ReportError(ex.Message & vbCrLf & ex.StackTrace)

            Return False
        End Try

        Return True

    End Function

    Private Sub IsDirtyEverything()

        Try
            For Each lrValueType In Me.Model.ValueType
                lrValueType.isDirty = True
            Next
            For Each lrEntityType In Me.Model.EntityType
                lrEntityType.isDirty = True
            Next
            For Each lrFactType In Me.Model.FactType
                lrFactType.isDirty = True
            Next
            For Each lrRoleConstraint In Me.Model.RoleConstraint
                lrRoleConstraint.isDirty = True
            Next

        Catch ex As Exception
            Dim lsMessage As String
            Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

            lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
            lsMessage &= vbCrLf & vbCrLf & ex.Message
            prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace)
        End Try

    End Sub


    '20211001-VM-Add code here for instance for Periodic-Event in TypeDB Social Network, where the PrimaryKey is
    'not for what it relates, but rather the Columns, Start-Date and End-Date.
    Private Sub createTablesForAllMissingTables()

        Try
            Dim larTablesToCreate = From Table In Me.TempModel.RDS.Table
                                    Where Table.IsDBRelation
                                    Where Table.RelatedRoleNames.Count = 0
                                    Select Table

            For Each lrTable In larTablesToCreate

                'Create an ObjectifiedFactType
                Dim lrFactType As FBM.FactType

                Dim lrModelElement As FBM.ModelObject = Nothing
                Dim larModelObject As New List(Of FBM.ModelObject)

                For Each lrColumn In lrTable.Column
                    If lrColumn.hasOutboundRelation Then
                        Dim lrRelation = lrColumn.Relation.Find(Function(x) x.OriginTable IsNot Nothing)
                        Dim lrDestinationTable As RDS.Table = Me.Model.RDS.getTableByName(lrRelation.DestinationTable.Name)

                        If lrDestinationTable Is Nothing Then
                            Throw New Exception($"Creating Objectified Fact Types. For: {lrTable.Name}. Destination Table does not exist: " & lrRelation.DestinationTable.Name)
                        End If

                        lrModelElement = lrDestinationTable.FBMModelElement
                        If lrModelElement Is Nothing Then
                            lrModelElement = Me.Model.GetModelObjectByName(lrColumn.Name)
                        End If
                    Else
                        lrModelElement = Me.Model.GetModelObjectByName(lrColumn.Name)
                    End If

                    If lrModelElement.isReferenceModeValueType Then
                        lrModelElement = Me.Model.getEntityTypeByReferenceModeValueType(lrModelElement)
                    End If
                    larModelObject.Add(lrModelElement)
                Next

                'FactTypes joining FactTypes only have one Role. See TimetableBookings FT in University model.
                Dim larFTModelElement = (From ModelElement In larModelObject
                                         Where ModelElement.GetType = GetType(FBM.FactType)
                                         Select ModelElement Distinct).ToList

                If larFTModelElement.Count > 0 Then
                    While larModelObject.Contains(larFTModelElement.First) And larModelObject.FindAll(Function(x) x Is larFTModelElement.First).Count > 1
                        larModelObject.Remove(larFTModelElement.First)
                    End While
                End If

                If larModelObject.Count > 0 Then

                    lrFactType = Me.Model.CreateFactType(lrTable.Name, larModelObject, False, True, , , False, )
                    Me.Model.AddFactType(lrFactType)
                    lrFactType.Objectify(True) 'Add to model first, so LinkFactTypes have something to join to.

                    'Create the internalUniquenessConstraint.
                    Dim larRole As New List(Of FBM.Role)
                    For Each lrRole In lrFactType.RoleGroup
                        larRole.Add(lrRole)
                    Next
                    Call lrFactType.CreateInternalUniquenessConstraint(larRole)

                    '-----------------------------------------------------------------------------------------------
                    'Create the FactTypeReading                            
                    If lrFactType.Arity <= 4 Then
                        Dim larRoleGroupPermutations As IEnumerable(Of IEnumerable(Of FBM.Role))

                        larRoleGroupPermutations = lrFactType.RoleGroup.Permute

                        Dim liInd = 1
                        For Each larRoleGroup In larRoleGroupPermutations
                            If liInd <= 3 Then
                                Dim lrSentence As New Language.Sentence("random sentence")
                                Dim liInd2 = 1
                                For Each lrRole In larRoleGroup.ToList
                                    If liInd2 < larRoleGroup.ToList.Count Then
                                        lrSentence.PredicatePart.Add(New Language.PredicatePart("has"))
                                    Else
                                        lrSentence.PredicatePart.Add(New Language.PredicatePart(""))
                                    End If
                                    liInd2 += 1
                                Next
                                Dim lrFactTypeReading As New FBM.FactTypeReading(lrFactType, larRoleGroup.ToList, lrSentence)
                                lrFactType.FactTypeReading.Add(lrFactTypeReading)
                            End If
                            liInd += 1
                        Next
                    End If

                Else
                    Dim lsMessage As String = "Error: Creating Objectified Fact Types: For Table, " & lrTable.Name & "."
                    If lrTable.Column.Count > 0 Then
                        lsMessage.AppendString(" Can't find Model Elements for the following: ")
                        For Each lrColumn In lrTable.Column
                            lsMessage.AppendLine(lrColumn.Name)
                        Next
                    Else
                        lsMessage.AppendString(" Table has no primary key columns")
                    End If

                    Call Me.ReportError(lsMessage)
                End If

            Next

        Catch ex As Exception
            Me.ReportError(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub

    ''' <summary>
    ''' Creates Entity Types For Tables That Are Not Subtypes Of Other Tables And Have No Columns
    ''' </summary>
    Private Sub createEntityTypesForTablesThatAreNotSubtypesOfOtherTablesAndHaveNoColumns(ByVal asReverseEngineeringDefaultReferenceMode As String,
                                                                                          ByVal asReverseEngineeringDefaultPrimaryKeyDataType As String)

        Try
            Dim larTablesToCreate = From Table In Me.TempModel.RDS.Table
                                    Where Table.PrimarySupertype = "Entity"
                                    Where Table.Column.Count = 0


            For Each lrTable In larTablesToCreate

                Dim lrEntityType As FBM.EntityType
                lrEntityType = New FBM.EntityType(Me.Model, pcenumLanguage.ORMModel, lrTable.Name, lrTable.Name)
                Me.Model.AddEntityType(lrEntityType, True)
                lrEntityType.SetDBName(lrTable.DatabaseName)

                Dim lsValueTypeName As String
                Dim lsReferenceMode As String = ""
                lsValueTypeName = lrTable.Name & "_" & asReverseEngineeringDefaultReferenceMode
                lsReferenceMode = "." & asReverseEngineeringDefaultReferenceMode

                Dim liBostonDataType As pcenumORMDataType = Me.Model.DatabaseConnection.getBostonDataTypeByDatabaseDataType(asReverseEngineeringDefaultPrimaryKeyDataType)

                lrEntityType.SetReferenceMode(lsReferenceMode, False, lsValueTypeName, False, liBostonDataType, True)
                lrEntityType.ReferenceModeValueType.SetDBName(asReverseEngineeringDefaultReferenceMode)

            Next

        Catch ex As Exception
            Me.ReportError(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub


    Private Sub createSubtypeRelationships()

        Try
            For Each lrTable In Me.TempModel.RDS.Table

                Dim lrBaseModelelement As FBM.ModelObject = Me.Model.GetModelObjectByName(lrTable.Name)
                Dim lrModelElement As FBM.ModelObject = Me.Model.GetModelObjectByName(lrTable.PrimarySupertype)

                If lrModelElement IsNot Nothing And lrBaseModelelement Is Nothing Then
                    Select Case lrModelElement.GetType
                        Case = GetType(FBM.EntityType)
                            lrBaseModelelement = Me.Model.CreateEntityType(lrTable.Name, True)
                    End Select
                End If

                If lrBaseModelelement IsNot Nothing Then
                    Select Case lrBaseModelelement.GetType
                        Case Is = GetType(FBM.EntityType)
                            If lrModelElement IsNot Nothing Then

                                Select Case lrModelElement.GetType
                                    Case Is = GetType(FBM.EntityType)

                                        'Create the Subtype Relationship
                                        Call CType(lrBaseModelelement, FBM.EntityType).CreateSubtypeRelationship(CType(lrModelElement, FBM.EntityType), False)
                                End Select
                            End If
                    End Select
                End If
            Next

        Catch ex As Exception
            Me.ReportError(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub

    Private Sub createFactTypesForAllOtherRelations()

        Try
            Dim lsMessage As String = ""
            Call Me.SetProgressBarValue(80)

            For Each lrRelation In Me.TempModel.RDS.Relation.FindAll(Function(x) Not x.isPrimaryKeyBasedRelation).ToArray
                'Relations to other Tables.
                Dim larModelElement As New List(Of FBM.ModelObject)
                Dim lrModelElement1 As FBM.ModelObject = Nothing
                Dim lrModelElement2 As FBM.ModelObject = Nothing

                Try
                    Dim lrOriginTable As RDS.Table = Me.Model.RDS.getTableByName(lrRelation.OriginTable.Name)
                    If lrOriginTable Is Nothing Then
                        lsMessage = "Couldn't find Origin Table, " & lrRelation.OriginTable.Name & ", in the Model."
                        lsMessage.AppendDoubleLineBreak($"Relation between: Origin Table: {lrRelation.OriginTable.Name}, Destination Table: {lrRelation.DestinationTable.Name}")
                        Throw New Exception(lsMessage)
                    End If
                    lrModelElement1 = lrOriginTable.FBMModelElement

                    Dim lrDestinationTable = Me.Model.RDS.getTableByName(lrRelation.DestinationTable.Name)
                    If lrDestinationTable Is Nothing Then
                        Throw New Exception("Couldn't find Destination Table, " & lrRelation.DestinationTable.Name & ", in the Model.")
                    End If

                    lrModelElement2 = lrDestinationTable.FBMModelElement

                    larModelElement.Add(lrModelElement1)
                    larModelElement.Add(lrModelElement2)

                    Dim lsFactTypeName As String = lrModelElement1.Id & lrModelElement2.Id
                    lsFactTypeName = Me.Model.CreateUniqueFactTypeName(lsFactTypeName, 0)

                    Dim lrFactType As FBM.FactType = Me.Model.CreateFactType(lsFactTypeName, larModelElement, False, True, , , False, )
                    Me.Model.AddFactType(lrFactType)

                    'Create the internalUniquenessConstraint.
                    Dim larRole As New List(Of FBM.Role)
                    larRole.Add(lrFactType.RoleGroup(0))

                    Call lrFactType.CreateInternalUniquenessConstraint(larRole)

                    '-----------------------------------------------------------------------------------------------
                    'Create the FactTypeReadings
                    For liInd = 1 To 2
                        Dim lrSentence As New Language.Sentence("random sentence")
                        lrSentence.PredicatePart.Add(New Language.PredicatePart("has"))
                        lrSentence.PredicatePart.Add(New Language.PredicatePart(""))

                        Dim larRoleGroup As New List(Of FBM.Role)
                        If liInd = 1 Then
                            larRoleGroup.Add(lrFactType.RoleGroup(0))
                            larRoleGroup.Add(lrFactType.RoleGroup(1))
                        Else
                            larRoleGroup.Add(lrFactType.RoleGroup(1))
                            larRoleGroup.Add(lrFactType.RoleGroup(0))
                        End If
                        Dim lrFactTypeReading As New FBM.FactTypeReading(lrFactType, larRoleGroup, lrSentence)
                        lrFactType.FactTypeReading.Add(lrFactTypeReading)
                    Next
                    For Each lrFactTypeReading In lrFactType.FactTypeReading.ToArray
                        Call lrFactType.SetFactTypeReading(lrFactTypeReading, False)
                    Next

                    '-----------------------------------------------------------------------------------------------
                    'Name the Column based on the OriginColumn from the TempTable, because creating the 
                    '  RoleConstraint/ IUC(above) does not preserve the Column name, but names it after what it points to.
                    '  This may be a Column with a different name, in the DestinationTable.
                    Try
                        For Each lrColumn In lrOriginTable.Column.FindAll(Function(x) larRole.Contains(x.Role))
                            Dim lrDestinationColumn As RDS.Column = lrDestinationTable.Column.Find(Function(x) x.ActiveRole Is lrColumn.ActiveRole)
                            Dim lrTempDestinationColumn = lrRelation.DestinationColumns.Find(Function(x) LCase(x.Name) = LCase(lrDestinationColumn.Name))
                            Dim liSequenceNr As Integer = lrRelation.DestinationColumns.IndexOf(lrTempDestinationColumn)
                            Dim lsColumnName As String = lrRelation.OriginColumns(liSequenceNr).Name
                            Call lrColumn.setName(lsColumnName)
                        Next

                    Catch ex As Exception
                        'Not a biggie at this stage.
                        Call Me.ReportError("Information: Couldn't rename column/s in Table, " & lrOriginTable.Name & ". Sticking with the name/s generated.", pcenumReverseEngineeringErrorReportType.Information)
                    End Try

                Catch ex As Exception
                    Throw New Exception("Error creating Fact Type: " & ex.Message)
                End Try

                Call Me.AppendProgress(".")
            Next 'Relation in TempModel

            Call Me.SetProgressBarValue(85, "Created Fact Types for all other Relations.")

        Catch ex As Exception
            Dim lsMessage As String = ""
            lsMessage = "Creating Fact Types for all other Relations: " & ex.Message & "...Check to see that the relevant Table/s have a Primary Key set in the database."
            Throw New Exception(lsMessage)
        End Try

    End Sub

    ''' <summary>
    ''' Creates the ValueTypes for Columns that that are ValueTypes (even if they are referenced by a Relation)
    '''   Value Types are, in this instance,:
    '''     1. Not the ValueTypes created for ReferenceShemes of EntityTypes created for Tables with single Column PrimaryKeys.
    '''        a) Including those Columns/ValueTypes that are referenced by a Column that references an EntityType for a Table with a single column PrimaryKey.
    '''     2. Are Columns/ValueTypes where the Column has no Relation/reference to another table.
    ''' </summary>
    Private Sub createValueTypesByTable(ByRef arTable As RDS.Table)

        Try
            For Each lrColumn In arTable.Column

                If Not (lrColumn.isPartOfPrimaryKey Or lrColumn.hasOutboundRelation) Or
                    (lrColumn.isPartOfPrimaryKey And Not lrColumn.hasOutboundRelation) Then

                    Dim lrModelElement = Me.Model.GetModelObjectByName(lrColumn.Name)

                    If lrModelElement IsNot Nothing Then
                        If arTable.IsDBRelation And lrModelElement.GetType <> GetType(FBM.ValueType) Then
                            GoTo Skip
                        End If
                    End If

                    Dim lsUniqueValueTypeName As String
                    If lrModelElement IsNot Nothing Then
                        lsUniqueValueTypeName = lrColumn.Name
                        If lrModelElement.GetType <> GetType(FBM.ValueType) Then
                            lsUniqueValueTypeName = Me.Model.CreateUniqueValueTypeName(lrColumn.Name, 0)
                        End If
                    Else
                        lsUniqueValueTypeName = lrColumn.Name
                    End If

                    'Create the ValueType
                    Dim lrValueType As FBM.ValueType

                    lrValueType = New FBM.ValueType(Me.Model, pcenumLanguage.ORMModel, lrColumn.Name, True)
                    Try
                        lrValueType.DataType = Me.Model.DatabaseConnection.getBostonDataTypeByDatabaseDataType(lrColumn.DataType.DataType)
                    Catch ex As Exception
                        lrValueType.DataType = pcenumORMDataType.TextVariableLength
                    End Try

                    'Add the ValueType to the Model
                    Me.Model.AddValueType(lrValueType,,,, True)
                    lrValueType.SetDBName(lrColumn.DatabaseName)

Skip: 'Because is not a ValueType

                End If

            Next

        Catch ex As Exception
            Call Me.ReportError(ex.Message)
        End Try
    End Sub

    Private Sub DisplayData(ByRef table As DataTable)

        For Each row As DataRow In table.Rows
            For Each col As DataColumn In table.Columns
                MsgBox(col.ColumnName & " = " & row(col))
            Next
        Next

    End Sub

    Private Sub getRelations()

        Try
            For Each lrTable In Me.TempModel.RDS.Table
                lrTable.Model = Me.TempModel.RDS

                Dim larRelation As New List(Of RDS.Relation)
                Try
                    larRelation = Me.TempModel.DatabaseConnection.getForeignKeyRelationshipsByTable(lrTable)
                Catch ex As Exception
                    Me.ReportError(ex.Message)
                End Try

                For Each lrRelation In larRelation
                    Me.TempModel.RDS.Relation.Add(lrRelation)
                Next
            Next

        Catch ex As Exception
            Dim lsMessage As String
            Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

            lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
            lsMessage &= vbCrLf & vbCrLf & ex.Message
            prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace)
        End Try

    End Sub


    Private Sub getTables()

        Dim lrTable As RDS.Table

        Try

            For Each lrTable In Me.Model.DatabaseConnection.getTables()
                lrTable.Model = Me.TempModel.RDS
                Me.TempModel.RDS.Table.Add(lrTable)
            Next

        Catch ex As Exception
            Dim lsMessage As String
            Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

            lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
            lsMessage &= vbCrLf & vbCrLf & ex.Message
            prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace)
        End Try

    End Sub

    Private Sub createTablesForSingleColumnPKTables()

        Dim lrPage As FBM.Page

        For Each lrTable In Me.TempModel.RDS.Table

            If lrTable.hasSingleColumnPrimaryKey And lrTable.PrimarySupertype = "Entity" Then
                Dim lrEntityType As FBM.EntityType
                lrEntityType = New FBM.EntityType(Me.Model, pcenumLanguage.ORMModel, lrTable.Name, lrTable.Name)
                Me.Model.AddEntityType(lrEntityType, True)
                lrEntityType.SetDBName(lrTable.DatabaseName)

                Dim lsValueTypeName As String
                Dim lsReferenceMode As String = ""
                Dim lrPrimaryKeyColumn As RDS.Column
                lrPrimaryKeyColumn = lrTable.Index.Find(Function(x) x.IsPrimaryKey = True).Column(0)
                If lrPrimaryKeyColumn Is Nothing Then
                    Throw New Exception($"CreateTablesForSingleColumnPKTables: Table: {lrTable.Name}: No Primary Key Column found.")
                End If
                lsValueTypeName = lrPrimaryKeyColumn.Name

                Dim items As Array
                items = System.Enum.GetValues(GetType(pcenumReferenceModeEndings))
                Dim item As pcenumReferenceModeEndings
                For Each item In items
                    If lsValueTypeName.EndsWith(GetEnumDescription(item).Trim({"."c})) Then 'See https://msdn.microsoft.com/en-us/library/kxbw3kwc(v=vs.110).aspx
                        lsReferenceMode = GetEnumDescription(item).Trim({"."c})
                        Exit For
                    Else
                        lsReferenceMode = lsValueTypeName
                    End If
                Next


                Dim liBostonDataType As pcenumORMDataType = Me.Model.DatabaseConnection.getBostonDataTypeByDatabaseDataType(lrPrimaryKeyColumn.DataType.DataType)

                Dim lrVTTable As RDS.Table = Me.TempModel.RDS.Table.Find(Function(x) x.Name = lsValueTypeName)
                If lrVTTable IsNot Nothing Then
                    lsValueTypeName = lrEntityType.Id & "_" & lsValueTypeName
                End If

                '20220121-VM-Was. Moved test to top of If..then. If all fine, then remove this.
                'Select Case lrTable.PrimarySupertype
                'Case Is = "Entity"
                lrEntityType.SetReferenceMode(lsReferenceMode, False, lsValueTypeName, False, liBostonDataType, True)
                        lrEntityType.ReferenceModeValueType.SetDBName(lrPrimaryKeyColumn.DatabaseName)
                'End Select

            End If
            Call Me.AppendProgress(".")
        Next


    End Sub

    Private Sub GetIndexes()

        Dim lrTable As RDS.Table

        Try
            Dim larIndex As New List(Of RDS.Index)

            For Each lrTable In Me.TempModel.RDS.Table
                larIndex = Me.TempModel.DatabaseConnection.getIndexesByTable(lrTable)
                For Each lrIndex In larIndex
                    lrTable.Index.Add(lrIndex)
                    Me.TempModel.RDS.Index.Add(lrIndex)
                Next

                If larIndex.FindAll(Function(x) x.IsPrimaryKey).Count = 0 Then
                    'Need an alternate route for SQLite where a PK can be created that has no index.
                    larIndex = Me.TempModel.DatabaseConnection.getIndexesByTableByAlternateMeans(lrTable)
                    For Each lrIndex In larIndex
                        lrTable.Index.Add(lrIndex)
                    Next
                End If

            Next

        Catch ex As Exception
            Dim lsMessage As String
            Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

            lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
            lsMessage &= vbCrLf & vbCrLf & ex.Message
            Call Me.ReportError(lsMessage)
        End Try

    End Sub


    ''' <summary>
    ''' Used Predominantly with ODBC Connections where can retrieve the Data Types of a Database.
    ''' </summary>
    Private Sub GetDataTypes()

        Try
            Dim lrODBCTable As System.Data.DataTable

            If Not Me.ODBCConnection.State = ConnectionState.Open Then
                Me.ODBCConnection.Open()
            End If

            lrODBCTable = Me.ODBCConnection.GetSchema("DataTypes")

            For Each lrRow As DataRow In lrODBCTable.Rows

                Dim lrDataType As New RDS.DataType

                lrDataType.TypeName = lrRow(lrODBCTable.Columns("TypeName"))
                lrDataType.ProviderDBType = lrRow(lrODBCTable.Columns("ProviderDbType"))
                lrDataType.ColumnSize = lrRow(lrODBCTable.Columns("ColumnSize"))
                lrDataType.CreateFormat = NullVal(lrRow(lrODBCTable.Columns("CreateFormat")), "")
                lrDataType.CreateParameters = NullVal(lrRow(lrODBCTable.Columns("CreateParameters")), "")
                lrDataType.DataType = lrRow(lrODBCTable.Columns("Datatype"))
                lrDataType.IsAutoIncrementable = NullVal(lrRow(lrODBCTable.Columns("IsAutoIncrementable")), False)
                lrDataType.IsBestMatch = NullVal(lrRow(lrODBCTable.Columns("IsBestMatch")), False)
                lrDataType.IsCaseSensitive = NullVal(lrRow(lrODBCTable.Columns("IsCaseSensitive")), False)
                lrDataType.IsFixedLength = NullVal(lrRow(lrODBCTable.Columns("IsFixedLength")), False)
                lrDataType.IsFixedPrecisionScale = NullVal(lrRow(lrODBCTable.Columns("IsFixedPrecisionScale")), False)
                lrDataType.IsLong = NullVal(lrRow(lrODBCTable.Columns("IsLong")), False)
                lrDataType.IsNullable = NullVal(lrRow(lrODBCTable.Columns("IsNullable")), False)
                lrDataType.IsSearchable = NullVal(lrRow(lrODBCTable.Columns("IsSearchable")), False)
                lrDataType.IsSearchableWithLike = NullVal(lrRow(lrODBCTable.Columns("IsSearchableWithLike")), False)
                lrDataType.IsUnsigned = NullVal(lrRow(lrODBCTable.Columns("IsUnsigned")), False)
                lrDataType.MaximumScale = NullVal(lrRow(lrODBCTable.Columns("MaximumScale")), 0)
                lrDataType.MinimumScale = NullVal(lrRow(lrODBCTable.Columns("MinimumScale")), 0)
                lrDataType.IsConcurrencyType = NullVal(lrRow(lrODBCTable.Columns("IsConcurrencyType")), False)
                lrDataType.IsLiteralSupported = NullVal(lrRow(lrODBCTable.Columns("IsLiteralSupported")), False)
                lrDataType.LiteralPrefix = NullVal(lrRow(lrODBCTable.Columns("LiteralPrefix")), "")
                lrDataType.LiteralSuffix = NullVal(lrRow(lrODBCTable.Columns("LiteralSuffix")), "")
                lrDataType.SQLtype = NullVal(lrRow(lrODBCTable.Columns("SQLType")), 0)

                Me.Model.RDS.DataType.Add(lrDataType)
            Next

        Catch ex As Exception
            Dim lsMessage As String
            Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

            lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
            lsMessage &= vbCrLf & vbCrLf & ex.Message
            Call Me.ReportError(lsMessage)
        End Try

    End Sub

    Private Sub GetColumns()

        Try

            For Each lrTable In Me.TempModel.RDS.Table

                For Each lrColumn In Me.Model.DatabaseConnection.getColumnsByTable(lrTable)
                    lrTable.Column.Add(lrColumn)
                Next
            Next
        Catch ex As Exception
            Dim lsMessage As String
            Dim mb As MethodBase = MethodInfo.GetCurrentMethod()

            lsMessage = "Error: " & mb.ReflectedType.Name & "." & mb.Name
            lsMessage &= vbCrLf & vbCrLf & ex.Message
            prApplication.ThrowErrorMessage(lsMessage, pcenumErrorType.Critical, ex.StackTrace)
        End Try

    End Sub


    ''' <summary>
    ''' Only to be used on the TempModel.
    ''' </summary>
    Private Sub makeCapCamelCaseNames()

        Try
            'Tables
            For Each lrTable In Me.TempModel.RDS.Table
                lrTable.DatabaseName = lrTable.Name
                lrTable.Name = Viev.Strings.MakeCapCamelCase(lrTable.Name)

                'Columns 
                For Each lrColumn In lrTable.Column
                    lrColumn.DatabaseName = lrColumn.Name
                    lrColumn.Name = Viev.Strings.MakeCapCamelCase(lrColumn.Name)
                Next

                lrTable.PrimarySupertype = Viev.Strings.MakeCapCamelCase(lrTable.PrimarySupertype)
            Next


        Catch ex As Exception
            Call Me.ReportError("makeCapCamelCaseNames: " & ex.Message)
        End Try

    End Sub

    Private Sub ReportError(asErrorMessage As String,
                            Optional aiReverseEngineeringErrorType As pcenumReverseEngineeringErrorReportType = pcenumReverseEngineeringErrorReportType.Error)

        Dim lrProgressObject As New ProgressObject(True, asErrorMessage)

        Select Case aiReverseEngineeringErrorType
            Case Is = pcenumReverseEngineeringErrorReportType.Error
                Me.BackgroundWorker.ReportProgress(Me.ProgressPercentage, lrProgressObject)
            Case Is = pcenumReverseEngineeringErrorReportType.Information
                If Me.ShowExtraInformation Then
                    Me.BackgroundWorker.ReportProgress(Me.ProgressPercentage, lrProgressObject)
                End If
        End Select

    End Sub

    Private Sub SetProgressBarValue(ByVal aiValue As Integer, Optional asMessage As String = Nothing)

        Try
            Dim lrProgressObject As New ProgressObject(False, asMessage)
            Me.ProgressPercentage = aiValue
            If Me.BackgroundWorker IsNot Nothing Then
                Me.BackgroundWorker.ReportProgress(aiValue, lrProgressObject)
            End If

        Catch ex As Exception
            Call Me.ReportError(ex.Message)
        End Try
    End Sub

    Private Sub AppendProgress(ByVal asAppendString As String)

        Try
            Dim lrProgressObject As New ProgressObject(False, asAppendString, True)
            Me.BackgroundWorker.ReportProgress(Me.ProgressPercentage, lrProgressObject)

        Catch ex As Exception
            Call Me.ReportError(ex.Message)
        End Try
    End Sub




End Class
