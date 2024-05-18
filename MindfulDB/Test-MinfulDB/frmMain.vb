Imports MindfulDB

Public Class frmMain

    Private mrMindfulDB As MindfulDB.MindfulDB.MindfulDB

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.mrMindfulDB = New MindfulDB.MindfulDB.MindfulDB

        Me.mrMindfulDB.Model.TargetDatabaseType = FactEngineForServices.publicConstants.pcenumDatabaseType.SQLite
        Me.mrMindfulDB.ConnectionString = My.Settings.DatabaseConnectionString

        Call Me.mrMindfulDB.ConnectToDatabase()

        Call Me.mrMindfulDB.RetrieveRelationalDataStructure()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim lsQuery As String = Trim(Me.TextBox1.Text)

        If lsQuery = "" Then Exit Sub

        Dim lrRecordset = Me.mrMindfulDB.ProcessQuery(lsQuery, "Cypher")

#Region "Display the Results"
        Me.TextBox2.Text = ""

        Dim liInd = 0
        For Each lsColumnName In lrRecordset.Columns
            liInd += 1
            Me.TextBox2.Text &= " " & lsColumnName & " "
            If liInd < lrRecordset.Columns.Count Then Me.TextBox2.Text &= ","
        Next
        Me.TextBox2.Text &= vbCrLf & "=======================================" & vbCrLf

        For Each lrFact In lrRecordset.Facts

            Me.TextBox2.Text &= lrFact.EnumerateAsBracketedFact(True) & vbCrLf
        Next
#End Region

    End Sub

End Class
