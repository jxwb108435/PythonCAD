'(C) Copyright 2008 by Autodesk, Inc.
'
'
'By using this code, you are agreeing to the terms
'and conditions of the License Agreement that appeared
'and was accepted upon download or installation
'(or in connection with the download or installation)
'of the Autodesk software in which this code is included.
'All permissions on use of this code are as set forth
'in such License Agreement provided that the above copyright
'notice appears in all authorized copies and that both that
'copyright notice and the limited warranty and
'restricted rights notice below appear in all supporting
'documentation.
'
'AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
'AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
'MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
'DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
'UNINTERRUPTED OR ERROR FREE.
'
'Use, duplication, or disclosure by the U.S. Government is subject to
'restrictions set forth in FAR 52.227-19 (Commercial Computer
'Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
'(Rights in Technical Data and Computer Software), as applicable.

Option Explicit On
Option Strict Off




Imports DBTransactionManager = Autodesk.AutoCAD.DatabaseServices.TransactionManager

Public MustInherit Class SATemplate
    ' --------------------------------------------------------------------------
    ' Returns logical names used by this script
    Public Sub GetLogicalNames()
        Dim trans As Transaction = Nothing
        Dim corridorState As CorridorState = Nothing

        Try

            ' Start transaction
            trans = StartTransaction()
            ' Get the corridor stateobject
            corridorState = CivilApplication.ActiveDocument.CorridorState

            ' Retrieve parameter buckets from the corridor state
            GetLogicalNamesImplement(corridorState)

            trans.Commit()

        Catch e As System.Exception
            Utilities.RecordError(corridorState, e)
            If trans IsNot Nothing Then trans.Abort()
        Finally
            If trans IsNot Nothing Then trans.Dispose()
        End Try

    End Sub

    ' --------------------------------------------------------------------------
    ' Returns input parameters required by this script
    Public Sub GetInputParameters()

        Dim trans As Transaction = Nothing
        Dim corridorState As CorridorState = Nothing

        Try
            trans = StartTransaction()
            corridorState = CivilApplication.ActiveDocument.CorridorState

            GetInputParametersImplement(corridorState)

            trans.Commit()
        Catch e As System.Exception
            Utilities.RecordError(corridorState, e)
            If trans IsNot Nothing Then trans.Abort()
        Finally
            If trans IsNot Nothing Then trans.Dispose()
        End Try

    End Sub

    ' --------------------------------------------------------------------------
    ' Returns output parameters returned by this script
    Public Sub GetOutputParameters()

        Dim trans As Transaction = Nothing
        Dim corridorState As CorridorState = Nothing

        Try
            ' Start transaction
            trans = StartTransaction()
            corridorState = CivilApplication.ActiveDocument.CorridorState

            GetOutputParametersImplement(corridorState)

            trans.Commit()
        Catch e As System.Exception
            Utilities.RecordError(corridorState, e)
            If trans IsNot Nothing Then trans.Abort()
        Finally
            If trans IsNot Nothing Then trans.Dispose()
        End Try

    End Sub

    Public Sub Draw()
        Dim corridorState As CorridorState = Nothing

        Try
            corridorState = CivilApplication.ActiveDocument.CorridorState

            DrawImplement(corridorState)
        Catch e As System.Exception
            Utilities.RecordError(corridorState, e)
        End Try
    End Sub



    Protected Overridable Sub GetLogicalNamesImplement(ByVal corridorState As CorridorState)
        'default do nothing
    End Sub
    Protected Overridable Sub GetInputParametersImplement(ByVal corridorState As CorridorState)
        'default do nothing
    End Sub
    Protected Overridable Sub GetOutputParametersImplement(ByVal corridorState As CorridorState)
        'default do nothing
    End Sub


    Protected MustOverride Sub DrawImplement(ByVal corridorState As CorridorState)


    Protected Function StartTransaction() As Transaction
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim tm As DBTransactionManager = db.TransactionManager
        Dim trans As Transaction = tm.StartTransaction()
        Return trans
    End Function

    '' *************************************************************************
    '' *************************************************************************
    '' *************************************************************************
    ''          Name: GetCorridorState
    ''
    ''   Description:  return the corridor state .NET object.
    ''                It has been opened for writing.
    ''                Attention: it method should be called only once under a transaction opened by transactionManager
    ''
    'Protected Function GetCorridorState(ByVal trans As Transaction) As CorridorState
    '    'If trans Is Nothing Then
    '    '    Throw New ArgumentNullException("trans")
    '    'End If
    '    'Dim oId As ObjectId
    '    'oId = CorridorApplication.Application.ActiveDocument.CorridorStateId
    '    'Return trans.GetObject(oId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, False)
    '    Return CorridorApplication.Application.ActiveDocument.CorridorState
    'End Function

    Protected Sub New()
        If Not Codes.CodesStructureFilled Then
            FillCodeStructure()
        End If
    End Sub
End Class
