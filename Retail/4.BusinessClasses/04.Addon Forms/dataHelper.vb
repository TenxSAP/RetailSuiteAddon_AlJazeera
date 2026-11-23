Imports System.Configuration
Imports Sap.Data.Hana
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Text


Public Class dataHelper
    Dim con As HanaConnection
    'Dim logger As Logger = LogManager.GetCurrentClassLogger()

    Public Sub New()
        'con = New HanaConnection(ConfigurationManager.AppSettings("luxryCon"))

        'Dim strCon As String = "Driver={HDBODBC32};UID=SYSTEM;PWD=e++3QsJZ-gLLtGU?;SERVERNODE=hana-vm-49.cloudiax.com:30013;Max Pool Size=200;DATABASENAME=HV49C21227P01;CS={GALFAR_T03}"
        'Dim strCon1 As String = ConfigurationManager.AppSettings("CS")
        Dim strCon1 As String = ""

        'strCon1 = objMain.objPriceUpdateService.ConStr

        'If strCon1 = "" Then
        '    'Error emssage
        '    objMain.objApplication.StatusBar.SetText("Please configure the connection string", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        'End If
        'con = New HanaConnection(strCon1)
    End Sub

    Public Function getDataSet(strSql As String) As DataSet
        con = New HanaConnection(objMain.ObjPriceChangeRequest.ConStr)
        Dim cmd As New HanaCommand(strSql, con)
        Dim da As New HanaDataAdapter(cmd)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            'logger.Error(ex.Message)
            Throw ex
        End Try
        Return ds

    End Function

    Public Function ExecuteScalar(strSql As String) As Object
        con = New HanaConnection(objMain.ObjPriceChangeRequest.ConStr)
        Dim cmd As New HanaCommand(strSql, con)

        Try
            con.Open()
            cmd.ExecuteScalar()
            con.Close()
        Catch ex As Exception
            'logger.Error(ex.Message)
            Throw ex
        End Try

    End Function

    Public Sub ExecuteNonQuery(strSql As String)
        con = New HanaConnection(objMain.ObjPriceChangeRequest.ConStr)
        Dim cmd As New HanaCommand(strSql, con)

        Try
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
        Catch ex As Exception
            'logger.Error(ex.Message)
            Throw ex
        End Try
    End Sub

End Class


