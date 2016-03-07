'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports System
Imports System.Data.SqlClient
'-------------------------------------------------------------------------------------------'
' Inicio de clase "REliminar_VDefecto"
'-------------------------------------------------------------------------------------------'
Partial Class rEliminar_VDefecto

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("------------")
            loComandoSeleccionar.AppendLine("------------")
            loComandoSeleccionar.AppendLine("-- Script---")
            loComandoSeleccionar.AppendLine("------------")
            loComandoSeleccionar.AppendLine("------------")
            loComandoSeleccionar.AppendLine("   USE Factory_Administrativo_Pruebas;   ")
            loComandoSeleccionar.AppendLine("   ")
            loComandoSeleccionar.AppendLine("   SELECT   ")
            loComandoSeleccionar.AppendLine("   			cast(OBJECT_NAME (tablas.object_id) as varchar) as tabla,  ")
            loComandoSeleccionar.AppendLine("   			cast(columnas.name as varchar)as columna   ")
            loComandoSeleccionar.AppendLine("     ")
            loComandoSeleccionar.AppendLine("   			  ")
            loComandoSeleccionar.AppendLine("   FROM  ")
            loComandoSeleccionar.AppendLine("   			sys.objects as tablas  ")
            loComandoSeleccionar.AppendLine("   			join sys.columns as columnas on tablas.object_id = columnas.object_id  ")
            loComandoSeleccionar.AppendLine("   			join sys.types as tipos on tipos.user_type_id = columnas.user_type_id  ")
            loComandoSeleccionar.AppendLine("   WHERE  ")
            loComandoSeleccionar.AppendLine("   			 tablas.type = 'U'  ")
            loComandoSeleccionar.AppendLine("   	  ")
            loComandoSeleccionar.AppendLine("   order by tablas.object_id;  ")
            loComandoSeleccionar.AppendLine("     ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("------------")
            loComandoSeleccionar.AppendLine("------------")
            loComandoSeleccionar.AppendLine("-- Script---")
            loComandoSeleccionar.AppendLine("------------")
            loComandoSeleccionar.AppendLine("------------")



            Dim sqlConnection1 As New SqlConnection("Data Source= GALILEO;Initial Catalog= Factory_Administrativo_Pruebas;User Id= factory;Password = factory")
            Dim cmd As New SqlCommand
            Dim reader As SqlDataReader
            Dim recordData As String = ""
            Dim recordCount As Integer = 0
            Dim i As Integer = 0


            cmd.CommandText = loComandoSeleccionar.ToString

            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader()

            Dim script As String
            script = ""
            While reader.Read()



                loComandoSeleccionar.AppendLine("    EXEC sp_unbindefault '" & reader(0).ToString() & "." & reader(1).ToString() & "'     ")


                recordData &= ControlChars.CrLf
                recordCount += 1
            End While

            sqlConnection1.Close()



            Response.Clear()
            Response.Write("<html><body><pre>" & vbNewLine)
            Response.Write(loComandoSeleccionar.ToString)
            Response.Write("</pre></body></html>")
            Response.Flush()
            Response.End()
        Catch ex As Exception
            Response.Write(ex.Message)

        End Try
    End Sub

End Class




'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' AAP: 30/06/09: Programacion inicial
'-------------------------------------------------------------------------------------------'

