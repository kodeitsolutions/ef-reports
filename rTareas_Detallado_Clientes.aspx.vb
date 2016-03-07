'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTareas_Detallado_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class rTareas_Detallado_Clientes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 			Clientes.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Cod_Tar, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Nom_Tar, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Por_Eje, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Tipo, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Prioridad, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Etapa, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Comentario")
            loComandoSeleccionar.AppendLine(" FROM		Tareas")
            loComandoSeleccionar.AppendLine(" JOIN		Clientes ON Clientes.Cod_Cli = Tareas.Cod_Reg")
            loComandoSeleccionar.AppendLine("				AND Clientes.Cod_Cli between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" WHERE		Tareas.Origen	=	'Clientes'")
            loComandoSeleccionar.AppendLine("			AND Tareas.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Tareas.Status IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine(" 			AND Tareas.Prioridad IN (" & lcParametro3Desde & ")")
			loComandoSeleccionar.AppendLine(" 			AND Tareas.Importancia IN (" & lcParametro4Desde & ")")            
			loComandoSeleccionar.AppendLine("			AND Tareas.Etapa IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Tareas.Cod_Usu between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Tareas.Cod_Usu between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro7Hasta) 
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTareas_Detallado_Clientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrTareas_Detallado_Clientes.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' MAT: 18/03/11: Codigo inicial.
'-------------------------------------------------------------------------------------------'