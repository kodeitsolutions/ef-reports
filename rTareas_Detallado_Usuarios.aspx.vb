'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTareas_Detallado_Usuarios"
'-------------------------------------------------------------------------------------------'
Partial Class rTareas_Detallado_Usuarios

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 			Factory_Global.dbo.Usuarios.Cod_Usu, ")
            loComandoSeleccionar.AppendLine(" 			Factory_Global.dbo.Usuarios.Nom_Usu, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Cod_Tar, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Nom_Tar, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Por_Eje, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Tipo, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Prioridad, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Etapa, ")
            loComandoSeleccionar.AppendLine(" 			Tareas.Comentario")
            loComandoSeleccionar.AppendLine(" FROM		Tareas,Factory_Global.dbo.Usuarios")
            loComandoSeleccionar.AppendLine(" WHERE		Tareas.Origen	=	'Usuarios'")
            loComandoSeleccionar.AppendLine("           AND	(Factory_Global.dbo.Usuarios.Cod_Usu COLLATE Modern_Spanish_CI_AS = Tareas.Cod_Reg COLLATE  Modern_Spanish_CI_AS)")
            loComandoSeleccionar.AppendLine("			AND  Factory_Global.dbo.Usuarios.Cod_Usu      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		    AND Factory_Global.dbo.Usuarios.Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))
            loComandoSeleccionar.AppendLine("		    AND Factory_Global.dbo.Usuarios.Status   IN (" & lcParametro1Desde & ")")


            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            ' Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTareas_Detallado_Usuarios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrTareas_Detallado_Usuarios.ReportSource = loObjetoReporte

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
' MAT: 11/10/11: Codigo inicial.
'-------------------------------------------------------------------------------------------'