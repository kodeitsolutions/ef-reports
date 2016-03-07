'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fRequerimientos_Tesoreria"
'-------------------------------------------------------------------------------------------'
Partial Class fRequerimientos_Tesoreria
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            loComandoSeleccionar.AppendLine("SELECT	Requerimientos.cod_req, ")
            loComandoSeleccionar.AppendLine("		(CASE Requerimientos.status  ")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO'  ")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO'  ")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO'  ")
            loComandoSeleccionar.AppendLine("		END) AS status,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.nom_req,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.tip_pro, ")
            loComandoSeleccionar.AppendLine("		Requerimientos.tip_com, ")
            loComandoSeleccionar.AppendLine("		Requerimientos.grupo,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.cod_com,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.comentario,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.Sistema,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.modulo,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.seccion,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.Opcion,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.Usuarios_Permitidos,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.Grupos_Permitidos,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.Registros_Eximidos,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.Usuarios_Eximidos,  ")
            loComandoSeleccionar.AppendLine("		Requerimientos.Grupos_Eximidos  ")
            loComandoSeleccionar.AppendLine("FROM Requerimientos ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fRequerimientos_Tesoreria", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfRequerimientos_Tesoreria.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' EAG: 03/09/2015 : Codigo inicial
'-------------------------------------------------------------------------------------------'
