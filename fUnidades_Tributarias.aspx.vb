'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fUnidades_Tributarias"
'-------------------------------------------------------------------------------------------'
Partial Class fUnidades_Tributarias
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Unidades_Tributarias.cod_uni, ")
            loComandoSeleccionar.AppendLine("		(CASE Unidades_Tributarias.status ")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO' ")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO' ")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO' ")
            loComandoSeleccionar.AppendLine("		END) AS status, ")
            loComandoSeleccionar.AppendLine("		Unidades_Tributarias.gaceta, ")
            loComandoSeleccionar.AppendLine("		Unidades_Tributarias.Fec_Gac, ")
            loComandoSeleccionar.AppendLine("		Unidades_Tributarias.fec_ini, ")
            loComandoSeleccionar.AppendLine("		Unidades_Tributarias.fec_fin, ")
            loComandoSeleccionar.AppendLine("		Unidades_Tributarias.monto, ")
            loComandoSeleccionar.AppendLine("		Unidades_Tributarias.comentario, ")
            loComandoSeleccionar.AppendLine("		Unidades_Tributarias.factor ")
            loComandoSeleccionar.AppendLine("FROM Unidades_Tributarias ")
            loComandoSeleccionar.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fUnidades_Tributarias", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfUnidades_Tributarias.ReportSource = loObjetoReporte

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
' EAG: 02/09/2015 : Codigo inicial
'-------------------------------------------------------------------------------------------'
