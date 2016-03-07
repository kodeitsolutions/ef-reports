'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fImpuestos_Nacionales"
'-------------------------------------------------------------------------------------------'
Partial Class fImpuestos_Nacionales
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Impuestos.cod_imp, ")
            loComandoSeleccionar.AppendLine(" 		(CASE Impuestos.status ")
            loComandoSeleccionar.AppendLine(" 			WHEN 'A' THEN 'ACTIVO' ")
            loComandoSeleccionar.AppendLine(" 			WHEN 'I' THEN 'INACTIVO' ")
            loComandoSeleccionar.AppendLine(" 			WHEN 'S' THEN 'SUSPENDIDO' ")
            loComandoSeleccionar.AppendLine(" 		END) AS status, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.nom_imp , ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.tip_imp, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp1, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp1, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp2, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp2, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp3, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp3, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp4, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp4, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp5, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp5, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp6, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp6, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp7, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp7, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp8, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp8, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp9, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp9, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.por_imp10, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.mon_imp10, ")
            loComandoSeleccionar.AppendLine(" 		Impuestos.formula ")
            loComandoSeleccionar.AppendLine(" FROM Impuestos ")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fImpuestos_Nacionales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfImpuestos_Nacionales.ReportSource = loObjetoReporte

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
