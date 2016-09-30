'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_fConceptos_Movimientos_ConCC"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fConceptos_Movimientos_ConCC
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT	Conceptos.Cod_Con		AS Codigo_Concepto, ")
            loComandoSeleccionar.AppendLine("        Conceptos.Nom_Con		AS Concepto, ")
            loComandoSeleccionar.AppendLine("        ISNULL(CAST(Conceptos.Contable AS XML).value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')		AS CC_Pagina1,")
            loComandoSeleccionar.AppendLine("		ISNULL((SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = ISNULL(CAST(Conceptos.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[1]', 'varchar(12)'),'')),'')								AS Nom_CC_Pagina1,")
            loComandoSeleccionar.AppendLine("		ISNULL(CAST(Conceptos.Contable AS XML) .value('(/contable/ficha/cue_con)[2]', 'varchar(12)'),'')	AS CC_Pagina2,")
            loComandoSeleccionar.AppendLine("		ISNULL((SELECT Nom_Cue FROM Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("				WHERE Cod_Cue = ISNULL(CAST(Conceptos.Contable AS XML)")
            loComandoSeleccionar.AppendLine("				.value('(/contable/ficha/cue_con)[2]', 'varchar(12)'),'')),'')								AS Nom_CC_Pagina2")
            loComandoSeleccionar.AppendLine("FROM	Conceptos ")
            loComandoSeleccionar.AppendLine("ORDER BY Conceptos.Cod_Con ")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '-------------------------------------------------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------'
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------'
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fConceptos_Movimientos_ConCC", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_fConceptos_Movimientos_ConCC.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 31/10/14: Codigo inicial															    '
'-------------------------------------------------------------------------------------------'
' JJD: 18/12/14: Inclusion del Len de la Cuenta Contable                                    '
'-------------------------------------------------------------------------------------------'

