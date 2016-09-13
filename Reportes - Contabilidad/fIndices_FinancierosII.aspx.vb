Imports System.Data
Partial Class fIndices_FinancierosII

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Indices_Financieros.cod_ind								AS cod_ind, ")
            loComandoSeleccionar.AppendLine("		Indices_Financieros.nom_ind								AS nom_ind, ")
            loComandoSeleccionar.AppendLine("       CASE Indices_Financieros.status ")
            loComandoSeleccionar.AppendLine("           WHEN 'A' THEN 'ACTIVO' ")
            loComandoSeleccionar.AppendLine("           WHEN 'I' THEN 'INACTIVO' ")
            loComandoSeleccionar.AppendLine("           WHEN 'S' THEN 'SUSPENDIDO' ")
            loComandoSeleccionar.AppendLine("       End                                                     AS status, ")
            loComandoSeleccionar.AppendLine("		Indices_Financieros.definicion							AS definicion, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_IFinancieros.cod_cue,'')				AS Cuenta, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_IFinancieros.nom_cue,'')				AS Nombre, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_IFinancieros.Cuentas_Contables,'')	AS Cuentas_Contables, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_IFinancieros.Centros_Costos,'')		AS Centros_Costos, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_IFinancieros.Cuentas_Gastos,'')		AS Cuentas_Gastos, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_IFinancieros.Auxiliares,'')			AS Auxiliares, ")
            loComandoSeleccionar.AppendLine("		Renglones_IFinancieros.fec_ini							AS fec_ini, ")
            loComandoSeleccionar.AppendLine("		Renglones_IFinancieros.fec_fin							AS fec_fin, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_IFinancieros.Nivel1,'')				AS Nivel1, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_IFinancieros.Nivel2,'')				AS Nivel2, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_IFinancieros.Nivel3,'')				AS Nivel3 ")
            loComandoSeleccionar.AppendLine("FROM Indices_Financieros ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Renglones_IFinancieros  ")
            loComandoSeleccionar.AppendLine("			ON	Renglones_IFinancieros.tab_ori		=	Indices_Financieros.tab_ori ")
            loComandoSeleccionar.AppendLine("			AND	Renglones_IFinancieros.doc_ori		=	Indices_Financieros.doc_ori ")
            loComandoSeleccionar.AppendLine("			AND	Renglones_IFinancieros.cla_ori		=	Indices_Financieros.cla_ori ")
            loComandoSeleccionar.AppendLine("			AND	Renglones_IFinancieros.ren_ori		=	Indices_Financieros.ren_ori ")
            loComandoSeleccionar.AppendLine("			AND	Renglones_IFinancieros.tip_ind		=	Indices_Financieros.tip_ind ")
            loComandoSeleccionar.AppendLine("			AND	Renglones_IFinancieros.cod_ind		=	Indices_Financieros.cod_ind ")
            loComandoSeleccionar.AppendLine("			AND	Renglones_IFinancieros.Adicional	=	Indices_Financieros.Adicional ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fModelos_Modelo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfIndices_FinancierosII.ReportSource = loObjetoReporte

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
' EAG: 24/09/15: Codigo inicial
'-------------------------------------------------------------------------------------------'


