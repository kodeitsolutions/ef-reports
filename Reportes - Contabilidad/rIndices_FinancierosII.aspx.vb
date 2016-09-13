Imports System.Data
Partial Class rIndices_FinancierosII
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

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
            loComandoSeleccionar.AppendLine("		REPLACE(COALESCE(Renglones_IFinancieros.Cuentas_Contables,''),';',';'+CHAR(10))	AS Cuentas_Contables, ")
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
            loComandoSeleccionar.AppendLine("WHERE      Indices_Financieros.cod_ind BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Indices_Financieros.status IN(" & lcParametro1Desde & ")")
            If lcParametro2Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   AND Indices_Financieros.grupo = " & lcParametro2Desde)
            End If

            If lcParametro3Desde <> "''" Then
                loComandoSeleccionar.AppendLine("   AND Indices_Financieros.clase = " & lcParametro3Desde)
            End If
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rIndices_FinancierosII", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrIndices_FinancierosII.ReportSource = loObjetoReporte


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
' EAG:  24/09/15 : Codigo inicial
'-------------------------------------------------------------------------------------------' 
