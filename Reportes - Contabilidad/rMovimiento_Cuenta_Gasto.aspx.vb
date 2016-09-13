'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMovimiento_Cuenta_Gasto"
'-------------------------------------------------------------------------------------------'
Partial Class rMovimiento_Cuenta_Gasto
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim loComandoSeleccionar As New StringBuilder()
            loComandoSeleccionar.AppendLine("SELECT	Cuentas_Gastos.Cod_Gas, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Gastos.Nom_Gas, ")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Comprobantes.mon_deb) AS Debe, ")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Comprobantes.mon_hab) AS Haber, ")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Comprobantes.mon_deb - Renglones_Comprobantes.mon_hab) AS Saldo ")
            loComandoSeleccionar.AppendLine("FROM   comprobantes ")
            loComandoSeleccionar.AppendLine("	JOIN		Renglones_Comprobantes ON Renglones_Comprobantes.documento = Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("		AND		Renglones_Comprobantes.Adicional = Comprobantes.Adicional")
            loComandoSeleccionar.AppendLine("		AND		Comprobantes.Tipo = 'Diario' ")
            loComandoSeleccionar.AppendLine("		AND		Comprobantes.status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("		AND		Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("	JOIN		Cuentas_Contables ON Cuentas_Contables.Cod_cue = Renglones_Comprobantes.Cod_cue ")
            loComandoSeleccionar.AppendLine("	JOIN		Cuentas_Gastos ON Cuentas_Gastos.cod_gas = Cuentas_Contables.cod_gas ")
            loComandoSeleccionar.AppendLine("	WHERE		Comprobantes.Documento          BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND		Comprobantes.Fecha              BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND		Comprobantes.Cod_Suc            BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND		Comprobantes.Cod_Rev            BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND		Renglones_Comprobantes.Fec_Ori  BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND		Cuentas_Contables.Cod_cue  BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND		Cuentas_Gastos.Cod_gas  BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("	        AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Cuentas_Gastos.Nom_Gas,Cuentas_Gastos.Cod_Gas ")
            loComandoSeleccionar.AppendLine("ORDER BY Cuentas_Gastos.Nom_Gas ")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMovimiento_Cuenta_Gasto", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMovimiento_Cuenta_Gasto.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' EAG: 08/08/15: Codigo inicial																'
'-------------------------------------------------------------------------------------------'

