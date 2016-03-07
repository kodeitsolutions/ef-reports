Imports System.Data
Partial Class rLCierre_CuentasCobrar

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden


        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro11Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro12Desde As String = cusAplicacion.goReportes.paParametrosIniciales(12)
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro14Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(14), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro14Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(14), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Clientes.Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 		Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 		Cuentas_Cobrar.Documento,")
            loComandoSeleccionar.AppendLine(" 		Cuentas_Cobrar.Cod_Tip,")
            loComandoSeleccionar.AppendLine(" 		Cuentas_Cobrar.Fec_Ini AS Fecha_Emision,")
            loComandoSeleccionar.AppendLine(" 		Cuentas_Cobrar.Fec_Fin AS Fecha_Vencimiento,")
            loComandoSeleccionar.AppendLine(" 		CONVERT(DATETIME," & lcParametro1Hasta & ") AS Fecha_Actual,")
            loComandoSeleccionar.AppendLine(" 		DATEDIFF(DAY,Cuentas_Cobrar.Fec_Ini,Cuentas_Cobrar.Fec_Fin) AS Dias_Reales,")
            loComandoSeleccionar.AppendLine(" 		DATEDIFF(DAY,Cuentas_Cobrar.Fec_Fin," & lcParametro1Hasta & ") AS Dias_Vencidos,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Net *(-1) ")
            loComandoSeleccionar.AppendLine(" 			ELSE Cuentas_Cobrar.Mon_Net ")
            loComandoSeleccionar.AppendLine(" 		END As Mon_Net,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal *(-1) ")
            loComandoSeleccionar.AppendLine(" 			ELSe Cuentas_Cobrar.Mon_Sal ")
            loComandoSeleccionar.AppendLine(" 		END As Mon_Sal,")
            loComandoSeleccionar.AppendLine(" 		Cuentas_Cobrar.Comentario")
            loComandoSeleccionar.AppendLine(" FROM	Cuentas_Cobrar,")
            loComandoSeleccionar.AppendLine(" 		Clientes,")
            loComandoSeleccionar.AppendLine(" 		Vendedores,")
            loComandoSeleccionar.AppendLine(" 		Transportes,")
            loComandoSeleccionar.AppendLine(" 		Monedas")
            loComandoSeleccionar.AppendLine(" WHERE	Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Tra = Transportes.Cod_Tra")
            loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Mon = Monedas.Cod_Mon")
            loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Ven      BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Clientes.Cod_Zon			BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Status       IN ( " & lcParametro6Desde & " )")
            loComandoSeleccionar.AppendLine("       AND Clientes.Cod_Tip			BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("       AND Clientes.Cod_Cla			BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Tra      BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Mon      BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("       AND (( " & lcParametro11Desde & "  = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine(" 			OR ( " & lcParametro11Desde & "  <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 OR Cuentas_Cobrar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro13Desde & " AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Rev      BETWEEN " & lcParametro14Desde & " AND " & lcParametro14Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine(" ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLCierre_CuentasCobrar", laDatosReporte)

            If lcParametro12Desde.ToString = "Si" Then
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 1
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line3"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line3"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
            Else
                loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Height = 0
                loObjetoReporte.ReportDefinition.ReportObjects("text1").Top = 0
                loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Top = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line9"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line9"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line9"), CrystalDecisions.CrystalReports.Engine.LineObject).Top = 0
                CType(loObjetoReporte.ReportDefinition.ReportObjects("Line9"), CrystalDecisions.CrystalReports.Engine.LineObject).Bottom = 0
                loObjetoReporte.ReportDefinition.Sections(3).Height = 200
                loObjetoReporte.ReportDefinition.Sections("DetailSection2").Height = 10
            End If

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLCierre_CuentasCobrar.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message & " - Pila: " & loExcepcion.StackTrace, _
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
' Douglas Cortez: 10/06/2010: Programacion inicial
'-------------------------------------------------------------------------------------------'
' MAT:  18/02/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
