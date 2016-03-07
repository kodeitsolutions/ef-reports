'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEstimacion_Pagos"
'-------------------------------------------------------------------------------------------'
Partial Class rEstimacion_Pagos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Pagar.Fec_Fin,")
            loComandoSeleccionar.AppendLine(" 			CASE ")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(DAY, Cuentas_Pagar.Fec_Fin) >= 1 AND DATEPART(DAY, Cuentas_Pagar.Fec_Fin) <= 7 THEN Cuentas_Pagar.Mon_Sal")
            loComandoSeleccionar.AppendLine(" 					ELSE 0")
            loComandoSeleccionar.AppendLine(" 			END AS Semana_1, CASE ")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(DAY, Cuentas_Pagar.Fec_Fin) >= 8 AND DATEPART(DAY, Cuentas_Pagar.Fec_Fin) <= 15 THEN Cuentas_Pagar.Mon_Sal")
            loComandoSeleccionar.AppendLine(" 					ELSE 0")
            loComandoSeleccionar.AppendLine(" 			END AS Semana_2,")
            loComandoSeleccionar.AppendLine(" 			CASE ")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(DAY, Cuentas_Pagar.Fec_Fin) >= 16 AND DATEPART(DAY, Cuentas_Pagar.Fec_Fin) <= 22 THEN Cuentas_Pagar.Mon_Sal")
            loComandoSeleccionar.AppendLine(" 					ELSE 0")
            loComandoSeleccionar.AppendLine(" 			END AS Semana_3,")
            loComandoSeleccionar.AppendLine(" 			CASE ")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(DAY, Cuentas_Pagar.Fec_Fin) >= 23 AND DATEPART(DAY, Cuentas_Pagar.Fec_Fin) <= 31 THEN Cuentas_Pagar.Mon_Sal")
            loComandoSeleccionar.AppendLine(" 					ELSE 0")
            loComandoSeleccionar.AppendLine(" 			END AS Semana_4")
            loComandoSeleccionar.AppendLine(" INTO #Temp")
            loComandoSeleccionar.AppendLine(" FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")

            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Fec_Fin  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Pro  Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Ven  Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tip  Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Status   IN  (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Mon  Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Proveedores.Cod_Zon        Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Proveedores.Cod_Pro        Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Rev  Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Rev  Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY Fec_Fin DESC")
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR, CONVERT(nchar(30), Fec_Fin,112)) AS Año,")
            loComandoSeleccionar.AppendLine(" 			CASE ")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 1 THEN 'Enero'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 2 THEN 'Febrero'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 3 THEN 'Marzo'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 4 THEN 'Abril'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 5 THEN 'Mayo'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 6 THEN 'Junio'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 7 THEN 'Julio'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 8 THEN 'Agosto'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 9 THEN 'Septiembre'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 10 THEN 'Octubre'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 11 THEN 'Noviembre'")
            loComandoSeleccionar.AppendLine(" 					WHEN DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) = 12 THEN 'Diciembre'")
            loComandoSeleccionar.AppendLine(" 			END AS Mes,")
            loComandoSeleccionar.AppendLine(" 			SUM(Semana_1) AS Semana_1,")
            loComandoSeleccionar.AppendLine(" 			SUM(Semana_2) AS Semana_2,")
            loComandoSeleccionar.AppendLine(" 			SUM(Semana_3) AS Semana_3,")
            loComandoSeleccionar.AppendLine(" 			SUM(Semana_4) AS Semana_4,")
            loComandoSeleccionar.AppendLine(" 			SUM(Semana_1)+SUM(Semana_2)+SUM(Semana_3)+SUM(Semana_4) AS Total")
            loComandoSeleccionar.AppendLine(" FROM	#Temp")
            loComandoSeleccionar.AppendLine(" GROUP BY DATEPART(YEAR, CONVERT(nchar(30), Fec_Fin,112)) , DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112))")
            loComandoSeleccionar.AppendLine(" ORDER BY DATEPART(YEAR, CONVERT(nchar(30), Fec_Fin,112))  DESC, DATEPART(MONTH, CONVERT(nchar(30), Fec_Fin,112)) ASC")

          

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEstimacion_Pagos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrEstimacion_Pagos.ReportSource = loObjetoReporte

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
' CMS: 27/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'