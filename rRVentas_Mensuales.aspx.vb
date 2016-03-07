'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRVentas_Mensuales"
'-------------------------------------------------------------------------------------------'
Partial Class rRVentas_Mensuales
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosIniciales(6)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("   SELECT ")
            loComandoSeleccionar.AppendLine("           CAST(DATEPART(YEAR, Fec_Ini) AS INT) AS Año_CXC, ")
            loComandoSeleccionar.AppendLine("           CAST(DATEPART(MONTH, Fec_Ini) AS INT) AS Mes_CXC, ")
            'loComandoSeleccionar.AppendLine("           SUM((mon_Bas1 + mon_exe)) AS Mon_Bas1_CXC, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Bru) AS Mon_Bas1_CXC, ")
            loComandoSeleccionar.AppendLine("           SUM(mon_imp1) AS Mon_Imp1_CXC, ")
            loComandoSeleccionar.AppendLine("           SUM((mon_bas1/30))AS Mon_Bas1_CXCP, ")
            loComandoSeleccionar.AppendLine("           SUM((mon_imp1/30)) AS Mon_Imp1_CXCP  ")
            loComandoSeleccionar.AppendLine("   INTO	#temporal ")
            loComandoSeleccionar.AppendLine("   FROM	Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("   WHERE	Cod_Tip = 'Fact' ")
            loComandoSeleccionar.AppendLine("           AND Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("           AND Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND cod_cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND cod_ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND cod_mon BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            'loComandoSeleccionar.AppendLine("           AND Status IN (" & lcParametro4Desde & ")")

            If lcParametro6Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro5Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro5Desde)
            End If

            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
			loComandoSeleccionar.AppendLine("			Group BY  DATEPART(YEAR, Fec_Ini), DATEPART(MONTH, Fec_Ini)")
			loComandoSeleccionar.AppendLine("     ")		

            loComandoSeleccionar.AppendLine("   SELECT  ")
            loComandoSeleccionar.AppendLine("           DATEPART(YEAR, Fec_Ini) AS Año_DC, ")
            loComandoSeleccionar.AppendLine("           DATEPART(MONTH, Fec_Ini) AS Mes_DC, ")
            'loComandoSeleccionar.AppendLine("           SUM((mon_Bas1 + mon_exe)) AS Mon_Bas1_DC, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Bru) AS Mon_Bas1_DC, ")
            loComandoSeleccionar.AppendLine("           SUM(mon_imp1) AS Mon_Imp1_DC, ")
            loComandoSeleccionar.AppendLine("           SUM((mon_bas1/30))AS Mon_Bas1_DCP, ")
            loComandoSeleccionar.AppendLine("           SUM((mon_imp1/30)) AS Mon_Imp1_DCP  ")
            loComandoSeleccionar.AppendLine("   INTO	#temporal2 ")
            loComandoSeleccionar.AppendLine("   FROM	Devoluciones_Clientes  ")
            loComandoSeleccionar.AppendLine("   WHERE	Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("           AND Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND cod_cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND cod_ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND cod_mon BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            If lcParametro6Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev between " & lcParametro5Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cod_Rev NOT between " & lcParametro5Desde)
            End If

            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            'loComandoSeleccionar.AppendLine("           AND Status IN (" & lcParametro4Desde & ")")
			loComandoSeleccionar.AppendLine("			Group BY  DATEPART(YEAR, Fec_Ini), DATEPART(MONTH, Fec_Ini)")
			loComandoSeleccionar.AppendLine("     ")		

            loComandoSeleccionar.AppendLine("   SELECT  ")
            loComandoSeleccionar.AppendLine("           #temporal.Año_CXC, ")
            loComandoSeleccionar.AppendLine("           #temporal.Mes_CXC, ")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(#temporal.Mon_Bas1_CXC), 0) as Mon_Bas1_CXC, ")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(#temporal.Mon_Imp1_CXC), 0) as Mon_Imp1_CXC, ")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(#temporal2.Mon_Bas1_DC), 0) as Mon_Bas1_DC, ")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(#temporal2.Mon_Imp1_DC), 0) as Mon_Imp1_DC, ")
            loComandoSeleccionar.AppendLine("           sum(ISNULL(#temporal.Mon_Bas1_CXC, 0) - ISNULL(#temporal2.Mon_Bas1_DC, 0))  as Mon_Net, ")
            loComandoSeleccionar.AppendLine("           sum(ISNULL(#temporal.Mon_Imp1_CXC, 0) - ISNULL(#temporal2.Mon_Imp1_DC, 0))  as Imp_Net, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(Mon_Bas1_CXCP), 0) AS Mon_Bas1_CXCP, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(Mon_Imp1_CXCP), 0) AS Mon_Imp1_CXCP, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(Mon_Bas1_DCP), 0) AS Mon_Bas1_DCP, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(Mon_Imp1_DCP), 0) AS Mon_Imp1_DCP, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(Mon_Bas1_CXCP, 0) - ISNULL(Mon_Bas1_DCP, 0))  AS Mon_Netp, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(Mon_Imp1_CXCP, 0) - ISNULL(Mon_Imp1_DCP, 0))  AS Imp_Netp ")
            loComandoSeleccionar.AppendLine("   FROM	#Temporal #Temporal ")
            loComandoSeleccionar.AppendLine("   Full JOIN  #temporal2 as #temporal2 on ((#temporal.Año_CXC = #temporal2.Año_DC) AND (#temporal.Mes_CXC = #temporal2.Mes_DC)) ")
            loComandoSeleccionar.AppendLine("   GROUP BY #temporal.Año_CXC, #temporal.Mes_CXC ")
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRVentas_Mensuales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRVentas_Mensuales.ReportSource = loObjetoReporte

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
' CMS: 19/05/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  22/08/09: Se agrego el filtro tipo de revisión y la verificacion de registro
'-------------------------------------------------------------------------------------------'
