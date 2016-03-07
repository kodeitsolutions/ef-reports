'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRVentas_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class rRVentas_Clientes
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT ")
            loComandoSeleccionar.AppendLine("       Cod_Cli, ")
            loComandoSeleccionar.AppendLine("       SUM(Mon_Bru) AS Mon_Bas1_CXC, ")
            loComandoSeleccionar.AppendLine("       SUM(Mon_Imp1) AS Mon_Imp1_CXC, ")
            loComandoSeleccionar.AppendLine("       SUM((Mon_Bru/30)) AS Mon_Bas1_CXCP, ")
            loComandoSeleccionar.AppendLine("       SUM((Mon_Imp1/30)) AS Mon_Imp1_CXCP  ")
            loComandoSeleccionar.AppendLine("INTO	#Temporal ")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("WHERE	Cod_Tip = 'Fact' ")
            loComandoSeleccionar.AppendLine("       AND Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("       AND Cod_Suc BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Fec_Ini BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cod_Cli BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cod_Ven BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("       AND Cod_Mon BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Cod_Rev BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Cod_Cli")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT  ")
            loComandoSeleccionar.AppendLine("       Cod_Cli, ")
            loComandoSeleccionar.AppendLine("       SUM(Mon_Bru) AS Mon_Bas1_DC, ")
            loComandoSeleccionar.AppendLine("       SUM(Mon_Imp1) AS Mon_Imp1_DC, ")
            loComandoSeleccionar.AppendLine("       SUM((Mon_Bru/30)) AS Mon_Bas1_DCP, ")
            loComandoSeleccionar.AppendLine("       SUM((Mon_Imp1/30)) AS Mon_Imp1_DCP  ")
            loComandoSeleccionar.AppendLine("INTO	#Temporal2 ")
            loComandoSeleccionar.AppendLine("FROM	Devoluciones_Clientes  ")
            loComandoSeleccionar.AppendLine("WHERE	Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("       AND Cod_Suc BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Fec_Ini BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cod_Cli BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cod_Ven BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("       AND Cod_Mon BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("GROUP BY Cod_Cli")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT  ")
            loComandoSeleccionar.AppendLine("       Clientes.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("       Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(#Temporal.Mon_Bas1_CXC), 0) AS Mon_Bas1_CXC, ")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(#Temporal.Mon_Imp1_CXC), 0) AS Mon_Imp1_CXC, ")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(#Temporal2.Mon_Bas1_DC), 0) AS Mon_Bas1_DC, ")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(#Temporal2.Mon_Imp1_DC), 0) AS Mon_Imp1_DC, ")
            loComandoSeleccionar.AppendLine("       SUM(ISNULL(#Temporal.Mon_Bas1_CXC, 0) - ISNULL(#Temporal2.Mon_Bas1_DC, 0))  AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("       SUM(ISNULL(#Temporal.Mon_Imp1_CXC, 0) - ISNULL(#Temporal2.Mon_Imp1_DC, 0))  AS Imp_Net, ")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(Mon_Bas1_CXCP), 0) AS Mon_Bas1_CXCP, ")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(Mon_Imp1_CXCP), 0) AS Mon_Imp1_CXCP, ")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(Mon_Bas1_DCP), 0) AS Mon_Bas1_DCP, ")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(Mon_Imp1_DCP), 0) AS Mon_Imp1_DCP, ")
            loComandoSeleccionar.AppendLine("       SUM(ISNULL(Mon_Bas1_CXCP, 0) - ISNULL(Mon_Bas1_DCP, 0))  AS Mon_Netp, ")
            loComandoSeleccionar.AppendLine("       SUM(ISNULL(Mon_Imp1_CXCP, 0) - ISNULL(Mon_Imp1_DCP, 0))  AS Imp_Netp ")
            loComandoSeleccionar.AppendLine("FROM	#Temporal #Temporal ")
            loComandoSeleccionar.AppendLine("LEFT JOIN  #Temporal2 as #Temporal2 on #Temporal.Cod_Cli = #Temporal2.Cod_Cli")
            loComandoSeleccionar.AppendLine("JOIN Clientes ON Clientes.Cod_Cli = #Temporal.Cod_Cli")
            loComandoSeleccionar.AppendLine("GROUP BY Clientes.Cod_cli, Clientes.Nom_Cli")
            loComandoSeleccionar.AppendLine("ORDER BY  " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRVentas_Clientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRVentas_Clientes.ReportSource = loObjetoReporte

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
' DLC: 30/06/2010: Programacion inicial
'-------------------------------------------------------------------------------------------'
' DLC: 22/09/2010: - Ajuste de los calculos en la consulta a la base de datos.
'                   - Ajuste de la apariencia en el reporte.
'                   - Ajuste en el ordenamiento de los datos.
'-------------------------------------------------------------------------------------------'
