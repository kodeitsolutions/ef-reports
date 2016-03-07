'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRVentas_MensualesPaises"
'-------------------------------------------------------------------------------------------'
Partial Class rRVentas_MensualesPaises
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       DATEPART(YEAR, Cuentas_Cobrar.Fec_Ini) AS Anio_CXC,")
            loComandoSeleccionar.AppendLine("       DATEPART(MONTH, Cuentas_Cobrar.Fec_Ini) AS Mes_CXC,")
            loComandoSeleccionar.AppendLine("       SUM(Cuentas_Cobrar.Mon_Bru) AS Mon_Bru_CXC,")
            loComandoSeleccionar.AppendLine("       SUM(Cuentas_Cobrar.Mon_Imp1) AS Mon_Imp1_CXC,")
            loComandoSeleccionar.AppendLine("       SUM((Cuentas_Cobrar.Mon_Bru/30 ))AS Mon_Bru_CXCP,")
            loComandoSeleccionar.AppendLine("       SUM((Cuentas_Cobrar.Mon_Imp1/30)) AS Mon_Imp1_CXCP,")
            loComandoSeleccionar.AppendLine("       Paises.Cod_Pai AS Cod_Pai_CXC")
            loComandoSeleccionar.AppendLine("INTO	#Temporal")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("JOIN	Paises ON Clientes.Cod_Pai = Paises.Cod_Pai")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Cobrar.Cod_Tip = 'Fact'")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Cli BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Ven BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Mon BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta & "")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Suc BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "")
            loComandoSeleccionar.AppendLine("GROUP BY DATEPART(YEAR, Cuentas_Cobrar.Fec_Ini),DATEPART(MONTH, Cuentas_Cobrar.Fec_Ini),Paises.Cod_Pai")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       DATEPART(YEAR, Devoluciones_Clientes.Fec_Ini) AS Anio_DC,")
            loComandoSeleccionar.AppendLine("       DATEPART(MONTH, Devoluciones_Clientes.Fec_Ini) AS Mes_DC,")
            loComandoSeleccionar.AppendLine("       SUM(Devoluciones_Clientes.Mon_Bru) AS Mon_Bru_DC,")
            loComandoSeleccionar.AppendLine("       SUM(Devoluciones_Clientes.Mon_Imp1) AS Mon_Imp1_DC,")
            loComandoSeleccionar.AppendLine("       SUM((Devoluciones_Clientes.Mon_Bru/30 ))AS Mon_Bru_DCP,")
            loComandoSeleccionar.AppendLine("       SUM((Devoluciones_Clientes.Mon_Imp1/30)) AS Mon_Imp1_DCP,")
            loComandoSeleccionar.AppendLine("       Paises.Cod_Pai AS Cod_Pai_DCP")
            loComandoSeleccionar.AppendLine("INTO	#Temporal2")
            loComandoSeleccionar.AppendLine("FROM	Devoluciones_Clientes")
            loComandoSeleccionar.AppendLine("JOIN	Clientes ON Devoluciones_Clientes.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("JOIN	Paises ON Clientes.Cod_Pai = Paises.Cod_Pai")
            loComandoSeleccionar.AppendLine("WHERE	Devoluciones_Clientes.Status IN ('Confirmado','Afectado','Procesado')")
            loComandoSeleccionar.AppendLine("       AND Devoluciones_Clientes.Fec_Ini BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
            loComandoSeleccionar.AppendLine("       AND Devoluciones_Clientes.Cod_Cli BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
            loComandoSeleccionar.AppendLine("       AND Devoluciones_Clientes.Cod_Ven BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
            loComandoSeleccionar.AppendLine("       AND Devoluciones_Clientes.Cod_Mon BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
            loComandoSeleccionar.AppendLine("GROUP BY DATEPART(YEAR, Devoluciones_Clientes.Fec_Ini),DATEPART(MONTH, Devoluciones_Clientes.Fec_Ini),Paises.Cod_Pai")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("       #Temporal.Anio_CXC,")
            loComandoSeleccionar.AppendLine("       #Temporal.Mes_CXC,")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(#Temporal.Mon_Bru_CXC), 0) AS Mon_Bas1_CXC,")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(#Temporal.Mon_Imp1_CXC), 0) AS Mon_Imp1_CXC,")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(#Temporal2.Mon_Bru_DC), 0) AS Mon_Bas1_DC,")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(#Temporal2.Mon_Imp1_DC), 0) AS Mon_Imp1_DC,")
            loComandoSeleccionar.AppendLine("       SUM(ISNULL(#Temporal.Mon_Bru_CXC, 0) - ISNULL(#Temporal2.Mon_Bru_DC, 0)) AS Mon_Net,")
            loComandoSeleccionar.AppendLine("       SUM(ISNULL(#Temporal.Mon_Imp1_CXC, 0) - ISNULL(#Temporal2.Mon_Imp1_DC, 0))AS Imp_Net,")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(Mon_Bru_CXCP), 0) AS Mon_Bas1_CXCP,")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(Mon_Imp1_CXCP), 0) AS Mon_Imp1_CXCP,")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(Mon_Bru_DCP), 0) AS Mon_Bas1_DCP,")
            loComandoSeleccionar.AppendLine("       ISNULL(SUM(Mon_Imp1_DCP), 0) AS Mon_Imp1_DCP,")
            loComandoSeleccionar.AppendLine("       SUM(ISNULL(Mon_Bru_CXCP, 0) - ISNULL(Mon_Bru_DCP, 0)) AS Mon_Netp,")
            loComandoSeleccionar.AppendLine("       SUM(ISNULL(Mon_Imp1_CXCP, 0) - ISNULL(Mon_Imp1_DCP, 0)) AS Imp_Netp,")
            loComandoSeleccionar.AppendLine("       Paises.Cod_Pai,")
            loComandoSeleccionar.AppendLine("       Paises.Nom_Pai")
            loComandoSeleccionar.AppendLine("INTO   #Result")
            loComandoSeleccionar.AppendLine("FROM	#Temporal ")
            loComandoSeleccionar.AppendLine("LEFT JOIN  #Temporal2 ON ((#Temporal.Anio_CXC = #Temporal2.Anio_DC) AND (#Temporal.Mes_CXC = #Temporal2.Mes_DC) AND (#Temporal.Cod_Pai_CXC = #Temporal2.Cod_Pai_DCP))")
            loComandoSeleccionar.AppendLine("JOIN Paises ON #Temporal.Cod_Pai_CXC = Paises.Cod_Pai OR #Temporal2.Cod_Pai_DCP = Paises.Cod_Pai")
            loComandoSeleccionar.AppendLine("WHERE	Paises.Cod_Pai BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta & "")
            loComandoSeleccionar.AppendLine("GROUP BY #Temporal.Anio_CXC, #Temporal.Mes_CXC, Paises.Cod_Pai, Paises.Nom_Pai")
            loComandoSeleccionar.AppendLine("ORDER BY Paises.Cod_Pai, #temporal.Anio_CXC ASC,  #temporal.Mes_CXC ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT * FROM #Result ORDER BY " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")

           ' Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRVentas_MensualesPaises", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrRVentas_MensualesPaises.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message & loExcepcion.StackTrace, _
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
' DLC: 01/07/2010 - Programación inicial
'-------------------------------------------------------------------------------------------'
' DLC: 16/07/2010 - Se ajusto el ordenamiento de la consulta, ordenandolo por nombre de pais,
'                   luego por codigo de pais, y el ordenamiento establecido por el usuario
'-------------------------------------------------------------------------------------------'
' DLC: 20/09/2010 - Se ajusto la consulta a la base de datos:
'                   .- Se cambio el parametro para el monto base: de Mon_Bas1 a Mon_Bru.
'                   .- Se agregó el filtro por pais.
'                   .- Se acomodo el metodo de ordenamiento: ordenar por codigo y nombre de pais.
'-------------------------------------------------------------------------------------------'
' MAT: 27/05/11: Ajuste del Select y Mejora de la vista de Diseño							'
'-------------------------------------------------------------------------------------------'
' MAT: 27/05/11: Eliminación del Filtro Estatus					'
'-------------------------------------------------------------------------------------------'
