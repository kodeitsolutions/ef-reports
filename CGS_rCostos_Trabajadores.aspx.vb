Imports System.Data
Partial Class CGS_rCostos_Trabajadores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

            lcComandoSeleccionar.AppendLine("SELECT Trabajadores.Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		Cargos.Nom_Car,")
            lcComandoSeleccionar.AppendLine("		Departamentos_Nomina.Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		SUM (Renglones_Recibos.Mon_Net) AS Total")
            lcComandoSeleccionar.AppendLine("INTO #tmpCostos")
            lcComandoSeleccionar.AppendLine("FROM Recibos")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Recibos ON Renglones_Recibos.Documento = Recibos.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Conceptos_Nomina ON Renglones_Recibos.Cod_Con = Conceptos_Nomina.Cod_Con")
            lcComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Recibos.Cod_Tra = Trabajadores.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            lcComandoSeleccionar.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            lcComandoSeleccionar.AppendLine("WHERE (Renglones_Recibos.Tipo = 'Asignacion'")
            lcComandoSeleccionar.AppendLine("       OR (Renglones_Recibos.Tipo = 'Otro'")
            lcComandoSeleccionar.AppendLine("	        AND SUBSTRING(Renglones_Recibos.Cod_Con,1,1)='U'")
            lcComandoSeleccionar.AppendLine("	        AND Renglones_Recibos.Cod_Con  <> 'U011'))")
            lcComandoSeleccionar.AppendLine("	AND	Conceptos_Nomina.Status = 'A'")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con IN ('01','02','08')")
            lcComandoSeleccionar.AppendLine("   AND Recibos.Fecha BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("   AND Trabajadores.Cod_Tra BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("   AND " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("   AND Departamentos_Nomina.Cod_Dep BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("   AND " & lcParametro2Hasta)
            If lcParametro3Desde = "'A'" Then
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status = " & lcParametro3Desde)
            Else
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status IN (" & lcParametro3Desde & " )")
            End If
            lcComandoSeleccionar.AppendLine("GROUP BY Trabajadores.Nom_Tra, Trabajadores.Cod_Tra, Cargos.Nom_Car, Departamentos_Nomina.Nom_Dep")

            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("UNION ALL")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Trabajadores.Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		Cargos.Nom_Car,")
            lcComandoSeleccionar.AppendLine("		Departamentos_Nomina.Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		SUM (Renglones_Recibos.Mon_Net) AS Total")
            lcComandoSeleccionar.AppendLine("FROM Recibos")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Recibos ON Renglones_Recibos.Documento = Recibos.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Conceptos_Nomina ON Renglones_Recibos.Cod_Con = Conceptos_Nomina.Cod_Con")
            lcComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Recibos.Cod_Tra = Trabajadores.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            lcComandoSeleccionar.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            lcComandoSeleccionar.AppendLine("WHERE ((Renglones_Recibos.Tipo = 'Asignacion'")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Recibos.Cod_Con = 'B070') OR (Renglones_Recibos.Tipo = 'Otro'")
            lcComandoSeleccionar.AppendLine("											AND Renglones_Recibos.Cod_Con IN ('U403', 'U404')))")
            lcComandoSeleccionar.AppendLine("	AND Conceptos_Nomina.Status = 'A'")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con = '92'")
            lcComandoSeleccionar.AppendLine("   AND Recibos.Fecha BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("   AND Trabajadores.Cod_Tra BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("   AND " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("   AND Departamentos_Nomina.Cod_Dep BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("   AND " & lcParametro2Hasta)
            If lcParametro3Desde = "'A'" Then
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status = " & lcParametro3Desde)
            Else
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status IN (" & lcParametro3Desde & " )")
            End If
            lcComandoSeleccionar.AppendLine("GROUP BY Trabajadores.Nom_Tra, Trabajadores.Cod_Tra, Cargos.Nom_Car, Departamentos_Nomina.Nom_Dep")

            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("UNION ALL")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Trabajadores.Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		Cargos.Nom_Car,")
            lcComandoSeleccionar.AppendLine("		Departamentos_Nomina.Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		(SUM (Renglones_Recibos.Mon_Net)) * -1 AS Total")
            lcComandoSeleccionar.AppendLine("FROM Recibos")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Recibos ON Renglones_Recibos.Documento = Recibos.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Conceptos_Nomina ON Renglones_Recibos.Cod_Con = Conceptos_Nomina.Cod_Con")
            lcComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Recibos.Cod_Tra = Trabajadores.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            lcComandoSeleccionar.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Recibos.Tipo = 'Deduccion'")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Recibos.Cod_Con IN ('E001', 'E002', 'E005', 'E100', 'E101', 'E103')")
            lcComandoSeleccionar.AppendLine("	AND Conceptos_Nomina.Status = 'A'")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con IN ('01', '02', '03')")
            lcComandoSeleccionar.AppendLine("   AND Recibos.Fecha BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("   AND Trabajadores.Cod_Tra BETWEEN " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("   AND " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("   AND Departamentos_Nomina.Cod_Dep BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("   AND " & lcParametro2Hasta)
            If lcParametro3Desde = "'A'" Then
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status = " & lcParametro3Desde)
            Else
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status IN (" & lcParametro3Desde & " )")
            End If
            lcComandoSeleccionar.AppendLine("GROUP BY Trabajadores.Nom_Tra, Trabajadores.Cod_Tra, Cargos.Nom_Car, Departamentos_Nomina.Nom_Dep")

            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("UNION ALL")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Proveedores.Nom_Pro			AS Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		'Proveedor'					AS Nom_Car,")
            lcComandoSeleccionar.AppendLine("		''							AS Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		SUM(Cuentas_Pagar.Mon_Net)	AS Total")
            lcComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            lcComandoSeleccionar.AppendLine("	JOIN Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            lcComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'FACT'")
            lcComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Cod_Pro = 'J003274445'")
            lcComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Fec_Reg BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("GROUP BY Proveedores.Nom_Pro")

            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Nom_Tra, Nom_Car, Nom_Dep, SUM(Total) AS Total,")
            lcComandoSeleccionar.AppendLine("	CAST( " & lcParametro0Desde & " AS DATE) AS Desde,")
            lcComandoSeleccionar.AppendLine("	CAST( " & lcParametro0Hasta & " AS DATE) AS Hasta")
            lcComandoSeleccionar.AppendLine("FROM #tmpCostos")
            lcComandoSeleccionar.AppendLine("GROUP BY Nom_Tra, Nom_Car, Nom_Dep")
            lcComandoSeleccionar.AppendLine("ORDER BY Nom_Tra")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpCostos")

            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rCostos_Trabajadores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rCostos_Trabajadores.ReportSource = loObjetoReporte

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
' JJD: 06/12/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
