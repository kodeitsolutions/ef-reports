Imports System.Data
Partial Class CGS_rGastos_Trabajadores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

            lcComandoSeleccionar.AppendLine("DECLARE @ldFechaIni	AS DATETIME = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @ldFechaFin	AS DATETIME = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcTrabIni		AS VARCHAR(10) = " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcTrabFin		AS VARCHAR(10) = " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcDepIni		AS VARCHAR(15) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcDepFin		AS VARCHAR(15) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcContIni		AS VARCHAR(2) = " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcContFin		AS VARCHAR(2) =" & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	Trabajadores.Cod_Tra,")
            lcComandoSeleccionar.AppendLine("		Trabajadores.Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		Cargos.Nom_Car,")
            lcComandoSeleccionar.AppendLine("		Departamentos_Nomina.Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		Renglones_Recibos.Mon_Net")
            lcComandoSeleccionar.AppendLine("INTO #tmpSueldo")
            lcComandoSeleccionar.AppendLine("FROM Recibos")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Recibos ON Recibos.Documento = Renglones_Recibos.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Conceptos_Nomina ON Renglones_Recibos.Cod_Con = Conceptos_Nomina.Cod_Con")
            lcComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Recibos.Cod_Tra = Trabajadores.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            lcComandoSeleccionar.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Recibos.Cod_Con = 'Q024'")
            lcComandoSeleccionar.AppendLine("	AND	Conceptos_Nomina.Status = 'A'")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con IN ('01', '02', '03','08')")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Fecha BETWEEN @ldFechaIni AND @ldFechaFin")
            lcComandoSeleccionar.AppendLine("	AND Trabajadores.Cod_Tra BETWEEN @lcTrabIni AND @lcTrabFin")
            lcComandoSeleccionar.AppendLine("	AND Departamentos_Nomina.Cod_Dep BETWEEN @lcDepIni AND @lcDepFin")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con BETWEEN @lcContIni AND @lcContFin")
            If lcParametro4Desde = "'A'" Then
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status = " & lcParametro4Desde)
            Else
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status IN (" & lcParametro4Desde & " )")
            End If
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	Trabajadores.Cod_Tra,")
            lcComandoSeleccionar.AppendLine("		Trabajadores.Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		Cargos.Nom_Car,")
            lcComandoSeleccionar.AppendLine("		Departamentos_Nomina.Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		SUM(Renglones_Recibos.Mon_Net)	AS Mon_Net")
            lcComandoSeleccionar.AppendLine("INTO #tmpAportes")
            lcComandoSeleccionar.AppendLine("FROM Recibos")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Recibos ON Recibos.Documento = Renglones_Recibos.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Conceptos_Nomina ON Renglones_Recibos.Cod_Con = Conceptos_Nomina.Cod_Con")
            lcComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Recibos.Cod_Tra = Trabajadores.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            lcComandoSeleccionar.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Recibos.Cod_Con IN ('U001','U002','U003','U004','U301','U302','U303','U403','U404')")
            lcComandoSeleccionar.AppendLine("	AND	Conceptos_Nomina.Status = 'A'")
            lcComandoSeleccionar.AppendLine("	AND	Conceptos_Nomina.Integrar = 1")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con IN ('01', '02', '03','08')")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Fecha BETWEEN @ldFechaIni AND @ldFechaFin")
            lcComandoSeleccionar.AppendLine("	AND Trabajadores.Cod_Tra BETWEEN @lcTrabIni AND @lcTrabFin")
            lcComandoSeleccionar.AppendLine("	AND Departamentos_Nomina.Cod_Dep BETWEEN @lcDepIni AND @lcDepFin")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con BETWEEN @lcContIni AND @lcContFin")
            If lcParametro4Desde = "'A'" Then
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status = " & lcParametro4Desde)
            Else
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status IN (" & lcParametro4Desde & " )")
            End If
            lcComandoSeleccionar.AppendLine("GROUP BY Trabajadores.Cod_Tra, Trabajadores.Nom_Tra, Cargos.Nom_Car, Departamentos_Nomina.Nom_Dep")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	Trabajadores.Cod_Tra,")
            lcComandoSeleccionar.AppendLine("		Trabajadores.Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		Cargos.Nom_Car,")
            lcComandoSeleccionar.AppendLine("		Departamentos_Nomina.Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		SUM(Renglones_Recibos.Mon_Net)	AS Mon_Net")
            lcComandoSeleccionar.AppendLine("INTO #tmpVacaciones")
            lcComandoSeleccionar.AppendLine("FROM Recibos")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Recibos ON Recibos.Documento = Renglones_Recibos.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Conceptos_Nomina ON Renglones_Recibos.Cod_Con = Conceptos_Nomina.Cod_Con")
            lcComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Recibos.Cod_Tra = Trabajadores.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            lcComandoSeleccionar.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Recibos.Cod_Con IN ('A300','A301','A302','A303','A304','A305','A403','A404','A405','A405', 'A406')")
            lcComandoSeleccionar.AppendLine("	AND	Conceptos_Nomina.Status = 'A'")
            lcComandoSeleccionar.AppendLine("	AND	Conceptos_Nomina.Integrar = 1")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con IN ('91')")
            'lcComandoSeleccionar.AppendLine("	AND Recibos.Fecha BETWEEN @ldFechaIni AND @ldFechaFin")
            lcComandoSeleccionar.AppendLine("	AND Trabajadores.Cod_Tra BETWEEN @lcTrabIni AND @lcTrabFin")
            lcComandoSeleccionar.AppendLine("	AND Departamentos_Nomina.Cod_Dep BETWEEN @lcDepIni AND @lcDepFin")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con BETWEEN @lcContIni AND @lcContFin")
            If lcParametro4Desde = "'A'" Then
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status = " & lcParametro4Desde)
            Else
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status IN (" & lcParametro4Desde & " )")
            End If
            lcComandoSeleccionar.AppendLine("GROUP BY Trabajadores.Cod_Tra, Trabajadores.Nom_Tra, Cargos.Nom_Car, Departamentos_Nomina.Nom_Dep")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	Trabajadores.Cod_Tra,")
            lcComandoSeleccionar.AppendLine("		Trabajadores.Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		Cargos.Nom_Car,")
            lcComandoSeleccionar.AppendLine("		Departamentos_Nomina.Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		SUM(Renglones_Recibos.Mon_Net)	AS Mon_Net")
            lcComandoSeleccionar.AppendLine("INTO #tmpUtilidades")
            lcComandoSeleccionar.AppendLine("FROM Recibos")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Recibos ON Recibos.Documento = Renglones_Recibos.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Conceptos_Nomina ON Renglones_Recibos.Cod_Con = Conceptos_Nomina.Cod_Con")
            lcComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Recibos.Cod_Tra = Trabajadores.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            lcComandoSeleccionar.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Recibos.Cod_Con IN ('A200','A402')")
            lcComandoSeleccionar.AppendLine("	AND	Conceptos_Nomina.Status = 'A'")
            lcComandoSeleccionar.AppendLine("	AND	Conceptos_Nomina.Integrar = 1")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con IN ('90')")
            'lcComandoSeleccionar.AppendLine("	AND Recibos.Fecha BETWEEN @ldFechaIni AND @ldFechaFin")
            lcComandoSeleccionar.AppendLine("	AND Trabajadores.Cod_Tra BETWEEN @lcTrabIni AND @lcTrabFin")
            lcComandoSeleccionar.AppendLine("	AND Departamentos_Nomina.Cod_Dep BETWEEN @lcDepIni AND @lcDepFin")
            lcComandoSeleccionar.AppendLine("	AND Recibos.Cod_Con BETWEEN @lcContIni AND @lcContFin")
            If lcParametro4Desde = "'A'" Then
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status = " & lcParametro4Desde)
            Else
                lcComandoSeleccionar.AppendLine("   AND Trabajadores.Status IN (" & lcParametro4Desde & " )")
            End If
            lcComandoSeleccionar.AppendLine("GROUP BY Trabajadores.Cod_Tra, Trabajadores.Nom_Tra, Cargos.Nom_Car, Departamentos_Nomina.Nom_Dep")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT DISTINCT")
            lcComandoSeleccionar.AppendLine("		#tmpSueldo.Cod_Tra, #tmpSueldo.Nom_Tra, #tmpSueldo.Nom_Car, #tmpSueldo.Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		#tmpSueldo.Mon_Net                  AS Sueldo,")
            lcComandoSeleccionar.AppendLine("		COALESCE(#tmpAportes.Mon_Net,0)     AS Aportes,")
            lcComandoSeleccionar.AppendLine("		COALESCE(#tmpVacaciones.Mon_Net,0)  AS Vacaciones,")
            lcComandoSeleccionar.AppendLine("		COALESCE(#tmpUtilidades.Mon_Net,0)  AS Utilidades,")
            lcComandoSeleccionar.AppendLine("		#tmpSueldo.Mon_Net + COALESCE(#tmpAportes.Mon_Net,0) + COALESCE(#tmpVacaciones.Mon_Net,0) + COALESCE(#tmpUtilidades.Mon_Net,0) AS Total,")
            lcComandoSeleccionar.AppendLine("		@ldFechaIni AS Desde, @ldFechaFin AS Hasta")
            lcComandoSeleccionar.AppendLine("FROM #tmpSueldo")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN #tmpAportes ON #tmpSueldo.Cod_Tra = #tmpAportes.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN #tmpVacaciones ON #tmpSueldo.Cod_Tra = #tmpVacaciones.Cod_Tra")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN #tmpUtilidades ON #tmpSueldo.Cod_Tra = #tmpUtilidades.Cod_Tra")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("UNION ALL")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	Proveedores.Cod_Pro			        AS Cod_Tra,")
            lcComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			        AS Nom_Tra,")
            lcComandoSeleccionar.AppendLine("		'PROVEEDOR'					        AS Nom_Car,")
            lcComandoSeleccionar.AppendLine("		'BONO DE ALIMENTACIÓN'		        AS Nom_Dep,")
            lcComandoSeleccionar.AppendLine("		0 AS Sueldo, 0 AS Aportes, 0 AS Vacaciones, 0 AS Utilidades,")
            lcComandoSeleccionar.AppendLine("		SUM(Cuentas_Pagar.Mon_Net)	        AS Total, ")
            lcComandoSeleccionar.AppendLine("		@ldFechaIni AS Desde, @ldFechaFin   AS Hasta")
            lcComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            lcComandoSeleccionar.AppendLine("	JOIN Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            lcComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'FACT'")
            lcComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Cod_Pro IN ('J003274445', 'J298405635')")
            lcComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Fec_Reg BETWEEN @ldFechaIni AND @ldFechaFin")
            lcComandoSeleccionar.AppendLine("GROUP BY Proveedores.Cod_Pro, Proveedores.Nom_Pro")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("ORDER BY Nom_Tra")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpSueldo")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpAportes")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpVacaciones")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpUtilidades")


            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rGastos_Trabajadores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rGastos_Trabajadores.ReportSource = loObjetoReporte

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
