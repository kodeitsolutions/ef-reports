'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rLEstadoCuenta_HistoricoCompras"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rLEstadoCuenta_HistoricoCompras
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
           
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecIni			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecFin			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodPro_Desde	AS VARCHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodPro_Hasta	AS VARCHAR(10)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET	@sp_FecIni          = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET	@sp_FecFin          = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("SET	@sp_CodPro_Desde    = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("SET	@sp_CodPro_Hasta    = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero	AS DECIMAL(28, 10)	;")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio	AS VARCHAR(10)	;")
            loComandoSeleccionar.AppendLine("SET @lnCero		= 0")
            loComandoSeleccionar.AppendLine("SET @lcVacio		= ''")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Saldo Inicial")
            loComandoSeleccionar.AppendLine("SELECT	Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE -Cuentas_Pagar.Mon_Net	")
            loComandoSeleccionar.AppendLine("			END) AS Sal_Ini")
            loComandoSeleccionar.AppendLine("INTO	#tmpSaldos_Iniciales")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores  ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Fec_Ini < @sp_FecIni")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("GROUP BY Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN -Renglones_Pagos.Mon_Abo")
            loComandoSeleccionar.AppendLine("				ELSE Renglones_Pagos.Mon_Abo	")
            loComandoSeleccionar.AppendLine("			END) +(Pagos.Mon_Ret + Pagos.Mon_Des) AS Sal_Ini")
            loComandoSeleccionar.AppendLine("FROM	Pagos")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("		AND	Pagos.Fec_Ini < @sp_FecIni")
            loComandoSeleccionar.AppendLine("		AND Pagos.Cod_Pro BETWEEN @sp_CodPro_Desde AND @sp_CodPro_Hasta")
            loComandoSeleccionar.AppendLine("		AND Pagos.Automatico = 0")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE	Pagos.Status IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("GROUP BY Proveedores.Cod_Pro, Pagos.Mon_Ret, Pagos.Mon_Des")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Movimientos")
            loComandoSeleccionar.AppendLine("SELECT	'Cuentas_Pagar'									AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro								AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro								AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tip							AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Factura							AS Documento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Reg							AS Registro,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Control							AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Mon_Sal")
            loComandoSeleccionar.AppendLine("INTO	#tmpMovimientos")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores  ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Fec_Reg BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Pagar.Cod_Pro BETWEEN @sp_CodPro_Desde AND @sp_CodPro_Hasta			")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Cuentas_Pagar'									AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro								AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro								AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tip							AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini							AS Registro,")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Pagar.Cod_Tip = 'ADEL' THEN")
            loComandoSeleccionar.AppendLine("            (SELECT CONCAT(RTRIM(Tip_Ope),' ', RTRIM(Num_Doc))")
            loComandoSeleccionar.AppendLine("            FROM Detalles_Pagos")
            loComandoSeleccionar.AppendLine("            WHERE Documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("            ) ")
            loComandoSeleccionar.AppendLine("            ELSE Cuentas_Pagar.Comentario")
            loComandoSeleccionar.AppendLine("        END											AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Mon_Sal")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores  ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Tip <> 'FACT'")
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Pro BETWEEN @sp_CodPro_Desde AND @sp_CodPro_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Pagos'											AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro								AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro								AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		@lcVacio										AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Pagos.Documento									AS Documento,")
            loComandoSeleccionar.AppendLine("		Pagos.Fec_Ini									AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Pagos.Fec_Ini									AS Registro,")
            loComandoSeleccionar.AppendLine("		CONCAT((SELECT CONCAT(RTRIM(Tip_Ope),' ', RTRIM(Num_Doc))")
            loComandoSeleccionar.AppendLine("            FROM Detalles_Pagos")
            loComandoSeleccionar.AppendLine("            WHERE Detalles_Pagos.Documento = Pagos.Documento),")
            loComandoSeleccionar.AppendLine("			 '/ ',SUBSTRING(Pagos.Comentario, 6, 30))        AS Referencia, ")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN	@lnCero")
            loComandoSeleccionar.AppendLine("				ELSE	Renglones_Pagos.Mon_Abo	")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN	Renglones_Pagos.Mon_Abo")
            loComandoSeleccionar.AppendLine("				ELSE	@lnCero	")
            loComandoSeleccionar.AppendLine("			END) -(Pagos.Mon_Ret + Pagos.Mon_Des) 		AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Mon_Sal")
            loComandoSeleccionar.AppendLine("FROM	Pagos")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("		AND	Pagos.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("		AND Pagos.Cod_Pro BETWEEN @sp_CodPro_Desde AND @sp_CodPro_Hasta")
            loComandoSeleccionar.AppendLine("		AND Pagos.Automatico = 0")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE	Pagos.Status IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("GROUP BY	Proveedores.Cod_Pro, Proveedores.Nom_Pro, Pagos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("			Pagos.Documento, Pagos.Fec_Ini, Pagos.Registro, Pagos.Comentario,")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Ret, Pagos.Mon_Des")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE	#tmpMovimientos")
            loComandoSeleccionar.AppendLine("SET		Orden = M.Orden,")
            loComandoSeleccionar.AppendLine("		Mon_Sal = M.Mon_Deb-M.Mon_Hab,")
            loComandoSeleccionar.AppendLine("		Sal_Ini = M.Sal_Ini")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT	ROW_NUMBER() ")
            loComandoSeleccionar.AppendLine("						OVER (	PARTITION BY #tmpMovimientos.Cod_Pro ")
            loComandoSeleccionar.AppendLine("								ORDER BY #tmpMovimientos.Fec_Ini, (CASE WHEN #tmpMovimientos.Cod_Tip='' THEN 'zzzzzzzzz' ELSE #tmpMovimientos.Cod_Tip END ) ASC) AS Orden,")
            loComandoSeleccionar.AppendLine("					#tmpMovimientos.Tabla, #tmpMovimientos.Cod_Tip, #tmpMovimientos.Documento,")
            loComandoSeleccionar.AppendLine("					ISNULL(SI.Sal_Ini, @lnCero) AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("					#tmpMovimientos.Mon_Deb AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("					#tmpMovimientos.Mon_Hab AS Mon_Hab")
            loComandoSeleccionar.AppendLine("			FROM	#tmpMovimientos			")
            loComandoSeleccionar.AppendLine("			LEFT JOIN (SELECT Cod_Pro, SUM(Sal_Ini) AS Sal_Ini FROM #tmpSaldos_Iniciales GROUP BY Cod_Pro) AS SI")
            loComandoSeleccionar.AppendLine("				ON SI.Cod_Pro = #tmpMovimientos.Cod_Pro")
            loComandoSeleccionar.AppendLine("		) AS M		")
            loComandoSeleccionar.AppendLine("WHERE	M.Tabla = #tmpMovimientos.Tabla ")
            loComandoSeleccionar.AppendLine("	AND	M.Cod_Tip = #tmpMovimientos.Cod_Tip")
            loComandoSeleccionar.AppendLine("	AND	M.Documento = #tmpMovimientos.Documento")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	A.Orden, A.Tabla, A.Cod_Pro, A.Nom_Pro, A.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("		A.Documento, A.Fec_Ini, A.Registro, A.Referencia, A.Sal_Ini,")
            loComandoSeleccionar.AppendLine("		A.Mon_Deb, A.Mon_Hab, SUM(B.Mon_Sal) +  A.Sal_Ini AS Sal_Doc,")
            loComandoSeleccionar.AppendLine("       @sp_FecIni AS Desde, @sp_FecFin	AS Hasta")
            loComandoSeleccionar.AppendLine("FROM	#tmpMovimientos AS A")
            loComandoSeleccionar.AppendLine("	JOIN #tmpMovimientos AS B")
            loComandoSeleccionar.AppendLine("		ON B.Cod_Pro = A.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND B.Orden <= A.Orden")
            loComandoSeleccionar.AppendLine("GROUP BY A.Orden, A.Tabla, A.Cod_Pro, A.Nom_Pro, A.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("		A.Documento, A.Fec_Ini, A.Registro, A.Referencia, A.Sal_Ini,")
            loComandoSeleccionar.AppendLine("		A.Mon_Deb, A.Mon_Hab ")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro ASC, A.Fec_Ini ASC, (CASE WHEN A.Cod_Tip='' THEN 'zzzzzzzzz' ELSE A.Cod_Tip END ) ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldos_Iniciales")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")
         

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rLEstadoCuenta_HistoricoCompras", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rLEstadoCuenta_HistoricoCompras.ReportSource = loObjetoReporte

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
' DLC: 02/09/2010: Programacion inicial (Replica del reporte rLEstadoCuenta_HistoricoVentas)'
'                   - Cambio de la consulta a procedimiento almacenado.						'
'-------------------------------------------------------------------------------------------'
' DLC: 15/09/2010: Ajuste en la forma de obtener los detalles de Pagos, asi como también,	'
'                ajustar en el RPT, la forma de mostrar los detalles de Pagos.				'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Reprogramación del Reporte y su respectivo Store Procedure					'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Ajuste de la vista de Diseño.												'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Se elimino el filtro Detalle												'
'-------------------------------------------------------------------------------------------'
' RJG: 05/12/11: Eliminado el SP: ahora la consulta se hace desde un Query en línea para	'
'				 corregir cálculo de saldo y optimizar.										'
'-------------------------------------------------------------------------------------------'
