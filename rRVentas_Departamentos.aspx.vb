'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRVentas_Departamentos"
'-------------------------------------------------------------------------------------------'
Partial Class rRVentas_Departamentos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" DECLARE @lcFecIni DATETIME")
            loComandoSeleccionar.AppendLine(" DECLARE @lcFecFin DATETIME")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTablaBase TABLE(")
            loComandoSeleccionar.AppendLine("   Mes		INT		NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	Año		INT		NOT NULL")
            loComandoSeleccionar.AppendLine(" )")
            loComandoSeleccionar.AppendLine(" DECLARE @lcDifMeses INT")
            loComandoSeleccionar.AppendLine(" DECLARE @lcAño INT")
            loComandoSeleccionar.AppendLine(" DECLARE @lcMes INT")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @lcFecIni = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" SET @lcFecFin = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @lcFecIni = DATEADD(MONTH ,0, CONVERT(DATETIME ,CONVERT(VARCHAR(4 ),DATEPART(YEAR,@lcFecIni))+right('0'+CONVERT(VARCHAR (2),DATEPART(MONTH,@lcFecIni)),2)+ '01 00:00:00')) ")
            loComandoSeleccionar.AppendLine(" SET @lcFecFin = DATEADD(MONTH ,1, CONVERT(DATETIME ,CONVERT(VARCHAR(4 ),DATEPART(YEAR,@lcFecFin))+right('0'+CONVERT(VARCHAR (2),DATEPART(MONTH,@lcFecFin)),2)+ '01 23:59:59'))-1 ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" IF (DATEDIFF(MONTH,@lcFecIni,@lcFecFin) + 1) > 6")
            loComandoSeleccionar.AppendLine(" BEGIN")
            loComandoSeleccionar.AppendLine(" 	SET @lcFecIni = DATEADD(MONTH ,-5, CONVERT(DATETIME ,CONVERT(VARCHAR(4 ),DATEPART(YEAR,@lcFecFin))+right('0'+CONVERT(VARCHAR (2),DATEPART(MONTH,@lcFecFin)),2)+ '01 00:00:00'))")
            loComandoSeleccionar.AppendLine(" END")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @lcAño = DATEPART(YEAR,@lcFecIni)")
            loComandoSeleccionar.AppendLine(" SET @lcMes = DATEPART(MONTH,@lcFecIni)")
            loComandoSeleccionar.AppendLine(" SET @lcDifMeses = DATEDIFF(MONTH,@lcFecIni,@lcFecFin)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" WHILE @lcDifMeses >=0")
            loComandoSeleccionar.AppendLine(" BEGIN")
            loComandoSeleccionar.AppendLine(" 	INSERT INTO @lcTablaBase(Mes,Año) VALUES(@lcMes,@lcAño)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	SET @lcMes = @lcMes + 1")
            loComandoSeleccionar.AppendLine(" 	IF @lcMes > 12")
            loComandoSeleccionar.AppendLine(" 	BEGIN")
            loComandoSeleccionar.AppendLine(" 		SET @lcMes = 1")
            loComandoSeleccionar.AppendLine(" 		SET @lcAño = @lcAño +1")
            loComandoSeleccionar.AppendLine(" 	END")
            loComandoSeleccionar.AppendLine(" 	SET @lcDifMeses = @lcDifMeses - 1")
            loComandoSeleccionar.AppendLine(" END")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Cod_Dep,")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(SUM(Renglones_Facturas.Can_Art1),0) AS Ventas,")
            loComandoSeleccionar.AppendLine(" 		0 AS Compras,")
            loComandoSeleccionar.AppendLine(" 		0 AS Exi_Act1,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YEAR,Facturas.Fec_Ini) AS Año,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(MONTH,Facturas.Fec_Ini) AS Mes")
            loComandoSeleccionar.AppendLine(" INTO	#tempDATOS")
            loComandoSeleccionar.AppendLine(" FROM	Articulos")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN Renglones_Facturas ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Departamentos ON Articulos.Cod_dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine(" WHERE	Facturas.Fec_Ini BETWEEN @lcFecIni AND @lcFecFin")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.cod_Art BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Facturas.Cod_Cli BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Facturas.Cod_Ven BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Departamentos.Cod_Dep BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Facturas.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine(" 		AND Facturas.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine(" 		AND Facturas.Cod_Rev BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Facturas.Cod_Suc BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Departamentos.Cod_Dep,Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YEAR,Facturas.Fec_Ini), DATEPART(MONTH,Facturas.Fec_Ini)")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Cod_Dep,")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(-SUM(Renglones_dClientes.Can_Art1),0) AS Ventas,")
            loComandoSeleccionar.AppendLine(" 		0 AS Compras,")
            loComandoSeleccionar.AppendLine(" 		0 AS Exi_Act1,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YEAR,Devoluciones_Clientes.Fec_Ini) AS Año,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(MONTH,Devoluciones_Clientes.Fec_Ini) AS Mes")
            loComandoSeleccionar.AppendLine(" FROM	Articulos")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN Renglones_dClientes ON Articulos.Cod_Art = Renglones_dClientes.Cod_Art")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN Devoluciones_Clientes ON Renglones_dClientes.Documento = Devoluciones_Clientes.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Departamentos ON Articulos.Cod_dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine(" WHERE	Devoluciones_Clientes.Fec_Ini BETWEEN @lcFecIni AND @lcFecFin")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.cod_Art BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Clientes.Cod_Cli BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Clientes.Cod_Ven BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Departamentos.Cod_Dep BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Clientes.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Clientes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Clientes.Cod_Rev BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Clientes.Cod_Suc BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Departamentos.Cod_Dep, Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YEAR,Devoluciones_Clientes.Fec_Ini), DATEPART(MONTH,Devoluciones_Clientes.Fec_Ini)")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Cod_Dep,")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine(" 		0 AS Ventas,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(SUM(Renglones_Compras.Can_Art1),0) AS Compras,")
            loComandoSeleccionar.AppendLine(" 		0 AS Exi_Act1,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YEAR,Compras.Fec_Ini) AS Año,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(MONTH,Compras.Fec_Ini) AS Mes")
            loComandoSeleccionar.AppendLine(" FROM	Articulos")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN Renglones_Compras ON Articulos.Cod_Art = Renglones_Compras.Cod_Art")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN Compras ON Renglones_Compras.Documento = Compras.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Departamentos ON Articulos.Cod_dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine(" WHERE	Compras.Fec_Ini BETWEEN @lcFecIni AND @lcFecFin")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.cod_Art BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Compras.Cod_Pro BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Departamentos.Cod_Dep BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Compras.Status IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine(" 		AND Compras.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine(" 		AND Compras.Cod_Rev BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Compras.Cod_Suc BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Departamentos.Cod_Dep, Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YEAR,Compras.Fec_Ini),DATEPART(MONTH,Compras.Fec_Ini)")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Cod_Dep,")
            loComandoSeleccionar.AppendLine(" 		Departamentos.Nom_Dep,")
            loComandoSeleccionar.AppendLine(" 		0 AS Ventas,")
            loComandoSeleccionar.AppendLine(" 		0 AS Compras,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(SUM(Articulos.Exi_Act1),0) AS Exi_Act1,")
            loComandoSeleccionar.AppendLine(" 		tempBase.Año,")
            loComandoSeleccionar.AppendLine(" 		tempBase.Mes")
            loComandoSeleccionar.AppendLine(" FROM	Articulos")
            loComandoSeleccionar.AppendLine(" JOIN Departamentos ON Articulos.Cod_dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine(" CROSS JOIN @lcTablaBase tempBASE")
            loComandoSeleccionar.AppendLine(" WHERE Articulos.cod_Art BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Departamentos.Cod_Dep BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Departamentos.Cod_Dep, Departamentos.Nom_Dep, tempBase.Año, tempBase.Mes")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTablaResult	TABLE(")
            loComandoSeleccionar.AppendLine(" 	    Cod_Dep		VARCHAR(10)		NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	    Nom_Dep		VARCHAR(100)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	    Ventas		DECIMAL(28,2)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	    VentasFec	VARCHAR(500)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	    VentasDet	VARCHAR(500)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	    Compras		DECIMAL(28,2)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	    Exi_Act1	DECIMAL(28,2)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	    NumDecExi	INT			    NOT NULL")
            loComandoSeleccionar.AppendLine(" )")
            loComandoSeleccionar.AppendLine(" DECLARE @lcCodDep VARCHAR(10)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcNomDep VARCHAR(100)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcVentas DECIMAL(28,2)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcVentasFec VARCHAR(500)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcVentasDet VARCHAR(500)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcCompras DECIMAL(28,2)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcExiAct DECIMAL(28,2)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTotalVentas DECIMAL(28,2)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTotalCompras DECIMAL(28,2)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcCountRecord INT")
            loComandoSeleccionar.AppendLine(" DECLARE @lcMesStr CHAR(3)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @lcCountRecord = 0")
            loComandoSeleccionar.AppendLine(" SET @lcDifMeses = DATEDIFF(MONTH,@lcFecIni,@lcFecFin)+1")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" DECLARE CURSOR_RESULT CURSOR FOR")
            loComandoSeleccionar.AppendLine(" 	SELECT")
            loComandoSeleccionar.AppendLine(" 			#tempDATOS.Cod_Dep,")
            loComandoSeleccionar.AppendLine(" 			#tempDATOS.Nom_Dep,")
            loComandoSeleccionar.AppendLine(" 			SUM(#tempDATOS.Ventas) AS Ventas,")
            loComandoSeleccionar.AppendLine(" 			SUM(#tempDATOS.Compras) AS Compras,")
            loComandoSeleccionar.AppendLine(" 			SUM(#tempDATOS.Exi_Act1) AS Exi_Act1,")
            loComandoSeleccionar.AppendLine(" 			#tempDATOS.Año,")
            loComandoSeleccionar.AppendLine(" 			CASE")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine(" 				WHEN #tempDATOS.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine(" 			END AS Mes")
            loComandoSeleccionar.AppendLine(" 	FROM	#tempDATOS")
            loComandoSeleccionar.AppendLine(" 	GROUP BY #tempDATOS.Cod_Dep,#tempDATOS.Nom_Dep,#tempDATOS.Año,#tempDATOS.Mes")
            loComandoSeleccionar.AppendLine(" OPEN CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" FETCH NEXT FROM CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" INTO @lcCodDep,@lcNomDep,@lcVentas,@lcCompras,@lcExiAct,@lcAño,@lcMesStr")
            loComandoSeleccionar.AppendLine(" WHILE @@fetch_status = 0")
            loComandoSeleccionar.AppendLine(" BEGIN")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	SET @lcCountRecord = @lcCountRecord + 1")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	IF ( @lcCountRecord = 1 )")
            loComandoSeleccionar.AppendLine(" 	BEGIN")
            loComandoSeleccionar.AppendLine(" 		SET @lcTotalVentas = 0")
            loComandoSeleccionar.AppendLine(" 		SET @lcTotalCompras = 0")
            loComandoSeleccionar.AppendLine(" 		SET @lcVentasDet = ''")
            loComandoSeleccionar.AppendLine(" 		SET @lcVentasFec = ''")
            loComandoSeleccionar.AppendLine(" 	END")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	SET @lcTotalVentas = @lcTotalVentas + @lcVentas")
            loComandoSeleccionar.AppendLine(" 	SET @lcTotalCompras = @lcTotalCompras + @lcCompras")
            loComandoSeleccionar.AppendLine(" 	SET @lcVentasDet = @lcVentasDet + CAST(@lcVentas AS VARCHAR(30)) + '#'")
            loComandoSeleccionar.AppendLine(" 	SET @lcVentasFec = @lcVentasFec + @lcMesStr + '-' + CAST(@lcAño AS VARCHAR(4)) + '#'")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	IF ( @lcCountRecord = @lcDifMeses )")
            loComandoSeleccionar.AppendLine(" 	BEGIN")
            loComandoSeleccionar.AppendLine(" 		SET @lcCountRecord = 0")
            loComandoSeleccionar.AppendLine(" 		INSERT @lcTablaResult(Cod_Dep,Nom_Dep,Ventas,VentasFec,VentasDet,Compras,Exi_Act1,NumDecExi)")
            loComandoSeleccionar.AppendLine(" 		VALUES(@lcCodDep,@lcNomDep,@lcTotalVentas,@lcVentasFec,@lcVentasDet,@lcTotalCompras,@lcExiAct," & cusAplicacion.goOpciones.pnDecimalesParaCantidad & ")")
            loComandoSeleccionar.AppendLine(" 	END")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	FETCH NEXT FROM CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" 	INTO @lcCodDep,@lcNomDep,@lcVentas,@lcCompras,@lcExiAct,@lcAño,@lcMesStr")
            loComandoSeleccionar.AppendLine(" END")
            loComandoSeleccionar.AppendLine(" CLOSE CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" DEALLOCATE CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT * FROM @lcTablaResult ORDER BY " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRVentas_Departamentos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRVentas_Departamentos.ReportSource = loObjetoReporte

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
' DLC: 21/07/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 08/09/2010: Ajuste de los cálculos numéricos, con el tipo de datos FLOAt a DECIMAL(28,2)
'-------------------------------------------------------------------------------------------'
' DLC: 22/09/2010: - Se agrego los totales en cada columna que se muestra en el reporte.
'                   - Se verificó y ajustó la visulaización de los datos en el reporte.
'-------------------------------------------------------------------------------------------'

