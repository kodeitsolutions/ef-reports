'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "PAS_rECuentas_Proveedores_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class PAS_rECuentas_Proveedores_Clientes
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

            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCod_Desde AS VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCod_Hasta AS VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  'PROVEEDOR'								AS Tipo,")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro						AS Codigo,     ")
            loComandoSeleccionar.AppendLine("		(SUM(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				  THEN Cuentas_Pagar.Mon_Sal      ")
            loComandoSeleccionar.AppendLine("				  ELSE 0 END)      ")
            loComandoSeleccionar.AppendLine("		- SUM(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				   THEN Cuentas_Pagar.Mon_Sal     ")
            loComandoSeleccionar.AppendLine("				   ELSE 0 END))					AS Sal_Ini     ")
            loComandoSeleccionar.AppendLine("INTO #tempSALDOINICIAL_Pro          ")
            loComandoSeleccionar.AppendLine("FROM Proveedores     ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro     ")
            loComandoSeleccionar.AppendLine("WHERE Proveedores.Cod_Cla = 'MIXTO'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Fec_Ini < @ldFecha_Desde  ")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCod_Desde AND @lcCod_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("GROUP BY Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  'CLIENTE'								AS Tipo,")
            loComandoSeleccionar.AppendLine("		Clientes.Cod_Cli						AS Codigo,     ")
            loComandoSeleccionar.AppendLine("		(SUM(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				  THEN Cuentas_Cobrar.Mon_Sal      ")
            loComandoSeleccionar.AppendLine("				  ELSE 0 END)      ")
            loComandoSeleccionar.AppendLine("		- SUM(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				   THEN Cuentas_Cobrar.Mon_Sal     ")
            loComandoSeleccionar.AppendLine("				   ELSE 0 END))					AS Sal_Ini     ")
            loComandoSeleccionar.AppendLine("INTO #tempSALDOINICIAL_Cli           ")
            loComandoSeleccionar.AppendLine("FROM Clientes     ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Cobrar ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli     ")
            loComandoSeleccionar.AppendLine("WHERE Clientes.Cod_Cla = 'MIXTO'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Fec_Ini < @ldFecha_Desde  ")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Cobrar.Cod_Cli BETWEEN @lcCod_Desde AND @lcCod_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("GROUP BY Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT 'PROVEEDOR'					AS Tipo,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Pro		AS Codigo,     ")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			AS Nombre,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tip,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Documento,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Control,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Reg,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Factura,    ")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("			 THEN Cuentas_Pagar.Mon_Sal                    ")
            loComandoSeleccionar.AppendLine("             ELSE 0    ")
            loComandoSeleccionar.AppendLine("         END						AS Mon_Deb,    ")
            loComandoSeleccionar.AppendLine("         CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito'")
            loComandoSeleccionar.AppendLine("			  THEN Cuentas_Pagar.Mon_Sal ")
            loComandoSeleccionar.AppendLine("              ELSE 0    ")
            loComandoSeleccionar.AppendLine("         END						AS Mon_Hab,    ")
            loComandoSeleccionar.AppendLine("         Cuentas_Pagar.Comentario")
            loComandoSeleccionar.AppendLine("INTO #tempMOVIMIENTO_Pro    ")
            loComandoSeleccionar.AppendLine("FROM Proveedores    ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro    ")
            loComandoSeleccionar.AppendLine("WHERE Proveedores.Cod_Cla = 'MIXTO'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Fec_Reg BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCod_Desde AND @lcCod_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT 'PROVEEDOR'					AS Tipo,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Pro		AS Codigo,     ")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			AS Nombre,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tip,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Documento,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Control,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Reg,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Factura,    ")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("			 THEN Cuentas_Pagar.Mon_Sal                    ")
            loComandoSeleccionar.AppendLine("             ELSE 0    ")
            loComandoSeleccionar.AppendLine("        END							AS Mon_Deb,    ")
            loComandoSeleccionar.AppendLine("        CASE WHEN Cuentas_Pagar.Tip_Doc = 'Debito'")
            loComandoSeleccionar.AppendLine("			 THEN Cuentas_Pagar.Mon_Sal ")
            loComandoSeleccionar.AppendLine("             ELSE 0    ")
            loComandoSeleccionar.AppendLine("         END						AS Mon_Hab,    ")
            loComandoSeleccionar.AppendLine("        Cuentas_Pagar.Comentario")
            loComandoSeleccionar.AppendLine("FROM Proveedores    ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro    ")
            loComandoSeleccionar.AppendLine("WHERE Proveedores.Cod_Cla = 'MIXTO'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Cod_Tip <> 'FACT'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Mon_Sal <> 0 ")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCod_Desde AND @lcCod_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT 'CLIENTE'					AS Tipo,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Cli		AS Codigo,     ")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli			AS Nombre,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Tip,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Documento,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Control,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Reg,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Factura,    ")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("			 THEN Cuentas_Cobrar.Mon_Sal                    ")
            loComandoSeleccionar.AppendLine("             ELSE 0    ")
            loComandoSeleccionar.AppendLine("       END						AS Mon_Deb,    ")
            loComandoSeleccionar.AppendLine("       CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito'")
            loComandoSeleccionar.AppendLine("			 THEN Cuentas_Cobrar.Mon_Sal ")
            loComandoSeleccionar.AppendLine("            ELSE 0    ")
            loComandoSeleccionar.AppendLine("       END						AS Mon_Hab,    ")
            loComandoSeleccionar.AppendLine("       Cuentas_Cobrar.Comentario")
            loComandoSeleccionar.AppendLine("INTO #tempMOVIMIENTO_Cli    ")
            loComandoSeleccionar.AppendLine("FROM Clientes    ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Cobrar ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli    ")
            loComandoSeleccionar.AppendLine("WHERE Clientes.Cod_Cla = 'MIXTO'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Cod_Tip <> 'RETIVA'")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Cobrar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Cobrar.Cod_Cli BETWEEN @lcCod_Desde AND @lcCod_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT 'CLIENTE'					AS Tipo,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Cli		AS Codigo,     ")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli			AS Nombre,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Tip,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Documento,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Control,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Reg,    ")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Factura,    ")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("			 THEN Cuentas_Cobrar.Mon_Sal                    ")
            loComandoSeleccionar.AppendLine("            ELSE 0    ")
            loComandoSeleccionar.AppendLine("       END						AS Mon_Deb,    ")
            loComandoSeleccionar.AppendLine("       CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito'")
            loComandoSeleccionar.AppendLine("		     THEN Cuentas_Cobrar.Mon_Sal ")
            loComandoSeleccionar.AppendLine("            ELSE 0    ")
            loComandoSeleccionar.AppendLine("       END						AS Mon_Hab,    ")
            loComandoSeleccionar.AppendLine("       Cuentas_Cobrar.Comentario")
            loComandoSeleccionar.AppendLine("FROM Clientes    ")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Cobrar ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli    ")
            loComandoSeleccionar.AppendLine("WHERE Clientes.Cod_Cla = 'MIXTO'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Cobrar.Fec_Reg BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Cobrar.Cod_Cli BETWEEN @lcCod_Desde AND @lcCod_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#tempMOVIMIENTO_Pro.Tipo,")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Codigo,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Nombre,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Cod_Tip,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Documento,  ")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Control,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Fec_Ini,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Fec_Reg,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Factura,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Mon_Deb,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Mon_Hab,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Pro.Comentario, ")
            loComandoSeleccionar.AppendLine("		CAST(@ldFecha_Desde AS DATE) AS Fec_Desde,")
            loComandoSeleccionar.AppendLine("		CAST(@ldFecha_Hasta AS DATE) AS Fec_Hasta,")
            loComandoSeleccionar.AppendLine("		COALESCE(#tempSALDOINICIAL_Pro.Sal_Ini,0) AS Sal_Ini,  	")
            loComandoSeleccionar.AppendLine("		0 AS Sal_Doc  	")
            loComandoSeleccionar.AppendLine("FROM	#tempMOVIMIENTO_Pro  	")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tempSALDOINICIAL_Pro ON #tempSALDOINICIAL_Pro.Codigo = #tempMOVIMIENTO_Pro.Codigo")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#tempMOVIMIENTO_Cli.Tipo,")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Codigo,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Nombre,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Cod_Tip,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Documento,  ")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Control,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Fec_Ini,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Fec_Reg,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Factura,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Mon_Deb,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Mon_Hab,  	")
            loComandoSeleccionar.AppendLine("		#tempMOVIMIENTO_Cli.Comentario, ")
            loComandoSeleccionar.AppendLine("		CAST(@ldFecha_Desde AS DATE) AS Fec_Desde,")
            loComandoSeleccionar.AppendLine("		CAST(@ldFecha_Hasta AS DATE) AS Fec_Hasta,")
            loComandoSeleccionar.AppendLine("		COALESCE(#tempSALDOINICIAL_Cli.Sal_Ini,0) AS Sal_Ini,  	")
            loComandoSeleccionar.AppendLine("		0 AS Sal_Doc  	")
            loComandoSeleccionar.AppendLine("FROM	#tempMOVIMIENTO_Cli  	")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tempSALDOINICIAL_Cli ON #tempSALDOINICIAL_Cli.Codigo = #tempMOVIMIENTO_Cli.Codigo")
            loComandoSeleccionar.AppendLine("ORDER BY Tipo,Codigo,Fec_Ini ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempSALDOINICIAL_Pro")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempSALDOINICIAL_Cli")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempMOVIMIENTO_Pro")
            loComandoSeleccionar.AppendLine("DROP TABLE #tempMOVIMIENTO_Cli")


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            If laDatosReporte.Tables(0).Rows.Count <> 0 Then

                '******************************************************************************************
                ' Se Procesa manualmente los datos
                '******************************************************************************************

                Dim loTabla As New DataTable("curReportes")
                Dim loColumna As DataColumn

                loColumna = New DataColumn("Tipo", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Codigo", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Nombre", GetType(String))
                loColumna.MaxLength = 200
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Cod_Tip", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Documento", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Control", GetType(String))
                loColumna.MaxLength = 50
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Ini", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Reg", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Factura", GetType(String))
                loColumna.MaxLength = 20
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Deb", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Mon_Hab", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Comentario", GetType(String))
                loColumna.MaxLength = 500
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Desde", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Fec_Hasta", GetType(String))
                loColumna.MaxLength = 10
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Ini", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                loColumna = New DataColumn("Sal_Doc", GetType(Decimal))
                loTabla.Columns.Add(loColumna)

                Dim loNuevaFila As DataRow
                Dim Cuenta_Actual As String
                Dim TipoActual As String
                Dim SaldoAnterior As Decimal = 0
                Dim lnTotalFilas As Integer = laDatosReporte.Tables(0).Rows.Count
                Dim loFila As DataRow

                '***************
                loFila = laDatosReporte.Tables(0).Rows(0)
                loNuevaFila = loTabla.NewRow()
                loTabla.Rows.Add(loNuevaFila)

                SaldoAnterior = loFila("Sal_Ini")

                loNuevaFila.Item("Tipo") = loFila("Tipo")
                loNuevaFila.Item("Codigo") = loFila("Codigo")
                loNuevaFila.Item("Nombre") = loFila("Nombre")
                loNuevaFila.Item("Cod_Tip") = loFila("Cod_Tip")
                loNuevaFila.Item("Documento") = loFila("Documento")
                loNuevaFila.Item("Control") = loFila("Control")
                loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "dd/MM/yyyy")
                loNuevaFila.Item("Fec_Reg") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Reg")), "dd/MM/yyyy")
                loNuevaFila.Item("Factura") = loFila("Factura")
                loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                loNuevaFila.Item("Comentario") = loFila("Comentario")
                loNuevaFila.Item("Fec_Desde") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Desde")), "dd/MM/yyyy")
                loNuevaFila.Item("Fec_Hasta") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Hasta")), "dd/MM/yyyy")
                loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

                SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                Cuenta_Actual = loFila("Codigo")
                TipoActual = loFila("Tipo")

                loTabla.AcceptChanges()

                For lnNumeroFila As Integer = 1 To lnTotalFilas - 1

                    loFila = laDatosReporte.Tables(0).Rows(lnNumeroFila)
                    loNuevaFila = loTabla.NewRow()
                    loTabla.Rows.Add(loNuevaFila)


                    If loFila("Codigo") <> Cuenta_Actual Or loFila("Tipo") <> TipoActual Then
                        SaldoAnterior = loFila("Sal_Ini")
                    End If

                    loNuevaFila.Item("Tipo") = loFila("Tipo")
                    loNuevaFila.Item("Codigo") = loFila("Codigo")
                    loNuevaFila.Item("Nombre") = loFila("Nombre")
                    loNuevaFila.Item("Cod_Tip") = loFila("Cod_Tip")
                    loNuevaFila.Item("Documento") = loFila("Documento")
                    loNuevaFila.Item("Control") = loFila("Control")
                    loNuevaFila.Item("Fec_Ini") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Ini")), "dd/MM/yyyy")
                    loNuevaFila.Item("Fec_Reg") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Reg")), "dd/MM/yyyy")
                    loNuevaFila.Item("Factura") = loFila("Factura")
                    loNuevaFila.Item("Mon_Deb") = loFila("Mon_Deb")
                    loNuevaFila.Item("Mon_Hab") = loFila("Mon_Hab")
                    loNuevaFila.Item("Comentario") = loFila("Comentario")
                    loNuevaFila.Item("Fec_Desde") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Desde")), "dd/MM/yyyy")
                    loNuevaFila.Item("Fec_Hasta") = Microsoft.VisualBasic.Format(CDate(loFila("Fec_Hasta")), "dd/MM/yyyy")
                    loNuevaFila.Item("Sal_Ini") = loFila("Sal_Ini")
                    loNuevaFila.Item("Sal_Doc") = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")

                    SaldoAnterior = SaldoAnterior + loFila("Mon_Deb") - loFila("Mon_Hab")
                    Cuenta_Actual = loFila("Codigo")

                    loTabla.AcceptChanges()

                Next lnNumeroFila


                Dim loDatosReporteFinal As New DataSet("curReportes")
                loDatosReporteFinal.Tables.Add(loTabla)


                '--------------------------------------------------------------------------------------'
                ' Se llena el reporte con la tabla nueva												'
                '--------------------------------------------------------------------------------------'
                Me.mCargarLogoEmpresa(loDatosReporteFinal.Tables(0), "LogoEmpresa")
                '-------------------------------------------------------------------------------------------------------
                ' Verificando si el select (tabla nº0) trae registros
                '-------------------------------------------------------------------------------------------------------

                If (loDatosReporteFinal.Tables(0).Rows.Count <= 0) Then
                    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información",
                                              "No se Encontraron Registros para los Parámetros Especificados. ",
                                               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion,
                                               "350px",
                                               "200px")
                End If

                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("PAS_rECuentas_Proveedores_Clientes", loDatosReporteFinal)
            Else

                ''-------------------------------------------------------------------------------------------------------
                '' Verificando si el select (tabla nº0) trae registros
                ''-------------------------------------------------------------------------------------------------------

                If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información",
                                              "No se Encontraron Registros para los Parámetros Especificados. ",
                                               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion,
                                               "350px",
                                               "200px")
                End If


                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("PAS_rECuentas_Proveedores_Clientes", laDatosReporte)
            End If

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvPAS_rECuentas_Proveedores_Clientes.ReportSource = loObjetoReporte

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
' CMS: 01/07/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS:  31/07/09: Filtro “Revisión:”, verificacion de registro
'-------------------------------------------------------------------------------------------'
' CMS:  03/08/09: Filtro “Tipo Revisión:”
'-------------------------------------------------------------------------------------------'
' MAT:  13/05/11: Ajuste de la Consulta, Tomaba los Saldos de Docuemntos Anulados
'-------------------------------------------------------------------------------------------'