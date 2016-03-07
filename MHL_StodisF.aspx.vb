'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MHL_StodisF"
'-------------------------------------------------------------------------------------------'
Partial Class MHL_StodisF
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcCosto As String = "Cos_Pro1"
            Dim loComandoSeleccionar As New StringBuilder()
            Dim lcAlmacenTransito As String

            '---------------------------------------------------------------------------------------------------------------'
            ' El código del Almacén de Tránsito se requiere para los Traslados entre Almacénes Confirmados o Procesados.	'
            '---------------------------------------------------------------------------------------------------------------'
            lcAlmacenTransito = goServicios.mObtenerCampoFormatoSQL(goOpciones.mObtener("CODALMTRA", "C"))

            '---------------------------------------------------------------------------------------------------------------'
            ' Calculo de Stock disponible a la fecha (Solo aplica para el almacen 'VTRANS'
            '---------------------------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine("SELECT		Renglones_Ajustes.Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada'  THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("INTO		#curDisp ")
            loComandoSeleccionar.AppendLine("FROM		Ajustes, ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine("WHERE		Ajustes.Documento		=	Renglones_Ajustes.Documento ")
            loComandoSeleccionar.AppendLine(" 		AND	Articulos.Cod_Art		=	Renglones_Ajustes.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Status			=	'Confirmado' ")
            loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Tipo	IN	('Entrada', 'Salida') ")
            loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Cod_Alm		= 'VTRANS' ")

            ' Union con Select de la tabla de Traslados en las Salidas
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT		Renglones_Traslados.Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			Traslados.Alm_Ori						AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("FROM		Traslados, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine("WHERE		Traslados.Documento		=	Renglones_Traslados.Documento ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Art		=	Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 		AND Traslados.Status		IN ('Confirmado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 		AND	Traslados.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		AND	Traslados.Alm_Ori		= 'VTRANS' ")

            ' Union con Select de la tabla de Traslados en las Entradas
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT		Renglones_Traslados.Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			CASE Traslados.Status					")
            loComandoSeleccionar.AppendLine(" 				WHEN 'Confirmado'	THEN Traslados.Alm_Ori")
            loComandoSeleccionar.AppendLine(" 				WHEN 'Procesado'	THEN Traslados.Alm_Des")
            loComandoSeleccionar.AppendLine(" 			END										AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("FROM		Traslados, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine("WHERE		Traslados.Documento		=	Renglones_Traslados.Documento ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Art		=	Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 		AND Traslados.Status		IN ('Confirmado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 		AND	Traslados.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		AND	(CASE	Traslados.Status				")
            loComandoSeleccionar.AppendLine(" 				WHEN 'Confirmado' THEN Traslados.Alm_Ori ")
            loComandoSeleccionar.AppendLine(" 				ELSE Traslados.Alm_Des				")
            loComandoSeleccionar.AppendLine(" 			END)						= 'VTRANS' ")

            ' Union con Select de la tabla de Entregas
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT			Renglones_Entregas.Cod_Art			AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Entregas.Cod_Alm			AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Entregas.Can_Art1			AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine(" 				0.0									AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("FROM			Entregas, ")
            loComandoSeleccionar.AppendLine("				Renglones_Entregas, ")
            loComandoSeleccionar.AppendLine("				Articulos ")
            loComandoSeleccionar.AppendLine("WHERE			Entregas.Documento			=	Renglones_Entregas.Documento ")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art			=	Renglones_Entregas.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 				AND Entregas.Status			IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 		AND	Entregas.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Entregas.Cod_Alm		= 'VTRANS' ")

            ' Union con Select de la tabla de Facturas
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT			Renglones_Facturas.Cod_Art			AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Cod_Alm			AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas.Can_Art1			AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine(" 				0.0									AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("FROM			Facturas, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine(" 				Articulos ")
            loComandoSeleccionar.AppendLine("WHERE			Facturas.Documento			=	Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art			=	Renglones_Facturas.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Facturas.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("			AND Renglones_Facturas.Tip_Ori			<>	'Entregas' ")
            loComandoSeleccionar.AppendLine(" 		AND	Facturas.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.Cod_Alm		= 'VTRANS' ")

            ' Union con Select de la tabla de Recepciones
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT			Renglones_Recepciones.Cod_Art		AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Recepciones.Cod_Alm		AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" 				0.0									AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Recepciones.Can_Art1		AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("FROM			Recepciones, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Recepciones, ")
            loComandoSeleccionar.AppendLine(" 				Articulos ")
            loComandoSeleccionar.AppendLine("WHERE		Recepciones.Documento			=	Renglones_Recepciones.Documento ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art			=	Renglones_Recepciones.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Recepciones.Status			IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 		AND	Recepciones.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Recepciones.Cod_Alm		= 'VTRANS' ")

            ' Union con Select de la tabla de Compras
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT			Renglones_Compras.Cod_Art			AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Cod_Alm			AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine("				0.0									AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Can_Art1			AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("FROM			Compras, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras, ")
            loComandoSeleccionar.AppendLine("				Articulos ")
            loComandoSeleccionar.AppendLine("WHERE			Compras.Documento				=	Renglones_Compras.Documento ")
            loComandoSeleccionar.AppendLine("				AND Articulos.Cod_Art			=	Renglones_Compras.Cod_Art ")
            loComandoSeleccionar.AppendLine("				AND Compras.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Compras.Tip_Ori	<>	'Recepciones' ")
            loComandoSeleccionar.AppendLine(" 		AND	Compras.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Compras.Cod_Alm		= 'VTRANS' ")

            ' Union con Select de la tabla de Devoluciones_Clientes
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT			Renglones_DClientes.Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_DClientes.Cod_Alm				AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" 				0.0										AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_DClientes.Can_Art1			AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("FROM			Devoluciones_Clientes, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_DClientes, ")
            loComandoSeleccionar.AppendLine(" 				Articulos ")
            loComandoSeleccionar.AppendLine("WHERE			Devoluciones_Clientes.Documento		=	Renglones_DClientes.Documento ")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Art					=	Renglones_DClientes.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 		AND	Devoluciones_Clientes.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_DClientes.Cod_Alm		= 'VTRANS' ")

            ' Union con Select de la tabla de Devoluciones_Proveedores
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT		Renglones_DProveedores.Cod_Art			AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Alm			AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Can_Art1			AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("FROM		Devoluciones_Proveedores, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores, ")
            loComandoSeleccionar.AppendLine(" 			Articulos ")
            loComandoSeleccionar.AppendLine("WHERE		Devoluciones_Proveedores.Documento	=	Renglones_DProveedores.Documento ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Art					=	Renglones_DProveedores.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 		AND	Devoluciones_Proveedores.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_DProveedores.Cod_Alm		= 'VTRANS' ")

            loComandoSeleccionar.AppendLine(" SELECT    MAX(Cod_Art) As Cod_Art, ")
            loComandoSeleccionar.AppendLine("           MAX(Cod_Alm) As Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Cod_Cla		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine("           0 As Stock_Transito, ")
            loComandoSeleccionar.AppendLine("           SUM(Can_Ent)-SUM(Can_Sal) As Stock_Disponible ")
            loComandoSeleccionar.AppendLine(" INTO      #curDisponible ")
            loComandoSeleccionar.AppendLine(" FROM      #curDisp ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Cod_Art,Cod_Cla")

            '---------------------------------------------------------------------------------------------------------------'
            ' Calculo de Stock en Transito a la fecha (Solo aplica para los almacenes 'VNAV' Y 'ADNA')
            '---------------------------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine("SELECT			Renglones_OCompras.Cod_Art			AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("				Renglones_OCompras.Cod_Alm			AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Articulos.Cod_Cla = 'TC' THEN '1036' ELSE '1038' END)		AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine("				0.0									AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine("				Renglones_OCompras.Can_Art1			AS	Can_Ent ")
            loComandoSeleccionar.AppendLine("INTO		#curTran ")
            loComandoSeleccionar.AppendLine("FROM			Ordenes_Compras, ")
            loComandoSeleccionar.AppendLine("				Renglones_OCompras, ")
            loComandoSeleccionar.AppendLine("				Articulos ")
            loComandoSeleccionar.AppendLine("WHERE			Ordenes_Compras.Documento		=	Renglones_OCompras.Documento ")
            loComandoSeleccionar.AppendLine("				AND Articulos.Cod_Art			=	Renglones_OCompras.Cod_Art ")
            loComandoSeleccionar.AppendLine("				AND Ordenes_Compras.Status		=	'Confirmado' ")
            loComandoSeleccionar.AppendLine(" 		        AND	Ordenes_Compras.Fec_Ini	<= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			    AND Renglones_OCompras.Cod_Alm	IN ('VNAV','ADNA') ")

            loComandoSeleccionar.AppendLine(" SELECT    MAX(Cod_Art)                AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("           MAX(Cod_Alm)                AS Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Cod_Cla                     AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine("           SUM(Can_Ent)-SUM(Can_Sal)   AS Stock_Transito, ")
            loComandoSeleccionar.AppendLine("           0 As Stock_Disponible ")
            loComandoSeleccionar.AppendLine(" INTO      #curTransito ")
            loComandoSeleccionar.AppendLine(" FROM      #curTran ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Cod_Art, Cod_Cla")

            loComandoSeleccionar.AppendLine(" SELECT    * ")
            loComandoSeleccionar.AppendLine(" INTO      #curStodis ")
            loComandoSeleccionar.AppendLine(" FROM #curDisponible ")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine(" SELECT    * ")
            loComandoSeleccionar.AppendLine(" FROM #curTransito ")

            loComandoSeleccionar.AppendLine(" SELECT    MAX(Cod_Art) As Cod_Art, ")
            loComandoSeleccionar.AppendLine("           MAX(Cod_Alm) As Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Cod_Cla				AS	Cod_Cla, ")
            loComandoSeleccionar.AppendLine("           SUM(Stock_Transito) As Stock_Transito, ")
            loComandoSeleccionar.AppendLine("           SUM(Stock_Disponible) As Stock_Disponible ")
            loComandoSeleccionar.AppendLine(" INTO      #curStodisF ")
            loComandoSeleccionar.AppendLine(" FROM      #curStodis ")
            loComandoSeleccionar.AppendLine(" WHERE     (Stock_Transito+Stock_Disponible)<> 0")
            loComandoSeleccionar.AppendLine(" GROUP BY  Cod_Art, Cod_Cla")
            loComandoSeleccionar.AppendLine(" ORDER BY  Cod_Art")

            loComandoSeleccionar.AppendLine(" SELECT	SUBSTRING(CONVERT(VARCHAR(10),GETDATE(),111),3,2) AS Anio_Emision, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(CONVERT(VARCHAR(10),GETDATE(),111),6,2) AS Mes_Emision, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(CONVERT(VARCHAR(10),GETDATE(),111),9,2) AS Dia_Emision, ")
            loComandoSeleccionar.AppendLine("           Cod_Art             AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Cod_Cla             AS Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" RIGHT('0000000'+CAST(ROUND(#curStodisF.Stock_Transito,0) AS INT),7) as Stock_Transito, ")
            loComandoSeleccionar.AppendLine(" RIGHT('0000000'+CAST(#curStodisF.Stock_Disponible as VARCHAR(20)),7) as Stock_Disponible, ")
            loComandoSeleccionar.AppendLine(" RIGHT('00'+CAST(DATEPART(MONTH," & lcParametro0Desde & ") AS VARCHAR(20)),2) as Mes ")
            loComandoSeleccionar.AppendLine(" FROM      #curStodisF ")

            Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------------------'
            ' Genera el archivo de texto    
            '-------------------------------------------------------------------------------------------------------'

            'Dim lcRif As String = goEmpresa.pcRifEmpresa.Replace("-", "").Replace(".", "").Replace(" ", "")
            Dim lcPeriodo As String = Strings.Format(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)), "yyyyMM")
            Dim loSalida As New StringBuilder()
            Dim lcContador As Integer = 0

            For Each loFila As DataRow In laDatosReporte.Tables(0).Rows

                'Dim lcAnio As String = CStr(loFila("Anio")).Trim()
                Dim lcAnioE As String = CStr(loFila("Anio_Emision")).Trim()
                Dim lcMes As String = CStr(loFila("Mes")).Trim()
                Dim lcMesE As String = CStr(loFila("Mes_Emision")).Trim()
                Dim lcDiaE As String = CStr(loFila("Dia_Emision")).Trim()
                Dim lcArticulo As String = CStr(loFila("Cod_Art")).Trim()
                Dim lcClase As String = CStr(loFila("Cod_Cla")).Trim()
                Dim lcStockTransito As String = CStr(loFila("Stock_Transito")).Trim()
                Dim lcStockDisponible As String = CStr(loFila("Stock_Disponible")).Trim()

                If lcContador = 0 Then
                    loSalida.Append("1F172    STODIS").Append(lcAnioE).Append(lcMesE).Append(lcDiaE).Append("VE").Append(lcMes).Append("001")
                    'loSalida.Append("        ")
                    'loSalida.Append("                    ")
                    'loSalida.Append("                    ")
                    'loSalida.Append("                    ")
                    'loSalida.Append("                    ")
                    loSalida.AppendLine()
                End If

                loSalida.Append("2")
                loSalida.Append(lcArticulo)
                loSalida.Append(lcClase)
                loSalida.Append("     ")
                loSalida.Append(lcStockDisponible)
                loSalida.Append("                        CO ")
                loSalida.Append(lcStockTransito)
                loSalida.Append("   0000000   0000000   0000000   0000000   0000000   0000000   0000000   0000000   0000000")
                loSalida.Append("   0000000   0000000   0000000   0000000   0000000   0000000   0000000   0000000   0000000")
                loSalida.AppendLine()
                lcContador = lcContador + 1
            Next loFila

            If lcContador < 10 Then
                loSalida.Append("90000").Append(lcContador).Append("001")
            Else
                If lcContador < 100 Then
                    loSalida.Append("9000").Append(lcContador).Append("001")
                Else
                    loSalida.Append("900").Append(lcContador).Append("001")
                End If
            End If

            '-------------------------------------------------------------------------------------------------------
            ' Envia la salida a pantalla en un archivo descargable.
            '-------------------------------------------------------------------------------------------------------
            Me.Response.Clear()
            Me.Response.ContentEncoding = System.Text.Encoding.UTF8
            Me.Response.AppendHeader("content-disposition", "attachment; filename=F172VEN" & ".txt")
            Me.Response.ContentType = "plain/text"
            Me.Response.Write(loSalida.ToString())
            'Me.Response.Write(Strings.Space(20))	'A veces no todo el texto es enviado a pantalla, entonces se 
            Me.Response.End()                       'mandan algunos espacios en blanco adicionales para "rellenar".
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
' MAT: 17/03/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RJG: 15/05/12: Corrección de los filtros de almacén (en traslados) y de sucursal.			'
'-------------------------------------------------------------------------------------------'
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                