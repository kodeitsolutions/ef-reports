﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports System
Imports System.Collections.Specialized
Imports System.Net

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMargen_gVendedorArticulos_Resumido_MOL"
'-------------------------------------------------------------------------------------------'
Partial Class rMargen_gVendedorArticulos_Resumido_MOL
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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
			    
			Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosIniciales(7)
			Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
			Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
			Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
			Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
			Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
			Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
			Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
			Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))


			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim lcCosto As String = ""

			Select Case lcParametro7Desde
				Case "Promedio MP"
					lcCosto = "Cos_Pro1"
				Case "Ultimo MP"
					lcCosto = "Cos_Ult1"
				Case "Anterior MP"
					lcCosto = "Cos_Ant1"
				Case "Promedio MS"
					lcCosto = "Cos_Pro2"
				Case "Ultimo MS"
					lcCosto = "Cos_Ult2"
				Case "Anterior MS"
					lcCosto = "Cos_Ant2"
			End Select

			Dim loComandoSeleccionar As New StringBuilder()



			loComandoSeleccionar.AppendLine("CREATE TABLE #tmpGanancia(	Cod_Ven		CHAR(30),			")
			loComandoSeleccionar.AppendLine("							Nom_Ven		CHAR(100),			")
			loComandoSeleccionar.AppendLine("							Cod_Art		CHAR(30),			")
			loComandoSeleccionar.AppendLine("							Nom_Art		CHAR(100),			")
			loComandoSeleccionar.AppendLine("							Cod_Mar		CHAR(10),			")
			loComandoSeleccionar.AppendLine("							Can_Art		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Can_Fac		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Can_Dev		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Base_A		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Base_B		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Costo_A		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Costo_B		DECIMAL(28, 10))	")
			loComandoSeleccionar.AppendLine("															")
			loComandoSeleccionar.AppendLine("CREATE TABLE #tmpFinal(	Cod_Ven		CHAR(30),			")
			loComandoSeleccionar.AppendLine("							Nom_Ven		CHAR(100),			")
			loComandoSeleccionar.AppendLine("							Cod_Art		CHAR(30),			")
			loComandoSeleccionar.AppendLine("							Nom_Art		CHAR(100),			")
			loComandoSeleccionar.AppendLine("							Cod_Mar		CHAR(10),			")
			loComandoSeleccionar.AppendLine("							Can_Art		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Can_Fac		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Can_Dev		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Base_A		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Base_B		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Costo_A		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Costo_B		DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Ganancia_A	DECIMAL(28, 10),	")
			loComandoSeleccionar.AppendLine("							Ganancia_B	DECIMAL(28, 10))	")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loComandoSeleccionar.AppendLine("/* Datos de Venta									 										*/")
			loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loComandoSeleccionar.AppendLine("INSERT INTO #tmpGanancia(Cod_Ven, Nom_Ven, Cod_Art, Nom_Art, Cod_Mar, Can_Art, Can_Fac, Can_Dev, Base_A, Base_B, Costo_A, Costo_B)")
			loComandoSeleccionar.AppendLine("SELECT		Vendedores.Cod_Ven													AS Cod_Ven,")
			loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven													AS Nom_Ven,")
			loComandoSeleccionar.AppendLine("			Articulos.Cod_Art													AS Cod_Art,")
			loComandoSeleccionar.AppendLine("			Articulos.Nom_Art													AS Nom_Art,")
			loComandoSeleccionar.AppendLine("			Articulos.Cod_Mar													AS Cod_Mar,")
			loComandoSeleccionar.AppendLine("			SUM(Renglones_Facturas.Can_Art1) 									AS Can_Art,")
			loComandoSeleccionar.AppendLine("			COUNT(DISTINCT Facturas.Documento) 									AS Can_Fac,")
			loComandoSeleccionar.AppendLine("			0								 									AS Can_Dev,")
			loComandoSeleccionar.AppendLine("			SUM(Renglones_Facturas.Mon_Net) 									AS Base_A,")
			loComandoSeleccionar.AppendLine("			0								 									AS Base_B,")
			loComandoSeleccionar.AppendLine("			SUM(Renglones_Facturas.Can_Art1*Renglones_Facturas." & lcCosto & ")		AS Costo_A,")
			loComandoSeleccionar.AppendLine("			0																	AS Costo_B")
			loComandoSeleccionar.AppendLine("FROM		Clientes")
			loComandoSeleccionar.AppendLine(" 	JOIN 	Facturas ")
			loComandoSeleccionar.AppendLine(" 		ON	Facturas.Cod_Cli = Clientes.Cod_Cli")
			loComandoSeleccionar.AppendLine(" 	JOIN 	Renglones_Facturas ")
			loComandoSeleccionar.AppendLine(" 		ON	Renglones_Facturas.Documento = Facturas.Documento")
			loComandoSeleccionar.AppendLine(" 	JOIN 	Vendedores ")
			loComandoSeleccionar.AppendLine(" 		ON	Vendedores.Cod_Ven = Facturas.Cod_Ven")
			loComandoSeleccionar.AppendLine(" 	JOIN	Articulos ")
			loComandoSeleccionar.AppendLine(" 		ON	Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.AppendLine("WHERE		Facturas.Documento				BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			AND Facturas.Fec_Ini 			BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Cli 			BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Ven 			BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("			AND Facturas.Status IN ('Confirmado', 'Afectado', 'Procesado')")
			loComandoSeleccionar.AppendLine("			AND Renglones_Facturas.Cod_Art	BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Dep			BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Mar			BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Mon 			BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
			loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Rev 			BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta)
			loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Suc 			BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta)
			loComandoSeleccionar.AppendLine("GROUP BY	Vendedores.Cod_Ven, Vendedores.Nom_Ven, Articulos.Cod_Art, Articulos.Nom_Art, Articulos.Cod_Mar")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loComandoSeleccionar.AppendLine("/* Datos de Devoluciones							 										*/")
			loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loComandoSeleccionar.AppendLine("INSERT INTO #tmpGanancia(Cod_Ven, Nom_Ven, Cod_Art, Nom_Art, Cod_Mar, Can_Art, Can_Fac, Can_Dev, Base_A, Base_B, Costo_A, Costo_B)")
			loComandoSeleccionar.AppendLine("SELECT		Vendedores.Cod_Ven													AS Cod_Ven,")
			loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven													AS Nom_Ven,")
			loComandoSeleccionar.AppendLine("			Articulos.Cod_Art													AS Cod_Art,")
			loComandoSeleccionar.AppendLine("			Articulos.Nom_Art													AS Nom_Art,")
			loComandoSeleccionar.AppendLine("			Articulos.Cod_Mar													AS Cod_Mar,")
			loComandoSeleccionar.AppendLine("			-SUM(Renglones_dClientes.Can_Art1) 									AS Can_Art,")
			loComandoSeleccionar.AppendLine("			0								 									AS Can_Fac,")
			loComandoSeleccionar.AppendLine("			COUNT(DISTINCT Devoluciones_Clientes.Documento) 					AS Can_Dev,")
			loComandoSeleccionar.AppendLine("			0								 									AS Base_A,")
			loComandoSeleccionar.AppendLine("			SUM(Renglones_dClientes.Mon_Net) 									AS Base_B,")
			loComandoSeleccionar.AppendLine("			0																	AS Costo_A,")
			loComandoSeleccionar.AppendLine("			SUM(Renglones_dClientes.Can_Art1*Renglones_dClientes.Cos_Pro1)		AS Costo_B")
			loComandoSeleccionar.AppendLine("FROM		Clientes")
			loComandoSeleccionar.AppendLine(" 	JOIN 	Devoluciones_Clientes ")
			loComandoSeleccionar.AppendLine(" 		ON	Devoluciones_Clientes.Cod_Cli = Clientes.Cod_Cli")
			loComandoSeleccionar.AppendLine(" 	JOIN 	Renglones_dClientes ")
			loComandoSeleccionar.AppendLine(" 		ON	Renglones_dClientes.Documento = Devoluciones_Clientes.Documento")
			loComandoSeleccionar.AppendLine(" 		AND	Renglones_dClientes.tip_Ori = 'Facturas'")
			loComandoSeleccionar.AppendLine(" 	JOIN 	Vendedores ")
			loComandoSeleccionar.AppendLine(" 		ON	Vendedores.Cod_Ven = Devoluciones_Clientes.Cod_Ven")
			loComandoSeleccionar.AppendLine(" 	JOIN	Articulos ")
			loComandoSeleccionar.AppendLine(" 		ON	Renglones_dClientes.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.AppendLine("WHERE		Devoluciones_Clientes.Documento		BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Fec_Ini 	BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Cli 	BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Ven	BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Status IN ('Confirmado', 'Afectado', 'Procesado')")
			loComandoSeleccionar.AppendLine("			AND Renglones_dClientes.Cod_Art		BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Dep				BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Mar				BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Mon 	BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Rev 	BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta)
			loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Suc 	BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta)
			loComandoSeleccionar.AppendLine("GROUP BY	Vendedores.Cod_Ven, Vendedores.Nom_Ven, Articulos.Cod_Art, Articulos.Nom_Art, Articulos.Cod_Mar")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loComandoSeleccionar.AppendLine("/* Cálculo de ganancia								 										*/ ")
			loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
			loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(Cod_Ven, Nom_Ven, Cod_Art, Nom_Art, Cod_Mar, Can_Art, Can_Fac, Can_Dev, Base_A, Base_B, Costo_A, Costo_B, Ganancia_A, Ganancia_B)")
			loComandoSeleccionar.AppendLine("SELECT	Cod_Ven				AS Cod_Ven,")
			loComandoSeleccionar.AppendLine("		Nom_Ven				AS Nom_Ven,")
			loComandoSeleccionar.AppendLine("		Cod_Art				AS Cod_Art,")
			loComandoSeleccionar.AppendLine("		Nom_Art				AS Nom_Art,")
			loComandoSeleccionar.AppendLine("		Cod_Mar				AS Cod_Mar,")
			loComandoSeleccionar.AppendLine("		SUM(Can_Art)		AS Can_Art,")
			loComandoSeleccionar.AppendLine("		SUM(Can_Fac)		AS Can_Fac,")
			loComandoSeleccionar.AppendLine("		SUM(Can_Dev)		AS Can_Dev,")
			loComandoSeleccionar.AppendLine("		SUM(Base_A)			AS Base_A,")
			loComandoSeleccionar.AppendLine("		SUM(Base_B)			AS Base_B,")
			loComandoSeleccionar.AppendLine("		SUM(Costo_A)		AS Costo_A,")
			loComandoSeleccionar.AppendLine("		SUM(Costo_B)		AS Costo_B,")
			loComandoSeleccionar.AppendLine("		0					AS Ganancia_A,")
			loComandoSeleccionar.AppendLine("		0					AS Ganancia_B")
			loComandoSeleccionar.AppendLine("FROM	#tmpGanancia")
			loComandoSeleccionar.AppendLine("GROUP BY	Cod_Ven, Nom_Ven, Cod_Art, Nom_Art, Cod_Mar")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("UPDATE		#tmpFinal")
			loComandoSeleccionar.AppendLine("SET		Ganancia_A = (Base_A -Base_B) - (Costo_A - Costo_B),")
			loComandoSeleccionar.AppendLine("			Ganancia_B = (	CASE	")
			loComandoSeleccionar.AppendLine("								WHEN (Base_A - Base_B) <> 0")
			loComandoSeleccionar.AppendLine("								THEN ( (Base_A -Base_B) - (Costo_A - Costo_B))*100 / (Base_A - Base_B)")
			loComandoSeleccionar.AppendLine("								ELSE 0")
			loComandoSeleccionar.AppendLine("							END)")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	Cod_Ven				AS Cod_Ven,")
			loComandoSeleccionar.AppendLine("		Nom_Ven				AS Nom_Ven,")
			loComandoSeleccionar.AppendLine("		Cod_Art				AS Cod_Art,")
			loComandoSeleccionar.AppendLine("		Nom_Art				AS Nom_Art,")
			loComandoSeleccionar.AppendLine("		Cod_Mar				AS Cod_Mar,")
			loComandoSeleccionar.AppendLine("		Can_Art				AS Can_Art,")
			loComandoSeleccionar.AppendLine("		Can_Fac				AS Can_Fac,")
			loComandoSeleccionar.AppendLine("		Can_Dev				AS Can_Dev,")
			loComandoSeleccionar.AppendLine("		(Base_A-Base_B)		AS Monto_Real,")
			loComandoSeleccionar.AppendLine("		(Costo_A-Costo_B)	AS Costo_Real,")
			loComandoSeleccionar.AppendLine("		Base_A				AS Base_A,")
			loComandoSeleccionar.AppendLine("		Base_B				AS Base_B,")
			loComandoSeleccionar.AppendLine("		Costo_A				AS Costo_A,")
			loComandoSeleccionar.AppendLine("		Costo_B				AS Costo_B,")
			loComandoSeleccionar.AppendLine("		Ganancia_A			AS Utilidad,")
			loComandoSeleccionar.AppendLine("		Ganancia_B			AS Porcentaje")
			loComandoSeleccionar.AppendLine("FROM	#tmpFinal")
           
			Select Case lcParametro9Desde
				Case "Mayor"
					loComandoSeleccionar.AppendLine("WHERE Ganancia_B > " & lcParametro10Desde)
				Case "Menor"
					loComandoSeleccionar.AppendLine("WHERE Ganancia_B < " & lcParametro10Desde)
				Case "Igual"
					loComandoSeleccionar.AppendLine("WHERE Ganancia_B = " & lcParametro10Desde)
				Case "Todos"
					'No filtra por Ganancia_B
			End Select
			
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpFinal")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'-------------------------------------------------------------------
            ' Selección de opcion por excel (Microsoft Excel - xls)
			'-------------------------------------------------------------------
            If (Me.Request.QueryString("salida").ToLower() = "xls") Then
                ' Genera el archivo a partir de la tabla de datos y termina la ejecución
                Me.mGenerarArchivoExcel(laDatosReporte)

            End If

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMargen_gVendedorArticulos_Resumido_MOL", laDatosReporte)
			'-------------------------------------------------------------------------------------------------------
			' Verificando si el select (tabla nº0) trae registros
			'-------------------------------------------------------------------------------------------------------

			If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
				Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
					"No se Encontraron Registros para los Parámetros Especificados. ", _
					vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
					"350px", "200px")
			End If

			Me.mTraducirReporte(loObjetoReporte)
			Me.mFormatearCamposReporte(loObjetoReporte)
			Me.crvrMargen_gVendedorArticulos_Resumido_MOL.ReportSource = loObjetoReporte


        Catch loExcepcion As Exception

			Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
				"No se pudo Completar el Proceso: " & loExcepcion.Message, _
				vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
              	"auto", "auto")

        End Try

    End Sub

	Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

		Try

			loObjetoReporte.Close()

		Catch loExcepcion As Exception

		End Try

	End Sub

    Private Sub mGenerarArchivoExcel(ByVal loDatos As DataSet)

    '***********************************************************************'
    ' Prepara los datos para enviarlos al servicio web de Excel.            '
    '***********************************************************************'
        Dim loSalida As New IO.MemoryStream()
        loDatos.WriteXml(loSalida, XmlWriteMode.WriteSchema)


    '***********************************************************************'
    ' Prepara los parámetros adicionales para enviarlos junto con los datos.'
    '***********************************************************************'
        Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
        Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
        Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje

        Dim llGananciasRespectoAlCosto As Boolean = goOpciones.mObtener("GANCOSPRE", "L")

        Dim loParametros As New NameValueCollection()
        loParametros.Add("lcNombreEmpresa", cusAplicacion.goEmpresa.pcNombre)
        loParametros.Add("lcRifEmpresa", cusAplicacion.goEmpresa.pcRifEmpresa)
        loParametros.Add("lnDecimalesMonto", lnDecimalesMonto.ToString())
        loParametros.Add("lnDecimalesCantidad", lnDecimalesCantidad.ToString())
        loParametros.Add("lnDecimalesPorcentaje", lnDecimalesPorcentaje.ToString())
        loParametros.Add("llGananciasRespectoAlCosto", llGananciasRespectoAlCosto.ToString())

        Dim loClienteWeb As new WebClient()
        loClienteWeb.QueryString = loParametros

    '***********************************************************************'
    ' Envía los datos y parámetros, y espera la respuesta.                  '
    '***********************************************************************'
        Dim loRespuesta As Byte()  
        Try
            loRespuesta = loClienteWeb.UploadData("http://localhost:8010/Reportes/rMargen_gVendedorArticulos_Resumido_MOL_xlsx.aspx", loSalida.GetBuffer())
        Catch ex As Exception
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado" , _ 
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 ex.Message, vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return
        End Try

    '***********************************************************************'
    ' Vemos si la respuesta es TextoPlano (error) o no (el archivo Excel    '
    ' generado). Si el tipo está vacio : error desconocido.                 '
    '***********************************************************************'
        Dim loTipoRespuesta As String = loClienteWeb.ResponseHeaders("Content-Type") 

        If String.IsNullOrEmpty(loTipoRespuesta) Then 
            'Error no especificado!
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado" , _ 
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: El servicio que genera la salida XSLX no responde.<br/>", _
                                                                 vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        ElseIf loTipoRespuesta.ToLower().StartsWith("text/plain") Then 

            Dim lcMensaje As String = UTF32Encoding.UTF8.GetString(loRespuesta)
            'Dim lcMensaje As String = ASCIIEncoding.ASCII.GetString(loRespuesta)
            lcMensaje = Me.Server.HtmlEncode(lcMensaje)

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado" , _ 
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 lcMensaje, vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        Else
            'Generación exitosa: la respuesta es el archivo en excel para descargar

            Me.Response.Clear()
            Me.Response.Buffer = True
            Me.Response.AppendHeader("content-disposition", "attachment; filename=rMargen_gVendedorArticulos_Resumido_MOL.xlsx")
            Me.Response.ContentType = "application/excel"
            Me.Response.BinaryWrite(loRespuesta)
            Me.Response.End()

        End If


    End Sub
      
	Private Sub mGenerarArchivoExcel2(ByVal lcNombreArchivo As String, ByVal loDatos As DataTable, ByVal lcParametrosReporte As String)
		
		'Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
		'Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
		'Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje
		
		'Dim lcFormatoMontos As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesMonto)
		
		'Dim lcFormatoCantidad As String 
		'If (lnDecimalesCantidad > 0) Then 
		'	lcFormatoCantidad = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesCantidad)
		'Else
		'	lcFormatoCantidad = "###,###,###,###,##0"
		'End If
		
		'Dim lcFormatoPorcentaje As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesPorcentaje)

	 ''******************************************************************'
	 '' Declaración de objetos de excel: IMPORTANTE liberar recursos al	'
	 '' final usando el GARBAGE COLLECTOR y ReleaseComObject.			'
	 ''******************************************************************'
		'Dim loExcel		As Excel.Application	= Nothing
		'Dim laLibros	As Excel.Workbooks		= Nothing
		'Dim loLibro		As Excel.Workbook		= Nothing
  '      Dim loHoja		As Excel.Worksheet		= Nothing
		'Dim loCeldas	As Excel.Range			= Nothing
		'Dim loRango		As Excel.Range			= Nothing
		
		'Dim loFilas		As Excel.Range			= Nothing
		'Dim loColumnas	As Excel.Range			= Nothing
		'Dim loFormas	As Excel.Shapes			= Nothing
		'Dim loImagen	As Excel.Shape			= Nothing
		'Dim loFuente	As Excel.Font			= Nothing
		
		
  '      Try
        
  '      ' Se inicializa el objeto de aplicacion excel
  '          loExcel = New Excel.Application()
  '          loExcel.Visible = False
  '          loExcel.DisplayAlerts = False 

  '      ' Crea un nuevo libro de excel y activa la primera hoja
  '          laLibros = loExcel.Workbooks
  '          'loLibro = laLibros.Add()
            
  '          'Dim lcPlantilla As String = HttpContext.Current.Server.MapPath("~/Administrativo/Complementos/plantilla.xls")
  '          'System.IO.File.Copy(lcPlantilla, lcNombreArchivo)
  '          loLibro = laLibros.Open(lcNombreArchivo)
            
  '          loHoja = loLibro.Worksheets(1)
  '          loHoja.Activate()

		'' Formato por defecto de todas las celdas			
		'	loCeldas = loHoja.Range("A1:IV65536")
  '          'loCeldas = loHoja.Cells
		'	loCeldas.Clear()
  '          loFuente = loCeldas.Font
  '          loFuente.Size = 9
  '          loFuente.Name = "Tahoma"


		' '******************************************************************'
		' ' Encabezado de la hoja											'
		' '******************************************************************'
		'	'Dim lcLogo As String = goEmpresa.pcUrlLogo 
		'	'lcLogo = HttpContext.Current.Server.MapPath(lcLogo)
		'	'loFormas = loHoja.Shapes

		'	'loFormas.AddPicture(lcLogo,  Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 1, 1, 60, 60)
			
  '          loRango = loHoja.Range("A1")
  '          loRango.Value = cusAplicacion.goEmpresa.pcNombre
            
  '          loRango = loHoja.Range("A2")
  '          loRango.Value = cusAplicacion.goEmpresa.pcRifEmpresa

  '          loRango = loHoja.Range("B5:J5")
  '          loRango.Select()
  '          loRango.MergeCells = True
  '          loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
  '          loRango.Value = "Margen de Ganancia por Vendedor y Artículo Resumido (MOL)"
  '          loFuente = loRango.Font
  '          loFuente.Size = 14
  '          loFuente.Bold = True

  '          ' Fecha y hora de creacion
		'	Dim ldFecha As DateTime = Date.Now()
		'	loRango = loHoja.Range("J1")
		'	loRango.NumberFormat = "@"
		'	loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
		'	loRango.Value = ldFecha.ToString("dd/MM/yyyy")
			
		'	loRango = loHoja.Range("J2")
		'	loRango.NumberFormat = "@" 'La celda almacena un string
		'	loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
		'	loRango.Value = ldFecha.ToString("hh:mm:ss tt")

  '          ' Parametros del reporte
  '          loRango = loHoja.Range("B7:J7")
  '          loRango.Select()
  '          loRango.MergeCells = True
  '          loRango.Value = lcParametrosReporte
  '          loRango.WrapText = True
  '          loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify


		'	Dim lnFilaActual As Integer = 9

		' '******************************************************************'
		' ' Datos del Reporte												'
		' '******************************************************************'
			
		'	loRango = loHoja.Range("B" & lnFilaActual)
		'	loRango.Value = "Vendedor"
			
		'	loRango = loHoja.Range("C" & lnFilaActual)
		'	loRango.Value = "Código"
			
		'	loRango = loHoja.Range("D" & lnFilaActual)
		'	loRango.Value = "Nombre"
			
		'	loRango = loHoja.Range("E" & lnFilaActual)
		'	loRango.Value = "Unidades"
			
		'	loRango = loHoja.Range("F" & lnFilaActual)
		'	loRango.Value = "Costo"
			
		'	loRango = loHoja.Range("G" & lnFilaActual)
		'	loRango.Value = "Monto Facturado"
			
		'	loRango = loHoja.Range("H" & lnFilaActual)
		'	loRango.Value = "Utilidad"
			
		'	loRango = loHoja.Range("I" & lnFilaActual)
		'	loRango.Value = "%"
			
		'	loRango = loHoja.Range("J" & lnFilaActual)
		'	loRango.Value = "Marca"
						
		'	loRango = loHoja.Range("B" & lnFilaActual & ":J" & lnFilaActual)
		'	loFuente = loRango.Font
		'	loFuente.Bold = True
		'	loFuente.Color = Rgb(255, 255, 255)
		'	loRango.Interior.Color = Rgb(0, 51, 153)
			
		'	loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
		'	loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
		'	loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
			
		'	Dim lnFilaInicio As Integer  = lnFilaActual
		'	For lnRenglon As Integer = 0 To loDatos.Rows.Count - 1
		'		 Dim loRenglon As DataRow = loDatos.Rows(lnRenglon)
				 
		'		 lnFilaActual += 1
				 
		'		 'Vendedor
		'		 loRango = loHoja.Range("B" & lnFilaActual)
		'		 loRango.NumberFormat = "@"
		'		 loRango.Value = CStr(loRenglon("Cod_Ven")).Trim()
		'		 loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				 
		'		 'Código
		'		 loRango = loHoja.Range("C" & lnFilaActual)
		'		 loRango.NumberFormat = "@"
		'		 loRango.Value = CStr(loRenglon("Cod_Art")).Trim()
		'		 loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				 
		'		 'Nombre
		'		 loRango = loHoja.Range("D" & lnFilaActual)
		'		 loRango.NumberFormat = "@"
		'		 loRango.Value = CStr(loRenglon("Nom_Art")).Trim()
				 
		'		 'Unidades
		'		 loRango = loHoja.Range("E" & lnFilaActual)	
		'		 loRango.NumberFormat = lcFormatoCantidad '#.###.##0,00
		'		 loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Can_Art")), lnDecimalesCantidad)
						
		'		 'Costo
		'		 loRango = loHoja.Range("F" & lnFilaActual) 
		'		 loRango.NumberFormat = lcFormatoMontos '#.###.##0,00	
		'		 loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Costo_Real")), lnDecimalesMonto)
					
		'		 'Monto Facturado
		'		 loRango = loHoja.Range("G" & lnFilaActual) 
		'		 loRango.NumberFormat = lcFormatoMontos
		'		 loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Monto_Real")), lnDecimalesMonto)
				
		'		 'Utilidad
		'		 loRango = loHoja.Range("H" & lnFilaActual) 
		'		 loRango.NumberFormat = lcFormatoMontos '#.###.##0,00	
		'		 loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Utilidad")), lnDecimalesMonto)
					
		'		 '%
		'		 loRango = loHoja.Range("I" & lnFilaActual)   
		'		 loRango.NumberFormat = lcFormatoPorcentaje
		'		 loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Porcentaje")), lnDecimalesPorcentaje)

		'		 ' Marca
		'		 loRango = loHoja.Range("J" & lnFilaActual)
		'		 loRango.NumberFormat = "@"
		'		 loRango.Value = CStr(loRenglon("Cod_Mar")).Trim()
				 
		'	Next lnRenglon
			
		'	Dim lnTotal As Integer = loDatos.Rows.Count
		'	loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":J" & (lnFilaInicio + lnTotal))
		'	loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
					
		'	loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":J" & (lnFilaInicio + lnTotal))
			
		'	Dim lnDesde AS Integer = lnFilaInicio
		'	Dim lnHasta AS Integer = lnFilaInicio + lnTotal
			
		'	lnFilaInicio += lnTotal + 2
		'	loRango = loHoja.Range("B" & (lnFilaInicio) & ":C" & (lnFilaInicio))
		'	loRango.MergeCells = True
		'	loRango.NumberFormat = "@"
		'	loRango.Value = "Total Registros: " & lnTotal.ToString()
		'	loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

		'	loRango = loHoja.Range("D" & (lnFilaInicio))
		'	loRango.NumberFormat = "@"
		'	loRango.Value = "Total General: "
		'	loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

		'	loRango = loHoja.Range("E" & (lnFilaInicio))
		'	loRango.NumberFormat = lcFormatoMontos
		'	loRango.Formula = "=SUM(E" & lnDesde & ":E" & lnHasta	& ")"

		'	loRango = loHoja.Range("F" & (lnFilaInicio))
		'	loRango.NumberFormat = lcFormatoMontos
		'	loRango.Formula = "=SUM(F" & lnDesde & ":F" & lnHasta	& ")"

		'	loRango = loHoja.Range("G" & (lnFilaInicio))
		'	loRango.NumberFormat = lcFormatoMontos
		'	loRango.Formula = "=SUM(G" & lnDesde & ":G" & lnHasta	& ")"

		'	loRango = loHoja.Range("H" & (lnFilaInicio))
		'	loRango.NumberFormat = lcFormatoMontos
		'	loRango.Formula = "=SUM(H" & lnDesde & ":H" & lnHasta	& ")"

		'	loRango = loHoja.Range("I" & (lnFilaInicio))
		'	loRango.NumberFormat = lcFormatoMontos
		'	loRango.Formula = "=IF(G" & (lnFilaInicio) & ">0, H" & (lnFilaInicio) & "*100/G" & (lnFilaInicio) & ", 100)"

		'	loRango = loHoja.Range("B" & (lnFilaInicio) & ":J" & (lnFilaInicio))
		'	loFuente = loRango.Font
		'	loFuente.Bold = True
					
		'	loFilas = loCeldas.Rows
		'	loFilas.AutoFit()
			
		'	loColumnas = loCeldas.Rows
		'	loColumnas.AutoFit()
			
		'	loRango = loHoja.Range("B1:B" & lnFilaInicio)
		'	loRango.ColumnWidth = 12
			
		'	loRango = loHoja.Range("C1:C" & lnFilaInicio)
		'	loRango.ColumnWidth = 28
			
		'	loRango = loHoja.Range("D1:D" & lnFilaInicio)
		'	loRango.ColumnWidth = 60
			
		'	loRango = loHoja.Range("E1:E" & lnFilaInicio)
		'	loRango.ColumnWidth = 15
			
		'	loRango = loHoja.Range("F1:F" & lnFilaInicio)
		'	loRango.ColumnWidth = 15
			
		'	loRango = loHoja.Range("G1:G" & lnFilaInicio)
		'	loRango.ColumnWidth = 15
			
		'	loRango = loHoja.Range("H1:H" & lnFilaInicio)
		'	loRango.ColumnWidth = 15
			
		'	loRango = loHoja.Range("I1:I" & lnFilaInicio)
		'	loRango.ColumnWidth = 9
			
		'	loRango = loHoja.Range("J1:J" & lnFilaInicio)
		'	loRango.ColumnWidth = 12
			
  '          ' Seleccionamos la primera celda del libro
		'	loRango = loHoja.Range("A1")
  '          loRango.Select()

  '          'Guardamos los cambios del libro activo
  '          loLibro.SaveAs(lcNombreArchivo)
            
		' '******************************************************************'
		' ' IMPORTANTE: Forma correcta de liberar recursos!!!				'
		' '******************************************************************'
  '          ' Cerramos y liberamos recursos

  '      Catch loExcepcion As Exception
			
		'	Throw New Exception("No fue posible exportar los datos a excel. " & loExcepcion.Message, loExcepcion)
			
  '      Finally

		'	If (loFuente IsNot Nothing) Then
		'		System.Runtime.InteropServices.Marshal.ReleaseComObject(loFuente)
		'		loFuente = Nothing
		'	End If
			
		'	If (loFormas IsNot Nothing) Then
		'		System.Runtime.InteropServices.Marshal.ReleaseComObject(loFormas)
		'		loFormas = Nothing
		'	End If
			
		'	If (loRango IsNot Nothing) Then
		'		System.Runtime.InteropServices.Marshal.ReleaseComObject(loRango)
		'		loRango = Nothing
		'	End If
			
		'	If (loFilas IsNot Nothing) Then
		'		System.Runtime.InteropServices.Marshal.ReleaseComObject(loFilas)
		'		loFilas = Nothing
		'	End If
			
		'	If (loColumnas IsNot Nothing) Then
		'		System.Runtime.InteropServices.Marshal.ReleaseComObject(loColumnas)
		'		loColumnas = Nothing
		'	End If
			
		'	If (loCeldas IsNot Nothing) Then
		'		System.Runtime.InteropServices.Marshal.ReleaseComObject(loCeldas)
		'		loCeldas = Nothing
		'	End If
			
		'	If (loHoja IsNot Nothing) Then
		'		System.Runtime.InteropServices.Marshal.ReleaseComObject(loHoja)
		'		loHoja = Nothing
		'	End If
			
		'	If (loLibro IsNot Nothing) Then
		'		loLibro.Close(True)
		'		System.Runtime.InteropServices.Marshal.ReleaseComObject(loLibro)
		'		loLibro = Nothing
		'	End If

		'	If (laLibros IsNot Nothing) Then
		'		System.Runtime.InteropServices.Marshal.ReleaseComObject(laLibros)
		'		laLibros = Nothing
		'	End If
            
  '          loExcel.Quit()

		'	System.Runtime.InteropServices.Marshal.ReleaseComObject(loExcel)
  '          loExcel = Nothing 
            
  '          GC.Collect()
  '          GC.WaitForPendingFinalizers()
            
  '      End Try

    End Sub

  
    
End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 05/11/12: Programacion inicial
'-------------------------------------------------------------------------------------------'
