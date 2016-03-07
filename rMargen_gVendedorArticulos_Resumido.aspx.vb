'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMargen_gVendedorArticulos_Resumido"
'-------------------------------------------------------------------------------------------'
Partial Class rMargen_gVendedorArticulos_Resumido
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
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosIniciales(6)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcCosto As String = ""

            Select Case lcParametro6Desde
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
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpGanancia(Cod_Ven, Nom_Ven, Cod_Art, Nom_Art, Can_Art, Can_Fac, Can_Dev, Base_A, Base_B, Costo_A, Costo_B)")
            loComandoSeleccionar.AppendLine("SELECT		Vendedores.Cod_Ven													AS Cod_Ven,")
            loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven													AS Nom_Ven,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Art													AS Cod_Art,")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art													AS Nom_Art,")
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
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Mon 			BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Rev 			BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Suc 			BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Vendedores.Cod_Ven, Vendedores.Nom_Ven, Articulos.Cod_Art, Articulos.Nom_Art")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("/* Datos de Devoluciones							 										*/")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpGanancia(Cod_Ven, Nom_Ven, Cod_Art, Nom_Art, Can_Art, Can_Fac, Can_Dev, Base_A, Base_B, Costo_A, Costo_B)")
            loComandoSeleccionar.AppendLine("SELECT		Vendedores.Cod_Ven													AS Cod_Ven,")
            loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven													AS Nom_Ven,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Art													AS Cod_Art,")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art													AS Nom_Art,")
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
            loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Mon 	BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Rev 	BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("			AND Devoluciones_Clientes.Cod_Suc 	BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Vendedores.Cod_Ven, Vendedores.Nom_Ven, Articulos.Cod_Art, Articulos.Nom_Art")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("/* Cálculo de ganancia								 										*/ ")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(Cod_Ven, Nom_Ven, Cod_Art, Nom_Art, Can_Art, Can_Fac, Can_Dev, Base_A, Base_B, Costo_A, Costo_B, Ganancia_A, Ganancia_B)")
            loComandoSeleccionar.AppendLine("SELECT	Cod_Ven				AS Cod_Ven,")
            loComandoSeleccionar.AppendLine("		Nom_Ven				AS Nom_Ven,")
            loComandoSeleccionar.AppendLine("		Cod_Art				AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		Nom_Art				AS Nom_Art,")
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
            loComandoSeleccionar.AppendLine("GROUP BY	Cod_Ven, Nom_Ven, Cod_Art, Nom_Art")
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
        	loComandoSeleccionar.AppendLine("		Can_Art				AS Can_Art,")
        	loComandoSeleccionar.AppendLine("		Can_Fac				AS Can_Fac,")
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
            Select Case lcParametro8Desde
                Case "Mayor"
                    loComandoSeleccionar.AppendLine("WHERE Ganancia_B > " & lcParametro9Desde)
                Case "Menor"
                    loComandoSeleccionar.AppendLine("WHERE Ganancia_B < " & lcParametro9Desde)
                Case "Igual"
                    loComandoSeleccionar.AppendLine("WHERE Ganancia_B = " & lcParametro9Desde)
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMargen_gVendedorArticulos_Resumido", laDatosReporte)

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

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrMargen_gVendedorArticulos_Resumido.ReportSource = loObjetoReporte

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
' MAT: 08/09/11: Programacion inicial
'-------------------------------------------------------------------------------------------'
' RJG: 05/09/12: Corrección de SELECT.
'-------------------------------------------------------------------------------------------'
