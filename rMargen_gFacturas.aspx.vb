'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMargen_gFacturas"
'-------------------------------------------------------------------------------------------'
Partial Class rMargen_gFacturas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        Try
            ' Me.mescribirConsulta("Iniciales:" & cusAplicacion.goReportes.paParametrosFinales.Count)
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
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As String = cusAplicacion.goReportes.paParametrosIniciales(14)
            Dim lcParametro15Desde As String = cusAplicacion.goReportes.paParametrosIniciales(15)
            Dim lcParametro16Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(16))
            Dim lcParametro16Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(16))
            Dim lcParametro17Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(17))
            Dim lcParametro17Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(17))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcCosto As String  = "Cos_Pro1"

            Select Case lcParametro10Desde
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

		    Dim llGananciasRespectoAlCosto AS Boolean  = goOpciones.mObtener("GANCOSPRE", "L")

            Dim loComandoSeleccionar As New StringBuilder()

            
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpGanancia(	Documento	CHAR(10), 			")
            loComandoSeleccionar.AppendLine("							Fec_Ini		DATETIME, 			")
            loComandoSeleccionar.AppendLine("							Cod_Cli		CHAR(10), 			")
            loComandoSeleccionar.AppendLine("							Nom_Cli		CHAR(100),			")
            loComandoSeleccionar.AppendLine("							Cod_Ven		CHAR(10), 			")
            loComandoSeleccionar.AppendLine("							Base_A		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Base_B		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Costo_A		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Costo_B		DECIMAL(28, 10))	")
            loComandoSeleccionar.AppendLine("															")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpFinal(	Documento	CHAR(10), 			")
            loComandoSeleccionar.AppendLine("							Fec_Ini		DATETIME, 			")
            loComandoSeleccionar.AppendLine("							Cod_Cli		CHAR(10), 			")
            loComandoSeleccionar.AppendLine("							Nom_Cli		CHAR(100),			")
            loComandoSeleccionar.AppendLine("							Cod_Ven		CHAR(10), 			")
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
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpGanancia(Documento, Fec_Ini, Cod_Cli, Nom_Cli, Cod_Ven, Base_A, Base_B, Costo_A, Costo_B)")
            loComandoSeleccionar.AppendLine("SELECT			Facturas.Documento												AS Documento,")
            loComandoSeleccionar.AppendLine("				Facturas.Fec_Ini												AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("				Clientes.Cod_Cli 												AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli 												AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("				Facturas.Cod_Ven 												AS Cod_Ven,")
            'loComandoSeleccionar.AppendLine("				SUM(Facturas.Mon_Net - Facturas.mon_imp1) 						        AS Base_A,")
            loComandoSeleccionar.AppendLine("				SUM(  Renglones_Facturas.Mon_Net")
            loComandoSeleccionar.AppendLine("				    *(1-Facturas.por_des1/100+facturas.por_rec1/100) ")
			loComandoSeleccionar.AppendLine("			        *(1+")
			loComandoSeleccionar.AppendLine("			            CASE WHEN Facturas.Mon_Bru>0 ")
			loComandoSeleccionar.AppendLine("			                THEN (Facturas.mon_otr1+Facturas.mon_otr2+Facturas.mon_otr3)/Facturas.Mon_Bru")
			loComandoSeleccionar.AppendLine("			                ELSE 0")
			loComandoSeleccionar.AppendLine("			            END")
			loComandoSeleccionar.AppendLine("			        )) AS Base_A,")
            'loComandoSeleccionar.AppendLine("				SUM(COALESCE(Devoluciones.Mon_Net - Devoluciones.Mon_Imp1, 0))  AS Base_B,")
            loComandoSeleccionar.AppendLine("				SUM(COALESCE(Devoluciones.mon_net, 0))                                  AS Base_B,")
            loComandoSeleccionar.AppendLine("				SUM(Renglones_Facturas.Can_Art1*Renglones_Facturas.Cos_Pro1)	AS Costo_A,")
            loComandoSeleccionar.AppendLine("				SUM(COALESCE(Devoluciones.Can_Art1*Devoluciones.Cos_Pro1, 0))	AS Costo_B")
            loComandoSeleccionar.AppendLine("FROM			Facturas")
            loComandoSeleccionar.AppendLine(" 		JOIN 	Clientes")
            loComandoSeleccionar.AppendLine(" 			ON	Clientes.Cod_Cli = Facturas.Cod_Cli")
            loComandoSeleccionar.AppendLine(" 		JOIN 	Renglones_Facturas ")
            loComandoSeleccionar.AppendLine(" 			ON	Renglones_Facturas.Documento = Facturas.Documento")
            loComandoSeleccionar.AppendLine(" 	LEFT JOIN 	(	SELECT		Renglones_dClientes.Doc_Ori,")
            loComandoSeleccionar.AppendLine(" 								Renglones_dClientes.Ren_Ori,")
            loComandoSeleccionar.AppendLine(" 								Renglones_dClientes.Can_Art1,")
            loComandoSeleccionar.AppendLine(" 								Renglones_dClientes.Cos_Pro1,")
            loComandoSeleccionar.AppendLine(" 								(   Renglones_dClientes.Mon_Net")
            loComandoSeleccionar.AppendLine(" 								    *(1-Devoluciones_Clientes.por_des1/100+Devoluciones_Clientes.por_rec1/100)")
            loComandoSeleccionar.AppendLine(" 								    *(1+ ")
            loComandoSeleccionar.AppendLine(" 								        CASE WHEN Devoluciones_Clientes.Mon_Bru>0")
            loComandoSeleccionar.AppendLine(" 								        THEN ( Devoluciones_Clientes.mon_otr1")
            loComandoSeleccionar.AppendLine(" 								              +Devoluciones_Clientes.mon_otr2")
            loComandoSeleccionar.AppendLine(" 								              +Devoluciones_Clientes.mon_otr3")
            loComandoSeleccionar.AppendLine(" 								             )/Devoluciones_Clientes.Mon_Bru")
            loComandoSeleccionar.AppendLine(" 								        ELSE 0 END")
            loComandoSeleccionar.AppendLine(" 								)) AS Mon_Net")
            loComandoSeleccionar.AppendLine(" 					FROM		Devoluciones_Clientes")
            loComandoSeleccionar.AppendLine(" 						JOIN	Renglones_dClientes ")
            loComandoSeleccionar.AppendLine(" 							ON	Renglones_dClientes.Documento = Devoluciones_Clientes.Documento")
            loComandoSeleccionar.AppendLine(" 					WHERE		Devoluciones_Clientes.Status IN (" & lcParametro7Desde & ")")
            'loComandoSeleccionar.AppendLine("						AND Devoluciones_Clientes.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine(" 							AND	Renglones_dClientes.Tip_Ori = 'Facturas'")
            loComandoSeleccionar.AppendLine(" 				) AS Devoluciones")
            loComandoSeleccionar.AppendLine(" 			ON	Devoluciones.Doc_Ori = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine(" 			AND	Devoluciones.Ren_Ori = Renglones_Facturas.Renglon")
            loComandoSeleccionar.AppendLine(" 		JOIN 	Vendedores ")
            loComandoSeleccionar.AppendLine(" 			ON	Vendedores.Cod_Ven = Facturas.Cod_Ven")
            loComandoSeleccionar.AppendLine(" 		JOIN	Articulos ")
            loComandoSeleccionar.AppendLine(" 			ON	Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE		Facturas.Documento				BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Fec_Ini 			BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Cli 			BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Clientes.Cod_Tip 			BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Clientes.Cod_Cla 			BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Ven 			BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			AND Vendedores.Cod_Tip			BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Status IN (" & lcParametro7Desde & ")")
            'loComandoSeleccionar.AppendLine("			AND Facturas.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine("			AND Renglones_Facturas.Cod_Art	BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Dep			BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Mon 			BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Tra 			BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_For 			BETWEEN " & lcParametro13Desde & " AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Rev 			BETWEEN " & lcParametro16Desde & " AND " & lcParametro16Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Suc 			BETWEEN " & lcParametro17Desde & " AND " & lcParametro17Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Facturas.Documento, Facturas.Fec_Ini, Clientes.Cod_Cli, Clientes.Nom_Cli, Facturas.Cod_Ven")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("/* Cálculo de ganancia								 										*/ ")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(Documento, Fec_Ini, Cod_Cli, Nom_Cli, Cod_Ven, Base_A, Base_B, Costo_A, Costo_B, Ganancia_A, Ganancia_B)")
            loComandoSeleccionar.AppendLine("SELECT	Documento				AS Documento,")
            loComandoSeleccionar.AppendLine("		Fec_Ini					AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Cod_Cli					AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Nom_Cli					AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		Cod_Ven					AS Cod_Ven,")
            loComandoSeleccionar.AppendLine("		SUM(Base_A)				AS Base_A,")
            loComandoSeleccionar.AppendLine("		SUM(Base_B)				AS Base_B,")
            loComandoSeleccionar.AppendLine("		SUM(Costo_A)			AS Costo_A,")
            loComandoSeleccionar.AppendLine("		SUM(Costo_B)			AS Costo_B,")
            loComandoSeleccionar.AppendLine("		0						AS Ganancia_A,")
            loComandoSeleccionar.AppendLine("		0						AS Ganancia_B")
            loComandoSeleccionar.AppendLine("FROM	#tmpGanancia")
            loComandoSeleccionar.AppendLine("GROUP BY	Documento, Fec_Ini, Cod_Cli, Nom_Cli, Cod_Ven")
            loComandoSeleccionar.AppendLine("")
            If llGananciasRespectoAlCosto Then 
                loComandoSeleccionar.AppendLine("UPDATE		#tmpFinal")
                loComandoSeleccionar.AppendLine("SET		Ganancia_A = (Base_A -Base_B) - (Costo_A - Costo_B),")
                loComandoSeleccionar.AppendLine("			Ganancia_B = (	CASE	")
                loComandoSeleccionar.AppendLine("								WHEN (Costo_A - Costo_B) <> 0")
                loComandoSeleccionar.AppendLine("								THEN ( (Base_A -Base_B) - (Costo_A - Costo_B))*100 / (Costo_A - Costo_B)")
                loComandoSeleccionar.AppendLine("								ELSE 0")
                loComandoSeleccionar.AppendLine("							END)")
            Else
                loComandoSeleccionar.AppendLine("UPDATE		#tmpFinal")
                loComandoSeleccionar.AppendLine("SET		Ganancia_A = (Base_A -Base_B) - (Costo_A - Costo_B),")
                loComandoSeleccionar.AppendLine("			Ganancia_B = (	CASE	")
                loComandoSeleccionar.AppendLine("								WHEN (Base_A - Base_B) <> 0")
                loComandoSeleccionar.AppendLine("								THEN ( (Base_A -Base_B) - (Costo_A - Costo_B))*100 / (Base_A - Base_B)")
                loComandoSeleccionar.AppendLine("								ELSE 0")
                loComandoSeleccionar.AppendLine("							END)")
            End If
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Documento				AS Documento,")
        	loComandoSeleccionar.AppendLine("		Fec_Ini					AS Fec_Ini,")
        	loComandoSeleccionar.AppendLine("		Cod_Cli					AS Cod_Cli,")
        	loComandoSeleccionar.AppendLine("		Nom_Cli					AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		Cod_Ven					AS Cod_Ven,")
        	loComandoSeleccionar.AppendLine("		Base_A					AS Base_A,")
        	loComandoSeleccionar.AppendLine("		Base_B					AS Base_B,")
        	loComandoSeleccionar.AppendLine("		Costo_A					AS Costo_A,")
        	loComandoSeleccionar.AppendLine("		Costo_B					AS Costo_B,")
            loComandoSeleccionar.AppendLine("		Ganancia_A				AS Ganancia_A,")
            loComandoSeleccionar.AppendLine("		Ganancia_B			AS Ganancia_B,")
            If llGananciasRespectoAlCosto Then
			    loComandoSeleccionar.AppendLine("		CAST(1 AS BIT)			AS Ganancia_SobreCosto")
            Else
			    loComandoSeleccionar.AppendLine("		CAST(0 AS BIT)			AS Ganancia_SobreCosto")
            End If
        	loComandoSeleccionar.AppendLine("FROM	#tmpFinal")
            Select Case lcParametro14Desde
                Case "Mayor"
                    loComandoSeleccionar.AppendLine("WHERE Ganancia_B > " & lcParametro15Desde)
                Case "Menor"
                    loComandoSeleccionar.AppendLine("WHERE Ganancia_B < " & lcParametro15Desde)
                Case "Igual"
                    loComandoSeleccionar.AppendLine("WHERE Ganancia_B = " & lcParametro15Desde)
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


            Dim loServicios As New cusDatos.goDatos
			
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMargen_gFacturas", laDatosReporte)

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
            Me.crvrMargen_gFacturas.ReportSource = loObjetoReporte

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
' CMS: 28/08/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 30/08/09: Se modifico el calculo de la columna Ganancia_B, para evitar la division
'                por cero
'-------------------------------------------------------------------------------------------'
' RJG: 06/09/12: Corrección de SELECT.
'-------------------------------------------------------------------------------------------'
' RJG: 16/01/14: Se agregó la opción para el cálculo de ganancias con respecto al precio o  '
'                costo. Se ajustó el SELECT para considerar los Descuentos, Recargos y Otros. 
'-------------------------------------------------------------------------------------------'
