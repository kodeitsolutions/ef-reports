'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fVentas_Compras_uAños"
'-------------------------------------------------------------------------------------------'
Partial Class fVentas_Compras_uAños

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim loComandoSeleccionar As New StringBuilder()
            Dim ldFechaFinal		As DateTime
            Dim ldFechaInicial		As DateTime
            Dim lnMesFinal	AS Integer =  Date.Now.Month + 1
            Dim lnAñoFinal AS Integer =  Date.Now.Year
            Dim lnMesInicial	AS Integer 
            Dim lcCodigo As String = cusAplicacion.goFormatos.pcCondicionPrincipal.Replace("(articulos.cod_art=", "")
            lcCodigo = lcCodigo.Replace(")","")
            
            ldFechaFinal = New Date(lnAñoFinal, lnMesFinal, 1, 0, 0, 0, 0).AddSeconds(-1)
			ldFechaInicial   = New Date(lnAñoFinal,lnMesFinal, 1, 0, 0, 0, 0).AddMonths(-24)

			
			'************ INSTRUCCION PARA CREAR LA TABLA DE CADA AÑO, MES *************************
			loComandoSeleccionar.AppendLine("CREATE TABLE #tmpMeses(Codigo CHAR(50),Año INT, Renglon INT, Nombre CHAR(30))")
			
			lnMesInicial = ldFechaInicial.Month
			
			Dim lcMes AS String = ""
			
			For i AS Integer = 0 To 23
			
				Dim lnMes As Integer = (lnMesInicial +i) mod 12
				Dim lnAño AS Integer
				
				If lnMes = 0 Then
					lnMes = 12
					lnAño=	((lnMesInicial+i)\12) - 1
				Else
				
					lnAño=	((lnMesInicial+i)\12)
				
				End If
				
				SELECT CASE lnMes
				
					CASE 1
					    lcMes  =  "'Enero'"
					CASE 2
						lcMes  =  "'Febrero'"
					CASE 3
						lcMes  =  "'Marzo'"
					CASE 4
						lcMes  =  "'Abril'"
					CASE 5
						lcMes  =  "'Mayo'"
					CASE 6
						lcMes  =  "'Junio'"
					CASE 7
						lcMes  =  "'Julio'"
					CASE 8
						lcMes  =  "'Agosto'"
					CASE 9
						lcMes  =  "'Septiembre'"
					CASE 10
						lcMes  =  "'Octubre'"
					CASE 11
						lcMes  =  "'Noviembre'"
					CASE 12
						lcMes  =  "'Diciembre'"

				END SELECT
				
				
				  loComandoSeleccionar.AppendLine("INSERT INTO #tmpMeses VALUES(" & lcCodigo & ", " & ldFechaInicial.Year + lnAño &  " , " & lnMes & " , " & lcMes & ")")		

			Next i
			
			'************ CREACIÓN DE LA CONSULTA *************************


			loComandoSeleccionar.Appendline("--**********************************************************************")
			loComandoSeleccionar.Appendline("--**********************TOTAL FACTURAS VENTAS***************************")
			loComandoSeleccionar.Appendline("--**********************************************************************")

			loComandoSeleccionar.Appendline("SELECT																		")
			loComandoSeleccionar.AppendLine("		Articulos.Cod_Art							AS Cod_Art,				")
			loComandoSeleccionar.AppendLine("		Articulos.Nom_Art							As Nom_Art,				")
			loComandoSeleccionar.Appendline("		DATEPART(YEAR, Facturas.Fec_Ini)			AS Año,					")
			loComandoSeleccionar.Appendline("		DATEPART(MONTH, Facturas.Fec_Ini)			AS Mes,					")
			loComandoSeleccionar.Appendline("		ISNULL(SUM(Renglones_Facturas.Can_Art1),0)	AS MontoNetoVenta,		")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10)) 					AS MontoDevueltoVenta,	")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10)) 					AS MontoNetoCompra,		")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10)) 					AS MontoDevueltoCompra	")
			loComandoSeleccionar.Appendline("INTO	#TmpTemporal	")
			loComandoSeleccionar.Appendline("FROM	Articulos		")
			loComandoSeleccionar.Appendline("JOIN	Renglones_Facturas ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.Appendline("JOIN	Facturas ON Renglones_Facturas.Documento = Facturas.Documento AND Facturas.Status IN ('Confirmado','Afectado','Procesado')")
			loComandoSeleccionar.Appendline("				AND Facturas.Fec_Ini BETWEEN " & goServicios.mObtenerCampoFormatoSQL(ldFechaInicial) & " AND " & goServicios.mObtenerCampoFormatoSQL(ldFechaFinal))
			loComandoSeleccionar.Appendline("WHERE	" &  cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.Appendline("GROUP BY DATEPART(YEAR, Facturas.Fec_Ini),DATEPART(MONTH, Facturas.Fec_Ini),Articulos.Cod_Art,Articulos.Nom_Art")
			loComandoSeleccionar.Appendline("  ")

			loComandoSeleccionar.Appendline("UNION ALL")
			
			loComandoSeleccionar.Appendline("  ")
			loComandoSeleccionar.Appendline("--**********************************************************************")
			loComandoSeleccionar.Appendline("--*********************TOTAL DEVOLUCIONES VENTAS************************")
			loComandoSeleccionar.Appendline("--**********************************************************************")

			loComandoSeleccionar.Appendline("SELECT																			")
			loComandoSeleccionar.AppendLine("		Articulos.Cod_Art								AS Cod_Art,				")
			loComandoSeleccionar.AppendLine("		Articulos.Nom_Art								As Nom_Art,				")
			loComandoSeleccionar.Appendline("		DATEPART(YEAR, devoluciones_Clientes.Fec_Ini)	AS Año,					")
			loComandoSeleccionar.Appendline("		DATEPART(MONTH, devoluciones_Clientes.Fec_Ini)	AS Mes,					")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10))						AS MontoNetoVenta,		")
			loComandoSeleccionar.Appendline("		ISNULL(SUM(Renglones_DClientes.Can_Art1),0)		AS MontoDevueltoVenta,	")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10))						AS MontoNetoCompra,		")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10))						AS MontoDevueltoCompra	")
			loComandoSeleccionar.Appendline("FROM	Articulos")
			loComandoSeleccionar.Appendline("JOIN	Renglones_DClientes ON Renglones_DClientes.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.Appendline("JOIN	devoluciones_Clientes ON Renglones_DClientes.Documento = devoluciones_Clientes.Documento AND devoluciones_Clientes.Status IN ('Confirmado','Afectado','Procesado')")
			loComandoSeleccionar.Appendline("				AND devoluciones_Clientes.Fec_Ini BETWEEN " & goServicios.mObtenerCampoFormatoSQL(ldFechaInicial) & " AND " & goServicios.mObtenerCampoFormatoSQL(ldFechaFinal))
			loComandoSeleccionar.Appendline("WHERE	" &  cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.Appendline("GROUP BY DATEPART(YEAR, devoluciones_Clientes.Fec_Ini),DATEPART(MONTH, devoluciones_Clientes.Fec_Ini),Articulos.Cod_Art,Articulos.Nom_Art")
			loComandoSeleccionar.Appendline("  ")

			loComandoSeleccionar.Appendline("UNION ALL")
			
			loComandoSeleccionar.Appendline("  ")
			loComandoSeleccionar.Appendline("--**********************************************************************")
			loComandoSeleccionar.Appendline("--**********************TOTAL FACTURAS COMPRA***************************")
			loComandoSeleccionar.Appendline("--**********************************************************************")

			loComandoSeleccionar.Appendline("SELECT																			")
			loComandoSeleccionar.AppendLine("		Articulos.Cod_Art								AS Cod_Art,				")
			loComandoSeleccionar.AppendLine("		Articulos.Nom_Art								As Nom_Art,				")
			loComandoSeleccionar.Appendline("		DATEPART(YEAR, Compras.Fec_Ini)					AS Año,					")
			loComandoSeleccionar.Appendline("		DATEPART(MONTH, Compras.Fec_Ini)				AS Mes,					")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10))						AS MontoNetoVenta,		")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10))						AS MontoDevueltoVenta,	")
			loComandoSeleccionar.Appendline("		ISNULL(SUM(Renglones_Compras.Can_Art1),0)		AS MontoNetoCompra,		")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10))						AS MontoDevueltoCompra	")
			loComandoSeleccionar.Appendline("FROM	Articulos")
			loComandoSeleccionar.Appendline("JOIN	Renglones_Compras ON Renglones_Compras.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.Appendline("JOIN	Compras ON Renglones_Compras.Documento = Compras.Documento AND Compras.Status IN ('Confirmado','Afectado','Procesado')")
			loComandoSeleccionar.Appendline("				AND Compras.Fec_Ini BETWEEN " & goServicios.mObtenerCampoFormatoSQL(ldFechaInicial) & " AND " & goServicios.mObtenerCampoFormatoSQL(ldFechaFinal))
			loComandoSeleccionar.Appendline("WHERE	" &  cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.Appendline("GROUP BY DATEPART(YEAR, Compras.Fec_Ini),DATEPART(MONTH, Compras.Fec_Ini),Articulos.Cod_Art,Articulos.Nom_Art")
			loComandoSeleccionar.Appendline("  ")

			loComandoSeleccionar.Appendline("UNION ALL")
			
			loComandoSeleccionar.Appendline("  ")
			loComandoSeleccionar.Appendline("--**********************************************************************")
			loComandoSeleccionar.Appendline("--*********************TOTAL DEVOLUCIONES COMPRA************************")
			loComandoSeleccionar.Appendline("--**********************************************************************")

			loComandoSeleccionar.Appendline("SELECT																				")
			loComandoSeleccionar.AppendLine("		Articulos.Cod_Art									AS Cod_Art,				")
			loComandoSeleccionar.AppendLine("		Articulos.Nom_Art									As Nom_Art,				")
			loComandoSeleccionar.Appendline("		DATEPART(YEAR, devoluciones_Proveedores.Fec_Ini)	AS Año,					")
			loComandoSeleccionar.Appendline("		DATEPART(MONTH, devoluciones_Proveedores.Fec_Ini)	AS Mes,					")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10))							AS MontoNetoVenta,		")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10))							AS MontoDevueltoVenta,	")
			loComandoSeleccionar.Appendline("		CAST(0 AS Decimal(28,10))							AS MontoNetoCompra,		")
			loComandoSeleccionar.Appendline("		ISNULL(SUM(Renglones_DProveedores.Can_Art1),0)		AS MontoDevueltoCompra	")
			loComandoSeleccionar.Appendline("FROM	Articulos")
			loComandoSeleccionar.Appendline("JOIN	Renglones_DProveedores ON Renglones_DProveedores.Cod_Art = Articulos.Cod_Art")
			loComandoSeleccionar.Appendline("JOIN	devoluciones_Proveedores ON Renglones_DProveedores.Documento = devoluciones_Proveedores.Documento AND devoluciones_Proveedores.Status IN ('Confirmado','Afectado','Procesado')")
			loComandoSeleccionar.Appendline("				AND devoluciones_Proveedores.Fec_Ini BETWEEN " & goServicios.mObtenerCampoFormatoSQL(ldFechaInicial) & " AND " & goServicios.mObtenerCampoFormatoSQL(ldFechaFinal))
			loComandoSeleccionar.Appendline("WHERE	" &  cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.Appendline("GROUP BY DATEPART(YEAR, devoluciones_Proveedores.Fec_Ini),DATEPART(MONTH, devoluciones_Proveedores.Fec_Ini),Articulos.Cod_Art,Articulos.Nom_Art")
			loComandoSeleccionar.Appendline("  ")

			loComandoSeleccionar.Appendline("")
			
			loComandoSeleccionar.Appendline("  ")
			loComandoSeleccionar.Appendline("SELECT ") 
			loComandoSeleccionar.Appendline("		#tmpMeses.Codigo AS Cod_Art,")
			loComandoSeleccionar.Appendline("		Articulos.Nom_Art AS Nom_Art,")
			loComandoSeleccionar.Appendline("		#tmpMeses.Renglon,															")
			loComandoSeleccionar.Appendline("		#tmpMeses.Año,																")
			loComandoSeleccionar.Appendline("		#tmpMeses.Nombre,															")
			loComandoSeleccionar.Appendline("		ISNULL(SUM(#TmpTemporal.MontoNetoVenta),0)			AS MontoNetoVenta,		")
			loComandoSeleccionar.Appendline("		ISNULL(SUM(#TmpTemporal.MontoDevueltoVenta),0)		AS MontoDevueltoVenta,	")
			loComandoSeleccionar.Appendline("		ISNULL(SUM(#TmpTemporal.MontoNetoCompra),0)			AS MontoNetoCompra,		")
			loComandoSeleccionar.Appendline("		ISNULL(SUM(#TmpTemporal.MontoDevueltoCompra),0)		AS MontoDevueltoCompra	")
			loComandoSeleccionar.Appendline("FROM #tmpMeses")
			loComandoSeleccionar.AppendLine("JOIN Articulos ON Articulos.Cod_Art collate SQL_Latin1_General_CP1_CI_AS= #tmpMeses.Codigo collate SQL_Latin1_General_CP1_CI_AS")
			loComandoSeleccionar.Appendline("FULL JOIN #TmpTemporal ON #TmpTemporal.Año = #tmpMeses.Año AND #TmpTemporal.Mes = #tmpMeses.Renglon")
			loComandoSeleccionar.Appendline("GROUP BY #tmpMeses.Codigo,Articulos.Nom_Art,#tmpMeses.Año,#tmpMeses.Nombre,Renglon")
			loComandoSeleccionar.Appendline("ORDER BY Año DESC,Renglon DESC")
			loComandoSeleccionar.Appendline("  ")

		
			'me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fVentas_Compras_uAños", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfVentas_Compras_uAños.ReportSource = loObjetoReporte

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
' MAT: 22/07/11: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' MAT: 03/08/11: Adición del Collate para las tablas Artículos y #tmpMeses
'-------------------------------------------------------------------------------------------'