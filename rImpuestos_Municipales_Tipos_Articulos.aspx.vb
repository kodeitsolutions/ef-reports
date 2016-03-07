'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rImpuestos_Municipales_Tipos_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rImpuestos_Municipales_Tipos_Articulos
   Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

       
	Try	
		
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
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
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
			Dim lcParametro14Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(14))
            Dim lcParametro14Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(14))


			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			
			   
			Dim loComandoSeleccionar As New StringBuilder() 	
		
				loComandoSeleccionar.AppendLine(" SELECT	CAST(Tipos_Articulos.imp_mun as XML) AS impuesto,")
				loComandoSeleccionar.AppendLine(" 			Tipos_Articulos.Cod_Tip AS Cod_Tip")
				loComandoSeleccionar.AppendLine(" INTO	#tmpImpuesto")
				loComandoSeleccionar.AppendLine(" FROM	Tipos_Articulos ")
				loComandoSeleccionar.AppendLine("")
				
				loComandoSeleccionar.AppendLine(" SELECT	#tmpImpuesto.Cod_Tip,")
				loComandoSeleccionar.AppendLine(" 			D.C.value('@codigo', 'Varchar(5000)') as codigo,")
				loComandoSeleccionar.AppendLine(" 			D.C.value('@nombre', 'Varchar(5000)') as nombre,")
				loComandoSeleccionar.AppendLine(" 			D.C.value('@impuesto', 'Varchar(15)') as impuesto")
				loComandoSeleccionar.AppendLine(" INTO	#tmpTemporal1")
				loComandoSeleccionar.AppendLine(" FROM #tmpImpuesto")
				loComandoSeleccionar.AppendLine(" CROSS APPLY impuesto.nodes('elementos/elemento') D(c)")
				loComandoSeleccionar.AppendLine(" ORDER BY Cod_Tip ASC")
				loComandoSeleccionar.AppendLine("")
	
				loComandoSeleccionar.AppendLine(" SELECT	#tmpTemporal1.*,")
				loComandoSeleccionar.AppendLine("			ISNULL(Impuestos_Municipales.Nom_Imp,'') AS Nom_Impuesto,")
				loComandoSeleccionar.AppendLine("			ISNULL(Impuestos_Municipales.Por_Imp,0) AS Por_Imp	")
				loComandoSeleccionar.AppendLine(" INTO #tmpTemporal2")
				loComandoSeleccionar.AppendLine(" FROM #tmpTemporal1")
				loComandoSeleccionar.AppendLine(" JOIN Impuestos_Municipales ON (Impuestos_Municipales.Cod_Imp = #tmpTemporal1.Impuesto)")
				loComandoSeleccionar.AppendLine(" GROUP BY #tmpTemporal1.Cod_Tip,#tmpTemporal1.codigo, #tmpTemporal1.nombre,#tmpTemporal1.impuesto,Impuestos_Municipales.Nom_Imp,Impuestos_Municipales.Por_Imp")
				loComandoSeleccionar.AppendLine("")
	
				loComandoSeleccionar.AppendLine(" SELECT	SUM(Renglones_Facturas.Mon_Bru) As MontoVenta,")
				loComandoSeleccionar.AppendLine("			Facturas.Cod_Suc As Cod_Suc,")
				loComandoSeleccionar.AppendLine("			Tipos_Articulos.Cod_Tip As Cod_Tip,")
				loComandoSeleccionar.AppendLine("			Tipos_Articulos.Nom_Tip  As Nom_Tip,")
				loComandoSeleccionar.AppendLine("			#tmpTemporal2.nombre As Nombre,")
				loComandoSeleccionar.AppendLine("			#tmpTemporal2.impuesto As Impuesto,")
				loComandoSeleccionar.AppendLine("			#tmpTemporal2.Nom_Impuesto As Nom_Impuesto,")
				loComandoSeleccionar.AppendLine("			#tmpTemporal2.Por_Imp As Por_Imp")
				loComandoSeleccionar.AppendLine(" INTO #tmpVentas")
				loComandoSeleccionar.AppendLine(" FROM Facturas")
				loComandoSeleccionar.AppendLine(" JOIN #tmpTemporal2 ON (#tmpTemporal2.Codigo = Facturas.Cod_Suc)")
				loComandoSeleccionar.AppendLine(" JOIN Renglones_Facturas On (Renglones_Facturas.Documento = Facturas.Documento)")
				loComandoSeleccionar.AppendLine(" JOIN Sucursales ON Sucursales.Cod_Suc = Facturas.Cod_Suc")
				loComandoSeleccionar.AppendLine(" JOIN Articulos ON (Articulos.Cod_Art = Renglones_Facturas.Cod_Art)")
				loComandoSeleccionar.AppendLine(" JOIN Tipos_Articulos ON (Tipos_Articulos.Cod_Tip = Articulos.Cod_Tip AND Tipos_Articulos.Cod_Tip = #tmpTemporal2.Cod_Tip)")
				loComandoSeleccionar.AppendLine(" WHERE Facturas.Fec_Ini BETWEEN " & lcParametro0Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro0Hasta)
				loComandoSeleccionar.AppendLine("		AND Facturas.Documento BETWEEN " & lcParametro1Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro1Hasta)
				loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Art BETWEEN " & lcParametro2Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro2Hasta)
				loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Dep BETWEEN " & lcParametro3Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro3Hasta)
				loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Sec BETWEEN " & lcParametro4Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro4Hasta)
				loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Tip BETWEEN " & lcParametro5Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro5Hasta)
				loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Cla BETWEEN " & lcParametro6Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro6Hasta)
				loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Mar BETWEEN " & lcParametro7Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro7Hasta)
				loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Suc BETWEEN " & lcParametro8Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro8Hasta)
		
			If lcParametro10Desde = "Igual" Then
				loComandoSeleccionar.AppendLine(" 		AND Facturas.Cod_Rev BETWEEN " & lcParametro9Desde)
			Else
				loComandoSeleccionar.AppendLine(" 		AND Facturas.Cod_Rev NOT BETWEEN " & lcParametro9Desde)
			End If

				loComandoSeleccionar.AppendLine("       AND " & lcParametro9Hasta)
				loComandoSeleccionar.AppendLine("       AND Facturas.Status   IN (" & lcParametro11Desde & ")")
				loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Mon BETWEEN " & lcParametro12Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro12Hasta)
				loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Ven BETWEEN " & lcParametro13Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro13Hasta)
				loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Cli BETWEEN " & lcParametro14Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro14Hasta)
				loComandoSeleccionar.AppendLine(" GROUP BY Facturas.Cod_Suc,Tipos_Articulos.Cod_Tip,Tipos_Articulos.Nom_Tip,#tmpTemporal2.nombre,#tmpTemporal2.impuesto,#tmpTemporal2.Nom_Impuesto,#tmpTemporal2.Por_Imp")
				loComandoSeleccionar.AppendLine("")
				
		
				loComandoSeleccionar.AppendLine("")
				loComandoSeleccionar.AppendLine(" SELECT	SUM(Renglones_dclientes.Mon_Bru) As MontoDevolucion,")
				loComandoSeleccionar.AppendLine("			devoluciones_clientes.Cod_Suc As Cod_Suc,")
				loComandoSeleccionar.AppendLine("			Tipos_Articulos.Cod_Tip As Cod_Tip,")
				loComandoSeleccionar.AppendLine("			#tmpTemporal2.nombre As Nombre,")
				loComandoSeleccionar.AppendLine("			#tmpTemporal2.impuesto As Impuesto,")
				loComandoSeleccionar.AppendLine("			#tmpTemporal2.Nom_Impuesto As Nom_Impuesto,")
				loComandoSeleccionar.AppendLine("			#tmpTemporal2.Por_Imp As Por_Imp")
				loComandoSeleccionar.AppendLine(" INTO #tmpDevoluciones")
				loComandoSeleccionar.AppendLine(" FROM devoluciones_clientes")
				loComandoSeleccionar.AppendLine(" JOIN #tmpTemporal2 ON (#tmpTemporal2.Codigo = devoluciones_clientes.Cod_Suc)")
				loComandoSeleccionar.AppendLine(" JOIN Renglones_dclientes On (Renglones_dclientes.Documento = devoluciones_clientes.Documento)")
				loComandoSeleccionar.AppendLine(" JOIN Sucursales ON Sucursales.Cod_Suc = devoluciones_clientes.Cod_Suc")
				loComandoSeleccionar.AppendLine(" JOIN Articulos ON Articulos.Cod_Art = Renglones_dclientes.Cod_Art")
				loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Art BETWEEN " & lcParametro2Desde)
				loComandoSeleccionar.AppendLine("					AND" & lcParametro2Hasta)
				loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Dep BETWEEN " & lcParametro3Desde)
				loComandoSeleccionar.AppendLine("					AND" & lcParametro3Hasta)
				loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Sec BETWEEN " & lcParametro4Desde)
				loComandoSeleccionar.AppendLine("					AND" & lcParametro4Hasta)
				loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Tip BETWEEN " & lcParametro5Desde)
				loComandoSeleccionar.AppendLine("					AND" & lcParametro5Hasta)
				loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Cla BETWEEN " & lcParametro6Desde)
				loComandoSeleccionar.AppendLine("					AND" & lcParametro6Hasta)
				loComandoSeleccionar.AppendLine("					AND Articulos.Cod_Mar BETWEEN " & lcParametro7Desde)
				loComandoSeleccionar.AppendLine("					AND" & lcParametro7Hasta)						
				loComandoSeleccionar.AppendLine(" JOIN Tipos_Articulos ON (Tipos_Articulos.Cod_Tip = Articulos.Cod_Tip AND Tipos_Articulos.Cod_Tip = #tmpTemporal2.Cod_Tip)")
				loComandoSeleccionar.AppendLine(" WHERE devoluciones_clientes.Fec_Ini BETWEEN " & lcParametro0Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro0Hasta)
				loComandoSeleccionar.AppendLine("		AND devoluciones_clientes.Cod_Suc BETWEEN " & lcParametro8Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro8Hasta)
		
			If lcParametro10Desde = "Igual" Then
				loComandoSeleccionar.AppendLine(" 		AND devoluciones_clientes.Cod_Rev BETWEEN " & lcParametro9Desde)
			Else
				loComandoSeleccionar.AppendLine(" 		AND devoluciones_clientes.Cod_Rev NOT BETWEEN " & lcParametro9Desde)
			End If

				loComandoSeleccionar.AppendLine("       AND " & lcParametro9Hasta)
				loComandoSeleccionar.AppendLine("       AND devoluciones_clientes.Status   IN (" & lcParametro11Desde & ")")
				loComandoSeleccionar.AppendLine("		AND devoluciones_clientes.Cod_Mon BETWEEN " & lcParametro12Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro12Hasta)
				loComandoSeleccionar.AppendLine("		AND devoluciones_clientes.Cod_Ven BETWEEN " & lcParametro13Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro13Hasta)
				loComandoSeleccionar.AppendLine("		AND devoluciones_clientes.Cod_Cli BETWEEN " & lcParametro14Desde)
				loComandoSeleccionar.AppendLine("		AND" & lcParametro14Hasta)
				loComandoSeleccionar.AppendLine(" GROUP BY devoluciones_clientes.Cod_Suc,Tipos_Articulos.Cod_Tip,#tmpTemporal2.nombre,#tmpTemporal2.impuesto,#tmpTemporal2.Nom_Impuesto,#tmpTemporal2.Por_Imp")
				loComandoSeleccionar.AppendLine("")
				
				
				loComandoSeleccionar.AppendLine(" SELECT ")
				loComandoSeleccionar.AppendLine("		#tmpVentas.Cod_Suc As Cod_Suc,")
				loComandoSeleccionar.AppendLine("		#tmpVentas.Cod_Tip As Cod_Tip,")
				loComandoSeleccionar.AppendLine("		#tmpVentas.Nom_Tip As Nom_Tip,")
				loComandoSeleccionar.AppendLine("		#tmpVentas.nombre As Nombre,")
				loComandoSeleccionar.AppendLine("		#tmpVentas.impuesto As Impuesto,")
				loComandoSeleccionar.AppendLine("		#tmpVentas.Nom_Impuesto As Nom_Impuesto,")
				loComandoSeleccionar.AppendLine("		#tmpVentas.Por_Imp As Por_Imp,")
				loComandoSeleccionar.AppendLine("		ISNULL(#tmpVentas.MontoVenta,0) AS MontoVenta,")
				loComandoSeleccionar.AppendLine("		ISNULL(#tmpDevoluciones.MontoDevolucion,0) AS MontoDevolucion")
				loComandoSeleccionar.AppendLine(" FROM #tmpVentas")
				loComandoSeleccionar.AppendLine(" LEFT JOIN #tmpDevoluciones ON #tmpDevoluciones.Cod_Suc = #tmpVentas.Cod_Suc")
				loComandoSeleccionar.AppendLine("			AND #tmpDevoluciones.Cod_Tip = #tmpVentas.Cod_Tip")
				loComandoSeleccionar.AppendLine("			AND #tmpDevoluciones.Impuesto = #tmpVentas.Impuesto")
				loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)
		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rImpuestos_Municipales_Tipos_Articulos", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrImpuestos_Municipales_Tipos_Articulos.ReportSource = loObjetoReporte
			
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
' MAT:  10/05/11 : Codigo inicial
'-------------------------------------------------------------------------------------------'
