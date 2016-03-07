'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTrazabilidad_Compras"
'-------------------------------------------------------------------------------------------'
Partial Class rTrazabilidad_Compras

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()       
           
			
			loComandoSeleccionar.AppendLine("DECLARE @ldEmisionDesde		DATETIME")
			loComandoSeleccionar.AppendLine("DECLARE @ldEmisionHasta		DATETIME")
			loComandoSeleccionar.AppendLine("DECLARE @lcProveedorDesde		VARCHAR(10)")
			loComandoSeleccionar.AppendLine("DECLARE @lcProveedorHasta		VARCHAR(10)")
			loComandoSeleccionar.AppendLine("DECLARE @lcSucursalDesde		VARCHAR(10)")
			loComandoSeleccionar.AppendLine("DECLARE @lcSucursalHasta		VARCHAR(10)")
			loComandoSeleccionar.AppendLine("DECLARE @lcMonedaDesde			VARCHAR(10)")
			loComandoSeleccionar.AppendLine("DECLARE @lcMonedaHasta			VARCHAR(10)")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SET @ldEmisionDesde		= " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("SET @ldEmisionHasta		= " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("SET @lcProveedorDesde		= " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("SET @lcProveedorHasta		= " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("SET @lcSucursalDesde		= " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("SET @lcSucursalHasta		= " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("SET @lcMonedaDesde			= " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine("SET @lcMonedaHasta			= " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("-- Obtiene el 1º Nivel: Facturas de compra filtradas")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("SELECT		Compras.Cod_Pro				As Cod_Pro, ")
			loComandoSeleccionar.AppendLine("			Compras.Documento			As Compra, ")
			loComandoSeleccionar.AppendLine("			Renglones_Compras.Tip_Ori	AS TO_Compra, ")
			loComandoSeleccionar.AppendLine("			Renglones_Compras.Doc_Ori	AS DO_Compra,")
			loComandoSeleccionar.AppendLine("			Renglones_Compras.Ren_Ori	AS RO_Compra")
			loComandoSeleccionar.AppendLine("INTO		#tmpNivel_1")
			loComandoSeleccionar.AppendLine("FROM		Compras")
			loComandoSeleccionar.AppendLine("	JOIN	Renglones_Compras ")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Compras.Documento = Compras.Documento")
			loComandoSeleccionar.AppendLine("		AND	Compras.Fec_Ini BETWEEN @ldEmisionDesde AND @ldEmisionHasta")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("-- Obtiene el 2º nivel: documentos de origen de las Compras")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("SELECT		#tmpNivel_1.Cod_Pro					AS Cod_Pro,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_1.Compra					AS Compra,")
			loComandoSeleccionar.AppendLine("			Renglones_Recepciones.Documento		AS Recepcion,")
			loComandoSeleccionar.AppendLine("			Renglones_Recepciones.Tip_Ori		AS TO_Recepcion,")
			loComandoSeleccionar.AppendLine("			Renglones_Recepciones.Doc_Ori		AS DO_Recepcion,")
			loComandoSeleccionar.AppendLine("			Renglones_Recepciones.Ren_Ori		AS RO_Recepcion,")
			loComandoSeleccionar.AppendLine("			Renglones_oCompras.Documento		AS OrdenCompra,")
			loComandoSeleccionar.AppendLine("			Renglones_oCompras.Tip_Ori			AS TO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			Renglones_oCompras.Doc_Ori			AS DO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			Renglones_oCompras.Ren_Ori			AS RO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Documento	AS Presupuesto,")
			loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Req_Aso		AS DO_Presupuestos,")
			loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Documento	AS Requisicion")
			loComandoSeleccionar.AppendLine("INTO		#tmpNivel_2")
			loComandoSeleccionar.AppendLine("FROM		#tmpNivel_1")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Recepciones")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Recepciones.Documento = #tmpNivel_1.DO_Compra")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Recepciones.Renglon = #tmpNivel_1.RO_Compra")
			loComandoSeleccionar.AppendLine("		AND	#tmpNivel_1.TO_Compra = 'Recepciones'")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_oCompras ")
			loComandoSeleccionar.AppendLine("		ON	Renglones_oCompras.Documento = #tmpNivel_1.DO_Compra")
			loComandoSeleccionar.AppendLine("		AND	Renglones_oCompras.Renglon = #tmpNivel_1.RO_Compra")
			loComandoSeleccionar.AppendLine("		AND	#tmpNivel_1.TO_Compra = 'Ordenes_Compras'")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Presupuestos ")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Presupuestos.Documento = #tmpNivel_1.DO_Compra")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Presupuestos.Renglon = #tmpNivel_1.RO_Compra")
			loComandoSeleccionar.AppendLine("		AND	#tmpNivel_1.TO_Compra = 'Presupuestos'")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Requisiciones")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Requisiciones.Documento = #tmpNivel_1.DO_Compra")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Requisiciones.Renglon = #tmpNivel_1.RO_Compra")
			loComandoSeleccionar.AppendLine("		AND	#tmpNivel_1.TO_Compra = 'Requisiciones'")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpNivel_1")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("-- Obtiene el 3º nivel: documentos de origen de las recepciones")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("SELECT		#tmpNivel_2.Cod_Pro,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.Compra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.TO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.DO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.RO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.TO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.DO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.RO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.Presupuesto,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.DO_Presupuestos,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.Requisicion")
			loComandoSeleccionar.AppendLine("INTO		#tmpNivel_3")
			loComandoSeleccionar.AppendLine("FROM		#tmpNivel_2")
			loComandoSeleccionar.AppendLine("WHERE		#tmpNivel_2.Recepcion IS NULL")
			loComandoSeleccionar.AppendLine("UNION ALL")
			loComandoSeleccionar.AppendLine("SELECT		#tmpNivel_2.Cod_Pro					AS Cod_Pro,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.Compra					AS Compra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.Recepcion				AS Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.TO_Recepcion			AS TO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.DO_Recepcion			AS DO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_2.RO_Recepcion			AS RO_Recepcion,")
			loComandoSeleccionar.AppendLine("			Renglones_oCompras.Documento		AS OrdenCompra,")
			loComandoSeleccionar.AppendLine("			Renglones_oCompras.Tip_Ori			AS TO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			Renglones_oCompras.Doc_Ori			AS DO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			Renglones_oCompras.Ren_Ori			AS RO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Documento	AS Presupuesto,")
			loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Req_Aso		AS DO_Presupuestos,")
			loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Documento	AS Requisicion")
			loComandoSeleccionar.AppendLine("FROM		#tmpNivel_2	")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_oCompras")
			loComandoSeleccionar.AppendLine("		ON	Renglones_oCompras.Documento = #tmpNivel_2.DO_Recepcion")
			loComandoSeleccionar.AppendLine("		AND	Renglones_oCompras.Renglon = #tmpNivel_2.RO_Recepcion")
			loComandoSeleccionar.AppendLine("		AND	#tmpNivel_2.TO_Recepcion = 'ordenes_compras'")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Presupuestos")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Presupuestos.Documento = #tmpNivel_2.DO_Recepcion")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Presupuestos.Renglon = #tmpNivel_2.RO_Recepcion")
			loComandoSeleccionar.AppendLine("		AND	#tmpNivel_2.TO_Recepcion = 'presupuestos'")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Requisiciones")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Requisiciones.Documento = #tmpNivel_2.DO_Recepcion")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Requisiciones.Renglon = #tmpNivel_2.RO_Recepcion")
			loComandoSeleccionar.AppendLine("		AND	#tmpNivel_2.TO_Recepcion = 'requisiciones'")
			loComandoSeleccionar.AppendLine("WHERE		#tmpNivel_2.Recepcion IS NOT NULL")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpNivel_2")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("-- Obtiene el 4º nivel: documentos de origen de las Ordenes de Compra")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("SELECT		#tmpNivel_3.Cod_Pro,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.Compra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.TO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.DO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.RO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.TO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.DO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.RO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.Presupuesto,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.DO_Presupuestos,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.Requisicion")
			loComandoSeleccionar.AppendLine("INTO		#tmpNivel_4")
			loComandoSeleccionar.AppendLine("FROM		#tmpNivel_3")
			loComandoSeleccionar.AppendLine("WHERE		#tmpNivel_3.OrdenCompra IS NULL")
			loComandoSeleccionar.AppendLine("UNION ALL")
			loComandoSeleccionar.AppendLine("SELECT		#tmpNivel_3.Cod_Pro					AS Cod_Pro,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.Compra					AS Compra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.Recepcion				AS Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.TO_Recepcion			AS TO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.DO_Recepcion			AS DO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.RO_Recepcion			AS RO_Recepcion,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.OrdenCompra				AS OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.TO_OrdenCompra			AS TO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.DO_OrdenCompra			AS DO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			#tmpNivel_3.RO_OrdenCompra			AS RO_OrdenCompra,")
			loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Documento	AS Presupuesto,")
			loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Req_Aso		AS DO_Presupuestos,")
			loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Documento	AS Requisicion")
			loComandoSeleccionar.AppendLine("FROM		#tmpNivel_3	")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Presupuestos")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Presupuestos.Documento = #tmpNivel_3.DO_OrdenCompra")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Presupuestos.Renglon = #tmpNivel_3.RO_OrdenCompra")
			loComandoSeleccionar.AppendLine("		AND	#tmpNivel_3.TO_OrdenCompra = 'presupuestos'")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Requisiciones")
			loComandoSeleccionar.AppendLine("		ON	Renglones_Requisiciones.Documento = #tmpNivel_3.DO_OrdenCompra")
			loComandoSeleccionar.AppendLine("		AND	Renglones_Requisiciones.Renglon = #tmpNivel_3.RO_OrdenCompra")
			loComandoSeleccionar.AppendLine("		AND	#tmpNivel_3.TO_OrdenCompra = 'requisiciones'")
			loComandoSeleccionar.AppendLine("WHERE		#tmpNivel_3.OrdenCompra IS NOT NULL")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpNivel_3")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("-- Completa el 4º nivel: Agrega requisiciones asociadas a presupuestos")
			loComandoSeleccionar.AppendLine("--******************************************************************")
			loComandoSeleccionar.AppendLine("UPDATE	#tmpNivel_4")
			loComandoSeleccionar.AppendLine("SET		Requisicion = Origen.DO_Presupuestos")
			loComandoSeleccionar.AppendLine("FROM	(")
			loComandoSeleccionar.AppendLine("			SELECT		#tmpNivel_4.DO_Presupuestos,")
			loComandoSeleccionar.AppendLine("						#tmpNivel_4.Presupuesto")
			loComandoSeleccionar.AppendLine("			FROM		#tmpNivel_4	")
			loComandoSeleccionar.AppendLine("			WHERE		#tmpNivel_4.DO_Presupuestos IS NOT NULL")
			loComandoSeleccionar.AppendLine("					AND	#tmpNivel_4.DO_Presupuestos <> ''")
			loComandoSeleccionar.AppendLine("		) AS Origen")
			loComandoSeleccionar.AppendLine("WHERE	#tmpNivel_4.Presupuesto = Origen.Presupuesto")
			loComandoSeleccionar.AppendLine("	")
			loComandoSeleccionar.AppendLine("SELECT	#tmpNivel_4.Cod_Pro, ")
			loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro, ")
			loComandoSeleccionar.AppendLine("		#tmpNivel_4.Compra,  ")
			loComandoSeleccionar.AppendLine("		#tmpNivel_4.Recepcion,  ")
			loComandoSeleccionar.AppendLine("		#tmpNivel_4.OrdenCompra,  ")
			loComandoSeleccionar.AppendLine("		#tmpNivel_4.Presupuesto,  ")
			loComandoSeleccionar.AppendLine("		#tmpNivel_4.Requisicion ")
			loComandoSeleccionar.AppendLine("FROM	#tmpNivel_4")
			loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = #tmpNivel_4.Cod_Pro")
			loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro, Compra, Recepcion, OrdenCompra, Presupuesto, Requisicion")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpNivel_4")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

          
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTrazabilidad_Compras", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTrazabilidad_Compras.ReportSource = loObjetoReporte

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
' RJG: 09/01/12: Programacion inicial.														'
'-------------------------------------------------------------------------------------------'
