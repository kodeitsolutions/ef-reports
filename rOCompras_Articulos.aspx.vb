'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOCompras_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rOCompras_Articulos 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	 try 

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro12Desde As String = cusAplicacion.goReportes.paParametrosIniciales(12)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.appendline(" SELECT    Articulos.Cod_Art, ")
            loComandoSeleccionar.appendline("           Articulos.Nom_Art, ")
            loComandoSeleccionar.appendline("           Articulos.Cod_Dep, ")
            loComandoSeleccionar.appendline("           Articulos.Cod_Sec, ")
            loComandoSeleccionar.appendline("           Articulos.Cod_Mar, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Status, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Documento, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar.appendline("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar.appendline("           Renglones_OCompras.Renglon, ")
            loComandoSeleccionar.appendline("           Renglones_OCompras.Cod_Alm, ")
            'loComandoSeleccionar.appendline("           Renglones_OCompras.Can_Art1, ")

            Select Case lcParametro12Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("             Renglones_OCompras.Can_Art1, ")
                Case "Backorder"
                    loComandoSeleccionar.AppendLine("             Renglones_OCompras.Can_Pen1 AS Can_Art1, ")
                Case "Procesado"
                    loComandoSeleccionar.AppendLine("             (Renglones_OCompras.Can_Art1 - Renglones_OCompras.Can_Pen1) AS Can_Art1, ")
            End Select

            loComandoSeleccionar.appendline("           Renglones_OCompras.Cod_Uni, ")
            loComandoSeleccionar.appendline("           Renglones_OCompras.Precio1, ")
            loComandoSeleccionar.appendline("           Renglones_OCompras.Por_Des, ")
            loComandoSeleccionar.appendline("           Renglones_OCompras.Mon_Net ")
            loComandoSeleccionar.appendline(" FROM      Articulos, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras, ")
            loComandoSeleccionar.appendline("           Renglones_OCompras, ")
            loComandoSeleccionar.appendline("           Proveedores, ")
            loComandoSeleccionar.appendline("           Vendedores, ")
            loComandoSeleccionar.appendline("           Almacenes, ")
            loComandoSeleccionar.appendline("           Departamentos, ")
            loComandoSeleccionar.appendline("           Secciones, ")
            loComandoSeleccionar.appendline("           Marcas ")
            loComandoSeleccionar.appendline(" WHERE     Ordenes_Compras.Documento           =   Renglones_OCompras.Documento ")
            loComandoSeleccionar.appendline("           And Renglones_OCompras.Cod_Art      =   Articulos.Cod_Art " )
            loComandoSeleccionar.appendline("           And Articulos.Cod_Mar               =   Marcas.Cod_Mar ")
            loComandoSeleccionar.appendline("           And Articulos.Cod_Dep               =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.appendline("           And Secciones.Cod_Dep               =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.appendline("           And Articulos.Cod_Sec               =   Secciones.Cod_Sec ")
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Pro         =   Proveedores.Cod_Pro " )
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Ven         =   Vendedores.Cod_Ven ")								
            loComandoSeleccionar.AppendLine("           And Renglones_OCompras.Cod_Alm      =   Almacenes.Cod_Alm ")

            Select Case lcParametro12Desde
                Case "Backorder"
                    loComandoSeleccionar.AppendLine("             AND Renglones_OCompras.Can_Pen1 <> 0 ")
            End Select

            loComandoSeleccionar.appendline("           And Renglones_OCompras.Cod_Art      Between " & lcParametro0Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro0Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Fec_Ini         Between " & lcParametro1Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro1Hasta)
			loComandoSeleccionar.appendline("           And Renglones_OCompras.Cod_Art             Between " & lcParametro2Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro2Hasta)            
            loComandoSeleccionar.appendline("           And Proveedores.Cod_Pro             Between " & lcParametro3Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro3Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Ven         Between " & lcParametro4Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro4Hasta)
            loComandoSeleccionar.appendline("           And Articulos.Cod_Dep               Between " & lcParametro5Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro5Hasta)
            loComandoSeleccionar.appendline("           And Articulos.Cod_Sec               Between " & lcParametro6Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro6Hasta)
            loComandoSeleccionar.appendline("           And Articulos.Cod_Mar               Between " & lcParametro7Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro7Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Status          IN ( " & lcParametro8Desde & " ) ")
            loComandoSeleccionar.appendline("           And Renglones_OCompras.Cod_Alm      Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Compras.Cod_Rev      Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Compras.Cod_Suc      Between " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro11Hasta)
            'loComandoSeleccionar.appendline(" ORDER BY  Renglones_OCompras.Cod_Art, Ordenes_Compras.Fec_Ini, Ordenes_Compras.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY    Renglones_OCompras.Cod_Art, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rOCompras_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOCompras_Articulos.ReportSource =	 loObjetoReporte	

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
' JJD: 14/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:" Filtro "Revision:"
'-------------------------------------------------------------------------------------------'
' CMS:  13/07/09: Metode  de Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS: 22/07/09: Filtro BackOrder, lo conllevo al anexo del campo Can_Pen1,
'                 verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS: 19/03/10: se agrego el filtro cod_art
'-------------------------------------------------------------------------------------------'