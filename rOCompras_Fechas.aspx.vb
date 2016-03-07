Imports System.Data
Partial Class rOCompras_Fechas
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
   
		Dim lcParametro0Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro1Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro2Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
		Dim lcParametro2Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
		Dim lcParametro3Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
		Dim lcParametro3Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
		Dim lcParametro4Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
		Dim lcParametro4Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
		Dim lcParametro5Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
		Dim lcParametro5Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
		Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
		Dim lcParametro7Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
		Dim lcParametro7Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
        Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
		'Dim lcParametro9Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        'Dim lcParametro9Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

	Try
	
		Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.appendline(" SELECT    Ordenes_Compras.Status, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Documento, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar.appendline("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Mon_Bru, ")
            loComandoSeleccionar.appendline("           (Ordenes_Compras.Mon_Imp1 + Ordenes_Compras.Mon_Imp2 + Ordenes_Compras.Mon_Imp3) AS Mon_Imp, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Mon_Net, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Mon_Sal, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Cod_Tra, ")
            loComandoSeleccionar.appendline("           Transportes.Nom_Tra ")
            loComandoSeleccionar.appendline(" FROM      Ordenes_Compras, ")
            loComandoSeleccionar.appendline("           Proveedores, ")
            loComandoSeleccionar.appendline("           Transportes, ")
            loComandoSeleccionar.appendline("           Vendedores ")
            loComandoSeleccionar.appendline(" WHERE     Ordenes_Compras.Cod_Pro                 =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Ven             =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Tra             =   Transportes.Cod_Tra ")
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Documento           Between " & lcParametro0Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro0Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Fec_Ini             Between " & lcParametro1Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro1Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Pro             Between " & lcParametro2Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro2Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Ven             Between " & lcParametro3Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro3Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Tra             Between " & lcParametro4Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro4Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Mon             Between " & lcParametro5Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro5Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Status              IN ( " & lcParametro6Desde & " ) ")
			loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_rev             Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Compras.Cod_Suc             Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            'loComandoSeleccionar.appendline(" ORDER BY  Ordenes_Compras.Fec_Ini, Ordenes_Compras.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY   Ordenes_Compras.Fec_Ini,  " & lcOrdenamiento)

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rOCompras_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOCompras_Fechas.ReportSource =	 loObjetoReporte	

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
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  20/07/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'