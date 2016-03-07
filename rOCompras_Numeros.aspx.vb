Imports System.Data
Partial Class rOCompras_Numeros
    
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
		Dim lcParametro4Desde	As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
		Dim lcParametro5Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
		Dim lcParametro5Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
		Dim lcParametro6Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
		Dim lcParametro6Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
		Dim lcParametro7Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
		Dim lcParametro7Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
        Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
		
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

	Try
	
		Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.appendline(" SELECT    Ordenes_Compras.Status, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Documento, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Mon_Bru, ")
            loComandoSeleccionar.appendline("           (Ordenes_Compras.Mon_Imp1 + Ordenes_Compras.Mon_Imp2 + Ordenes_Compras.Mon_Imp3)    AS Mon_Imp, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Mon_Net, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Mon_Sal, ")
            loComandoSeleccionar.appendline("           Ordenes_Compras.Cod_Tra, ")
            loComandoSeleccionar.appendline("           Transportes.Nom_Tra ")
            loComandoSeleccionar.appendline(" FROM      Ordenes_Compras, ")
            loComandoSeleccionar.appendline("           Proveedores, ")
            loComandoSeleccionar.appendline("           Transportes, ")
            loComandoSeleccionar.appendline("           Vendedores ")
            loComandoSeleccionar.appendline(" WHERE     Ordenes_Compras.Cod_Pro				=   Proveedores.Cod_Pro ")
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Ven			=   Vendedores.Cod_Ven ")
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Tra			=   Transportes.Cod_Tra ")
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Documento		Between " & lcParametro0Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro0Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Fec_Ini			Between " & lcParametro1Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro1Hasta)
            loComandoSeleccionar.appendline("           And Proveedores.Cod_Pro             Between " & lcParametro2Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro2Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Ven         Between " & lcParametro3Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro3Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Status          IN ( " & lcParametro4Desde & " ) ")
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Tra         Between " & lcParametro5Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro5Hasta)
            loComandoSeleccionar.appendline("           And Ordenes_Compras.Cod_Mon         Between " & lcParametro6Desde)
            loComandoSeleccionar.appendline("           And " & lcParametro6Hasta)
			loComandoSeleccionar.appendline("           And Ordenes_Compras.cod_rev         Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Compras.Cod_Suc         Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)

            'loComandoSeleccionar.appendline(" ORDER BY  Ordenes_Compras.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

			Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rOCompras_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOCompras_Numeros.ReportSource =	 loObjetoReporte	

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
' CMS: 18/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
