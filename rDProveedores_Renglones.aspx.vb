Imports System.Data
Partial Class rDProveedores_Renglones
     Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		
		Dim lcParametro0Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro1Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro2Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
		Dim lcParametro2Hasta	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
		Dim lcParametro3Desde	As  String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
		Dim lcParametro4Desde	As  String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))


        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden



	Try
	
		Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT	Devoluciones_Proveedores.Documento, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Proveedores.Fec_Ini, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Proveedores.Fec_Fin, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Proveedores.Cod_Pro, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Proveedores.Cod_Ven, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Proveedores.Cod_Tra, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores.Renglon, ")
			loComandoSeleccionar.AppendLine("			Devoluciones_Proveedores.Cod_For, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores.Cod_Art, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores.Cod_Alm, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores.Precio1, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores.Can_Art1, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores.Cod_Uni, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores.Por_Imp1, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores.Por_Des, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores.Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("			Transportes.Nom_Tra, ")
            loComandoSeleccionar.AppendLine("			Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro ")
			loComandoSeleccionar.AppendLine(" FROM		Devoluciones_Proveedores, ")
			loComandoSeleccionar.AppendLine("			Renglones_DProveedores, ")
			loComandoSeleccionar.AppendLine("			Proveedores, ")
			loComandoSeleccionar.AppendLine("			Articulos, ")
			loComandoSeleccionar.AppendLine("			Vendedores, ")
			loComandoSeleccionar.AppendLine("			Formas_Pagos, ")
			loComandoSeleccionar.AppendLine("			Transportes ")
			loComandoSeleccionar.AppendLine(" WHERE		Devoluciones_Proveedores.Documento			=	Renglones_DProveedores.Documento ")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Proveedores.Cod_Pro		=	Proveedores.Cod_Pro ")
			loComandoSeleccionar.AppendLine("			And Renglones_DProveedores.Cod_Art			=	Articulos.Cod_Art ")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Proveedores.Cod_Ven		=	Vendedores.Cod_Ven ")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Proveedores.Cod_For		=	Formas_Pagos.Cod_For ")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Proveedores.Cod_Tra		=	Transportes.Cod_Tra ")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Proveedores.Documento		Between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			And Devoluciones_Proveedores.Fec_Ini		Between " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("			And Proveedores.Cod_Pro					    Between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("			And Devoluciones_Proveedores.Status		    IN (" & lcParametro3Desde & ")")
			loComandoSeleccionar.AppendLine("			And Devoluciones_Proveedores.Cod_rev		Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Devoluciones_Proveedores.Cod_Suc		Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   Devoluciones_Proveedores.Documento, " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY	Devoluciones_Proveedores.Documento")
           
        

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rDProveedores_Renglones", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte) 

            Me.crvrDProveedores_Renglones.ReportSource =	 loObjetoReporte	

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
' JJD: 13/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' AAP: 26/06/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'

