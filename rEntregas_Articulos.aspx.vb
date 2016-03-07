Imports System.Data
Partial Class rEntregas_Articulos 
    Inherits System.Web.UI.Page
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
		Dim lcComandoSelect As String

        lcComandoSelect = " SELECT  Articulos.Cod_Art, " _ 
								& " Articulos.Nom_art, " _
								& " Articulos.Cod_Mar, " _
								& " Articulos.Status, " _
								& " Entregas.Documento, " _
								& " Entregas.Cod_Cli, " _
								& " Entregas.Fec_Ini, " _
								& " Entregas.Cod_Ven, " _
								& " Renglones_Entregas.Renglon, " _
								& " Renglones_Entregas.Cod_Alm, " _
								& " Renglones_Entregas.Can_Art1, " _
								& " Renglones_Entregas.Cod_Uni, " _
								& " Renglones_Entregas.Precio1, " _
								& " Renglones_Entregas.Por_Des, " _
								& " Renglones_Entregas.Mon_Net " _
					& " From Articulos, " _
								& " Entregas, " _
								& " Renglones_Entregas, " _
								& " Clientes, " _
								& " Vendedores, " _
								& " Almacenes, " _
								& " Marcas " _
					& " WHERE Articulos.Cod_Art = Renglones_Entregas.Cod_Art " _
								& " And Renglones_Entregas.Documento = Entregas.Documento "  _
								& " And Articulos.Cod_Mar = Marcas.Cod_Mar " _
								& " And Entregas.Cod_Cli = Clientes.Cod_Cli "  _
								& " And Entregas.Cod_Ven = Vendedores.Cod_Ven" _								
								& " And Renglones_Entregas.Cod_Alm = Almacenes.Cod_Alm " _
								& " And Articulos.Cod_Art between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))  _
								& " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))  _
								& " And Entregas.Fec_Ini between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))  _
								& " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))  _
								& " And Clientes.Cod_Cli between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))  _
								& " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))  _
								& " And Vendedores.Cod_Ven between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))  _
								& " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))  _
								& " And Marcas.Cod_Mar  between" & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))  _
								& " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))  _
								& " And Articulos.Status between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))  _
								& " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))  _
								& " And Almacenes.Cod_Alm between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))  _
								& " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))  _
						& " ORDER BY Articulos.Cod_Art"
    


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rEntregas_Articulos", laDatosReporte)
            
            Me.crvrEntregas_Articulos.ReportSource =	 loObjetoReporte	


        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try
	End Sub

	Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

	loObjetoReporte.Close()
	
	End Sub
End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 25/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
