Imports System.Data
Partial Class rBcotizaciones_Detalladas 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
   '-------------------------------------------------------------------------------------------'
   ' Filtro de status. Para Combos con listas fijas
   '-------------------------------------------------------------------------------------------'
	Dim lcStatus	As String = cusAplicacion.goReportes.paParametrosIniciales(3)
	Dim laStatus()	As String = lcStatus.Split(",")
	
	For lnContador As Integer = 0 To laStatus.Length -1
		laStatus(lnContador) = goServicios.mObtenerCampoFormatoSQL(laStatus(lnContador))
	Next lnContador
	lcStatus = String.Join(",", laStatus)
	'-------------------------------------------------------------------------------------------'
	
	Try
	
		Dim lcComandoSelect As String
		
		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            lcComandoSelect = "SELECT	Cotizaciones.Documento, " _
             & "Cotizaciones.Fec_Ini, " _              
             & "Cotizaciones.Comentario As Comentario_Encabezado, " _
             & "Cotizaciones.Cod_For, " _
             & "Cotizaciones.Cod_Tra, " _	             
             & "Clientes.Cod_Cli, " _
             & "Clientes.Nom_Cli, " _
             & "Articulos.Cod_Art, " _
             & "Articulos.Nom_Art, " _
             & "Renglones_Cotizaciones.Comentario, " _
             & "Renglones_Cotizaciones.Can_Art1, " _
             & "Renglones_Cotizaciones.Can_Pen1, " _
             & "Vendedores.Cod_Ven, " _
             & "Vendedores.Nom_Ven " _
          & "FROM	 Cotizaciones, " _
             & "Renglones_Cotizaciones, " _
             & "Clientes, " _
             & "Articulos, " _
             & "Vendedores, " _
             & "Formas_Pagos " _
          & "WHERE	Cotizaciones.Documento = Renglones_Cotizaciones.Documento " _
             & " And Cotizaciones.Cod_Cli = Clientes.Cod_Cli " _
             & " And Renglones_Cotizaciones.Cod_Art = Articulos.Cod_Art " _
             & " And Cotizaciones.Cod_Ven = Vendedores.Cod_Ven " _
             & " And Cotizaciones.Cod_For = Formas_Pagos.Cod_For " _
             & " And Cotizaciones.Documento Between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0)) _
             & " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0)) _
             & " And Cotizaciones.Fec_Ini between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuTipoRedondeo.KN_Inferior,2,goServicios.enuTipoRedondeoFecha.KN_InicioDelDia) _
             & " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuTipoRedondeo.KN_Superior,2,goServicios.enuTipoRedondeoFecha.KN_FinDelDia) _
             & " And Clientes.Cod_Cli between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2)) _
             & " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2)) _ 
             & " And Cotizaciones.Status IN (" & lcStatus & ")" _
             & " And Vendedores.Cod_Ven between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4)) _
             & " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4)) _
             & " And Formas_Pagos.Cod_For between " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5)) _
             & " And " & goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5)) _
          & " ORDER BY Vendedores.Cod_Ven, " & lcOrdenamiento

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rBcotizaciones_Detalladas", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrBcotizaciones_Detalladas.ReportSource =	 loObjetoReporte	

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
' MVP: 13/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 31/09/09: Se agregaron los campos comentario (del encabezado), cod_for, nivel
'				 Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'