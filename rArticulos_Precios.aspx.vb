Imports System.Data
Partial Class rArticulos_Precios
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect as String
	 
	    lcComandoSelect = "SELECT	Cod_Art, "	_ 
								  & "Nom_Art, " _
								  & "Precio1, " _
		     					  & "Precio2, "	_
								  & "Precio3, " _
								  & "Precio4, " _
								  & "Precio5 " _
						& "FROM	 Articulos "		_
						& "WHERE Cod_Art between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'"	_
								 & " And '" & cusAplicacion.goReportes.paParametrosFinales(0)	& "'"	_						 
      							 & " And Status between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
								 & " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
						& " ORDER BY Cod_Art"
	
		Try 

		
			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

			Me.crvrArticulos_Precios.ReportSource = cusAplicacion.goReportes.mCargarReporte("rArticulos_Precios", laDatosReporte)

			
		Catch loExcepcion As Exception

			me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
																 "No se pudo Completar el Proceso: " & loExcepcion.Message, _
																  vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _ 
																  "auto", _
																  "auto")
		
		End Try
		
    End Sub
    
   
     
End Class
