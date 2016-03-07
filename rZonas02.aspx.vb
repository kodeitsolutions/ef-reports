Imports System.Data
Partial Class rZonas02
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect as String
	 
	    lcComandoSelect = "SELECT	Cod_Zon, "	_ 
						& "Nom_Zon, "			_
						& "status, "		_
						& "nivel as numero " _
						& "FROM	 Zonas "		_
						& "WHERE Cod_Zon between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'"	_
						&				" And '" & cusAplicacion.goReportes.paParametrosFinales(0)	& "'"	_						 
						& " ORDER BY Cod_Zon"
	
		Try 

		
			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

			Me.crvrrZonas02.ReportSource = cusAplicacion.goReportes.mCargarReporte("rZonas02", laDatosReporte)

			
		Catch loExcepcion As Exception

			me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
																 "No se pudo Completar el Proceso: " & loExcepcion.Message, _
																  vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _ 
																  "auto", _
																  "auto")
		
		End Try
		
    End Sub
    
   
     
End Class
