Imports System.Data
Partial Class rVendedores_Clientes 
   Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()
	

			 loComandoSeleccionar.AppendLine("SELECT		Vendedores.Cod_Ven, " )
			 loComandoSeleccionar.AppendLine("				Vendedores.Nom_Ven, " )
			 loComandoSeleccionar.AppendLine("				Clientes.Telefonos, " )
			 loComandoSeleccionar.AppendLine("				Vendedores.Cod_Tip, " )
			 loComandoSeleccionar.AppendLine("				Clientes.Fax, " )
			 loComandoSeleccionar.AppendLine("				Clientes.Cod_Zon, " )
			 loComandoSeleccionar.AppendLine("				Clientes.Cod_Cli, " )
			 loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli, " )
			 loComandoSeleccionar.AppendLine("				Zonas.Nom_Zon, " )
			 loComandoSeleccionar.AppendLine("				Tipos_Vendedores.Nom_Tip " )
			 loComandoSeleccionar.AppendLine("FROM			Vendedores, " )
			 loComandoSeleccionar.AppendLine("				Tipos_Vendedores, " ) 
			 loComandoSeleccionar.AppendLine("				Zonas, " ) 
			 loComandoSeleccionar.AppendLine("				Clientes " )
			 loComandoSeleccionar.AppendLine("WHERE			Vendedores.Cod_Tip = Tipos_Vendedores.Cod_Tip " )
			 loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Zon = Zonas.Cod_Zon "  )
			 loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Ven = Clientes.Cod_Ven "  )
			 loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Ven between " & lcParametro0Desde)
			 loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			 loComandoSeleccionar.AppendLine(" 				AND Vendedores.status IN (" & lcParametro1Desde & ")")
			 loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Tip between " & lcParametro2Desde)
			 loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			 loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Zon between " & lcParametro3Desde)
			 loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			 loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			 'loComandoSeleccionar.AppendLine(" ORDER BY		Vendedores.Cod_Ven, Vendedores.Nom_Ven")

    

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rVendedores_Clientes", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrVendedores_Clientes.ReportSource =	 loObjetoReporte	


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
' MVP   :  10/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP   :  11/07/08 : Adición de loObjetoReporte para eliminar los archivos temp en Uranus
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  02/04/09: Estandarizacion de código y ajustes al diseño.
'-------------------------------------------------------------------------------------------'