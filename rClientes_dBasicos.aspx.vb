Imports System.Data
Partial Class rClientes_dBasicos 
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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("SELECT			Clientes.Cod_Cli, " )
			loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli, " )
            loComandoSeleccionar.AppendLine("				Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("				Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("				Clientes.Fax, ")
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Tip, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Zon, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Cla, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Ven " )
			loComandoSeleccionar.AppendLine("FROM			Clientes, " )
			loComandoSeleccionar.AppendLine("				Tipos_Clientes, " ) 
			loComandoSeleccionar.AppendLine("				Zonas, " ) 
			loComandoSeleccionar.AppendLine("				Clases_Clientes, " ) 
			loComandoSeleccionar.AppendLine("				Vendedores " ) 
			loComandoSeleccionar.AppendLine("WHERE			Clientes.Cod_Tip = Tipos_Clientes.Cod_Tip " )
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Zon = Zonas.Cod_Zon "  )
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Cla = Clases_Clientes.Cod_Cla "  )
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Ven = Vendedores.Cod_Ven "  )
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Cli between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Clientes.status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Tip between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Zon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Cla between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Ven between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			'loComandoSeleccionar.AppendLine(" ORDER BY		Clientes.Cod_Cli, Clientes.Nom_Cli")


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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rClientes_dBasicos", laDatosReporte)
			            
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrClientes_dBasicos.ReportSource =	 loObjetoReporte	


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
' MVP:  10/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  11/07/08: Adición de loObjetoReporte para eliminar los archivos temp en Uranus
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  31/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  31/03/09: Verificación de registros
'-------------------------------------------------------------------------------------------'