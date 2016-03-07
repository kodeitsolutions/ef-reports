Imports System.Data
Partial Class rVendedores_dGenerales 
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
	

			loComandoSeleccionar.AppendLine("SELECT			Vendedores.Cod_Ven, " )
			loComandoSeleccionar.AppendLine("				Vendedores.Nom_Ven, " )
			loComandoSeleccionar.AppendLine("				Vendedores.Cod_Tip, " )
			loComandoSeleccionar.AppendLine("				Vendedores.Registro, " )
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Vendedores.Status = 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Vendedores.Status = 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine(" 				WHEN Vendedores.Status = 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine(" 			END AS Status,")
			loComandoSeleccionar.AppendLine("				Vendedores.Direccion, " )
			loComandoSeleccionar.AppendLine("				Vendedores.Telefonos, " )
			loComandoSeleccionar.AppendLine("				Vendedores.Por_Ven, " )
			loComandoSeleccionar.AppendLine("				Vendedores.Por_Cob, " )
			loComandoSeleccionar.AppendLine("				Vendedores.Comentario, " )
			loComandoSeleccionar.AppendLine("				Tipos_Vendedores.Nom_Tip, " )
			loComandoSeleccionar.AppendLine("				Zonas.Nom_Zon, " )
			loComandoSeleccionar.AppendLine("				Ciudades.Nom_Ciu, " )
			loComandoSeleccionar.AppendLine("				Paises.Nom_Pai " )
			loComandoSeleccionar.AppendLine("FROM			Vendedores " )
			loComandoSeleccionar.AppendLine("LEFT JOIN  Tipos_Vendedores ON (Vendedores.Cod_Tip = Tipos_Vendedores.Cod_Tip )" )
			loComandoSeleccionar.AppendLine("LEFT JOIN Zonas ON (Vendedores.Cod_Zon = Zonas.Cod_Zon )" )
			loComandoSeleccionar.AppendLine("LEFT JOIN Ciudades ON (Vendedores.Cod_Ciu = Ciudades.Cod_Ciu )")
			loComandoSeleccionar.AppendLine("LEFT JOIN Paises ON (Vendedores.Cod_Pai = Paises.Cod_Pai )")
			loComandoSeleccionar.AppendLine("WHERE			Vendedores.Cod_Ven between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Vendedores.status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Zon between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Pai between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Tip between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Ciu between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			
			  

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rVendedores_dGenerales", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrVendedores_dGenerales.ReportSource = loObjetoReporte
            

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
' MVP:  14/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' MAT:  01/02/11: Reestructuración del reporte. (no mostraba información alguna)
'-------------------------------------------------------------------------------------------'