Imports System.Data
Partial Class rProveedores
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	   
	Try	
	
		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
		Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
		Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
		Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
		Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
		Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
		Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
		Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
		Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
		
		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
		Dim loComandoSeleccionar As New StringBuilder()
	
		loComandoSeleccionar.AppendLine("SELECT	Cod_Pro, " )
		loComandoSeleccionar.AppendLine("Nom_Pro, " )
		loComandoSeleccionar.AppendLine("Proveedores.Status, " )
		loComandoSeleccionar.AppendLine("Proveedores.Cod_Tip, " )
		loComandoSeleccionar.AppendLine("Proveedores.Cod_Zon, " )
		loComandoSeleccionar.AppendLine("Proveedores.Cod_Cla, " )
		loComandoSeleccionar.AppendLine("Proveedores.Cod_Ven, " )
		loComandoSeleccionar.AppendLine("(Case When Proveedores.Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Proveedores " )
		loComandoSeleccionar.AppendLine("FROM	 Proveedores, " )
		loComandoSeleccionar.AppendLine(" Tipos_Proveedores, " )
		loComandoSeleccionar.AppendLine(" Zonas, " )
		loComandoSeleccionar.AppendLine(" Clases_Proveedores, " )
		loComandoSeleccionar.AppendLine(" Vendedores " )
		loComandoSeleccionar.AppendLine("WHERE Proveedores.Cod_Tip = Tipos_Proveedores.Cod_Tip " )
		loComandoSeleccionar.AppendLine(" And Proveedores.Cod_Zon = Zonas.Cod_Zon " )
		loComandoSeleccionar.AppendLine(" And Proveedores.Cod_Cla = Clases_Proveedores.Cod_Cla " )						  
		loComandoSeleccionar.AppendLine(" And Proveedores.Cod_Ven = Vendedores.Cod_Ven " )
		loComandoSeleccionar.AppendLine(" And Cod_Pro between " & lcParametro0Desde  )
		loComandoSeleccionar.AppendLine(" And " & lcParametro0Hasta )
		loComandoSeleccionar.AppendLine(" And Proveedores.Cod_Tip between " &  lcParametro1Desde )
		loComandoSeleccionar.AppendLine(" And " & lcParametro1Hasta )
		loComandoSeleccionar.AppendLine(" And Proveedores.Cod_Zon between " & lcParametro2Desde )
		loComandoSeleccionar.AppendLine(" And " & lcParametro2Hasta )
		loComandoSeleccionar.AppendLine(" And Proveedores.Cod_Cla between " & lcParametro3Desde )
		loComandoSeleccionar.AppendLine(" And " & lcParametro3Hasta )
		loComandoSeleccionar.AppendLine(" And Proveedores.Cod_Ven between " & lcParametro4Desde  )
		loComandoSeleccionar.AppendLine(" And " & lcParametro4Hasta )
		loComandoSeleccionar.AppendLine(" And Proveedores.Status IN (" & lcParametro5Desde & ")" )
		loComandoSeleccionar.AppendLine("ORDER BY proveedores."& lcOrdenamiento )


		   
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProveedores", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrProveedores.ReportSource = loObjetoReporte
			
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
' MJP   :  08/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP :  11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP :  14/07/08 : Filtros tipo, zonas, clase y vendedor agregados al reporte.
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'