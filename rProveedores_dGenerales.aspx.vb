Imports System.Data
Partial Class rProveedores_dGenerales 
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

            loComandoSeleccionar.AppendLine("SELECT	Proveedores.Cod_Pro, ")
			loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Rif, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Nit, " )
			loComandoSeleccionar.AppendLine("			(CASE WHEN Proveedores.Nacional = '1' THEN 'Si' ELSE 'No' END) As Nacional, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Cod_Tip, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Registro, " )
			loComandoSeleccionar.AppendLine("			(Case When Proveedores.Status = 'A' Then 'Activo' Else 'Inactivo' End) As Status, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Cod_Zon, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Cod_Cla, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Cod_Con, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Cod_Ciu, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Cod_Pai, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Cod_Pos, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Telefonos, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Fax, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Contacto, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Correo, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Web, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Cod_For, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Tip_Pag, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Pol_Com, " )
			loComandoSeleccionar.AppendLine("			Proveedores.Comentario, " )
			loComandoSeleccionar.AppendLine("			Tipos_Proveedores.Nom_Tip, " )
			loComandoSeleccionar.AppendLine("			Zonas.Nom_Zon, " )
			loComandoSeleccionar.AppendLine("			Clases_Proveedores.Nom_Cla, " )
			loComandoSeleccionar.AppendLine("			Conceptos.Nom_Con, " )
			loComandoSeleccionar.AppendLine("			Ciudades.Nom_Ciu, " )
			loComandoSeleccionar.AppendLine("			Paises.Nom_Pai, " )
			loComandoSeleccionar.AppendLine("			Formas_Pagos.Nom_For " )
			loComandoSeleccionar.AppendLine("FROM		Proveedores, " )
			loComandoSeleccionar.AppendLine("			Tipos_Proveedores, " )
			loComandoSeleccionar.AppendLine("			Zonas, " )
			loComandoSeleccionar.AppendLine("			Clases_Proveedores, " )
			loComandoSeleccionar.AppendLine("			Conceptos, " )
			loComandoSeleccionar.AppendLine("			Ciudades, " )
			loComandoSeleccionar.AppendLine("			Paises, " )
			loComandoSeleccionar.AppendLine("			Formas_Pagos " )
			loComandoSeleccionar.AppendLine("WHERE		Proveedores.Cod_Tip = Tipos_Proveedores.Cod_Tip " )
			loComandoSeleccionar.AppendLine("	AND 	Proveedores.Cod_Zon = Zonas.Cod_Zon " ) 
			loComandoSeleccionar.AppendLine("	AND 	Proveedores.Cod_Cla = Clases_Proveedores.Cod_Cla " ) 
			loComandoSeleccionar.AppendLine("	AND 	Proveedores.Cod_Con = Conceptos.Cod_Con " ) 
			loComandoSeleccionar.AppendLine("	AND 	Proveedores.Cod_Ciu = Ciudades.Cod_Ciu " ) 
			loComandoSeleccionar.AppendLine("	AND 	Proveedores.Cod_Pai = Paises.Cod_Pai " ) 
			loComandoSeleccionar.AppendLine("	AND 	Proveedores.Cod_For = Formas_Pagos.Cod_For " ) 
			loComandoSeleccionar.AppendLine("	AND 	Proveedores.Cod_Pro between " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine("	AND 	" & lcParametro0Hasta  )
			loComandoSeleccionar.AppendLine("	AND 	Proveedores.status IN (" & lcParametro1Desde & ")" )
			loComandoSeleccionar.AppendLine("	AND 	Zonas.Cod_Zon between " & lcParametro2Desde )
			loComandoSeleccionar.AppendLine("	AND 	" & lcParametro2Hasta )
			loComandoSeleccionar.AppendLine("	AND 	Paises.Cod_Pai between " & lcParametro3Desde )
			loComandoSeleccionar.AppendLine("	AND 	" & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("	AND 	Clases_Proveedores.Cod_Cla between " & lcParametro4Desde )
			loComandoSeleccionar.AppendLine("	AND 	" & lcParametro4Hasta )
			loComandoSeleccionar.AppendLine("	AND 	Tipos_Proveedores.Cod_Tip between " & lcParametro5Desde )
			loComandoSeleccionar.AppendLine("	AND 	" & lcParametro5Hasta )
            loComandoSeleccionar.AppendLine("ORDER BY proveedores." & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProveedores_dGenerales", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProveedores_dGenerales.ReportSource = loObjetoReporte
            

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
' MVP:  14/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' YJP:  21/04/09: Estandarizacion de codigos y correccion de campo estatus
'-------------------------------------------------------------------------------------------'