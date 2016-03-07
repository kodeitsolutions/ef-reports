'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rClientes_dGenerales"
'-------------------------------------------------------------------------------------------'
Partial Class rClientes_dGenerales 
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
		    Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("SELECT			Clientes.Cod_Cli, " )
			loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli, " )
			loComandoSeleccionar.AppendLine("				Clientes.Rif, " )
			loComandoSeleccionar.AppendLine("				Clientes.Nit, " )
			loComandoSeleccionar.AppendLine("				(Case When Clientes.Fiscal = '1' Then 'Si' Else 'No' End) As Contribuyente, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Tip, " )
			loComandoSeleccionar.AppendLine("				Clientes.Registro, " )
			loComandoSeleccionar.AppendLine("				(Case When Clientes.Status = 'A' Then 'Activo' Else 'Inactivo' End) As Status, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Zon, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Cla, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Ven, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Con, " )
			loComandoSeleccionar.AppendLine("				Clientes.Dir_Fis, " )
			loComandoSeleccionar.AppendLine("				Clientes.Dir_Ent, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Ciu, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Pai, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_Pos, " )
			loComandoSeleccionar.AppendLine("				Clientes.Telefonos, " )
			loComandoSeleccionar.AppendLine("				Clientes.Fax, " )
			loComandoSeleccionar.AppendLine("				Clientes.Contacto, " )
			loComandoSeleccionar.AppendLine("				Clientes.Correo, " )
			loComandoSeleccionar.AppendLine("				Clientes.Web, " )
			loComandoSeleccionar.AppendLine("				Clientes.Cod_For, " )
			loComandoSeleccionar.AppendLine("				Clientes.Tip_Pag, " )
			loComandoSeleccionar.AppendLine("				Clientes.Pol_Com, " )
			loComandoSeleccionar.AppendLine("				Tipos_Clientes.Nom_Tip, " )
			loComandoSeleccionar.AppendLine("				Zonas.Nom_Zon, " )
			loComandoSeleccionar.AppendLine("				Clases_Clientes.Nom_Cla, " )
			loComandoSeleccionar.AppendLine("				Vendedores.Nom_Ven, " )
			loComandoSeleccionar.AppendLine("				Conceptos.Nom_Con, " )
			loComandoSeleccionar.AppendLine("				Ciudades.Nom_Ciu, " )
			loComandoSeleccionar.AppendLine("				Paises.Nom_Pai, " )
			loComandoSeleccionar.AppendLine("				Formas_Pagos.Nom_For " )
			loComandoSeleccionar.AppendLine("				FROM	 Clientes, " )
			loComandoSeleccionar.AppendLine("				Tipos_Clientes, " )
			loComandoSeleccionar.AppendLine("				Zonas, " )
			loComandoSeleccionar.AppendLine("				Clases_Clientes, " )
			loComandoSeleccionar.AppendLine("				Vendedores, " )
			loComandoSeleccionar.AppendLine("				Conceptos, " )
			loComandoSeleccionar.AppendLine("				Ciudades, " )
			loComandoSeleccionar.AppendLine("				Paises, " )
			loComandoSeleccionar.AppendLine("				Formas_Pagos " )
			loComandoSeleccionar.AppendLine("WHERE			Clientes.Cod_Tip = Tipos_Clientes.Cod_Tip " )
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Zon = Zonas.Cod_Zon " ) 
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Cla = Clases_Clientes.Cod_Cla " ) 
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Ven = Vendedores.Cod_Ven " ) 
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Con = Conceptos.Cod_Con " ) 
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Ciu = Ciudades.Cod_Ciu " ) 
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Pai = Paises.Cod_Pai " ) 
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_For = Formas_Pagos.Cod_For " ) 
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Cli between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Clientes.status IN (" & lcParametro1Desde &  ")")
			loComandoSeleccionar.AppendLine(" 				AND Vendedores.Cod_Ven between " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Zon between " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Pai between " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Cla between " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine(" 				AND Clientes.Cod_Tip between " & lcParametro6Desde)
			loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
     
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rClientes_dGenerales", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrClientes_dGenerales.ReportSource = loObjetoReporte
            

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
' MVP: 11/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR: 31/03/09: Estandarizacion de codigo y ajustes al diseño.
'-------------------------------------------------------------------------------------------'
' RJG: 06/05/15: Ajuste mejor de diseño.
'-------------------------------------------------------------------------------------------'
