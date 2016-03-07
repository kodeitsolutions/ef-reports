Imports System.Data
Partial Class rClientes_dCredito 
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
			loComandoSeleccionar.AppendLine("				Clientes.Mon_Cre, " )
			loComandoSeleccionar.AppendLine("				Clientes.Dia_Cre, " )
			loComandoSeleccionar.AppendLine("				Clientes.Pro_Pag, " )
			loComandoSeleccionar.AppendLine("				Clientes.Por_Des, " )
			loComandoSeleccionar.AppendLine("				Clientes.Hor_Caj, " )
			loComandoSeleccionar.AppendLine("				Clientes.Mon_Sal, " )
			loComandoSeleccionar.AppendLine("				Clientes.Comentario, " )
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
			loComandoSeleccionar.AppendLine(" 				AND Tipos_Clientes.Cod_Tip between " & lcParametro2Desde)
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

          loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rClientes_dCredito", laDatosReporte)
          
            
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrClientes_dCredito.ReportSource = loObjetoReporte
            

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
' MVP:  11/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  31/03/09: Estandarizacion de codigo y ajustes al diseño.
'-------------------------------------------------------------------------------------------'