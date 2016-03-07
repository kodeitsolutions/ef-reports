'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rlistado_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class rlistado_Clientes
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		Try
	
		
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.Append("SELECT	Clientes.Cod_Cli, ")
            loComandoSeleccionar.Append("		Clientes.Nom_Cli, ")
            loComandoSeleccionar.Append("		Clientes.Rif, ")
            loComandoSeleccionar.Append("		Clientes.Telefonos, ")
            loComandoSeleccionar.Append("CASE")
            loComandoSeleccionar.Append("		WHEN Clientes.Status = 'A' THEN 'Activo'")
            loComandoSeleccionar.Append("		WHEN Clientes.Status = 'I' THEN 'Inactivo'")
            loComandoSeleccionar.Append("		WHEN Clientes.Status = 'S' THEN 'Suspendido'")
            loComandoSeleccionar.Append("END As Status")
            loComandoSeleccionar.Append("FROM	Clientes")
             loComandoSeleccionar.AppendLine("     Clientes.Cod_Cli Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("     AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("     AND Clientes.Status In (" & lcParametro1Desde & ")")
            loComandoSeleccionar.Append("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rlistado_Clientes", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
	
            Me.crvrlistado_Clientes.ReportSource = loObjetoReporte	


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
' MAT: 10/03/11: Codigo inicial 
'-------------------------------------------------------------------------------------------'
