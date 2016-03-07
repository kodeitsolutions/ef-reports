'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rClientes_Tipos"
'-------------------------------------------------------------------------------------------'
Partial Class rClientes_Tipos
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		Try
	
		
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.Append("SELECT	Cod_Tip, ")
            loComandoSeleccionar.Append("Nom_Tip, ")
            loComandoSeleccionar.Append("(Case when status = 'A' then 'Activo' Else 'Inactivo' End) As Status ")
            loComandoSeleccionar.Append("FROM	 Tipos_Clientes ")
            loComandoSeleccionar.Append("WHERE Cod_Tip between " & lcParametro0Desde)
            loComandoSeleccionar.Append(" And " & lcParametro0Hasta)
            loComandoSeleccionar.Append(" And status IN (" & lcParametro1Desde & ")")
            'loComandoSeleccionar.Append(" ORDER BY Cod_Tip, Nom_Tip")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rClientes_Tipos", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
	
            Me.crvrClientes_Tipos.ReportSource = loObjetoReporte	


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
' MVP:  08/07 08/: Codigo inicial 
'-------------------------------------------------------------------------------------------'
' MVP:  11/07/08: Adición de loObjetoReporte para eliminar los archivos temp en Uranus
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' YJP:  27/04/09: Agregar combo y estandarizacion
'-------------------------------------------------------------------------------------------'
' CMS:  06/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'