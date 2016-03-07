Imports System.Data
Partial Class rConstantes
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


	Try	
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()
		
		
			loComandoSeleccionar.AppendLine("SELECT			Cod_Con, " )
			loComandoSeleccionar.AppendLine("				Nom_Con, " )
			loComandoSeleccionar.AppendLine("				(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status, " )
			loComandoSeleccionar.AppendLine("				(Case tipo when 'C' then 'Caracter' when 'N' then 'Numérico' else 'Memo' end) as tipo, ")
			loComandoSeleccionar.AppendLine("				Val_Num, " )
			loComandoSeleccionar.AppendLine("				Val_Car, " )
			loComandoSeleccionar.AppendLine("				Val_Mem " )
			loComandoSeleccionar.AppendLine("FROM			Constantes " )
			loComandoSeleccionar.AppendLine("WHERE			Cod_Con between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("				AND status IN (" & lcParametro1Desde & ")")
            'loComandoSeleccionar.AppendLine("ORDER BY		Tipo " )
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

        
		
            Dim loServicios As New cusDatos.goDatos
            
            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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
   
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rConstantes", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
	   
			Me.crvrConstantes.ReportSource = loObjetoReporte
			
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
' MJP   :  09/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP :  11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP :  14/07/08 : Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' GCR :  24/03/09 : Estandarizacion de codigo y adicion de parametros
'-------------------------------------------------------------------------------------------'
' CMS:  31/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' MAT: 18/04/11: Mejora en la vista de Diseño
'-------------------------------------------------------------------------------------------'