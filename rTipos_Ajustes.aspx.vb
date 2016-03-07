'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTipos_Ajustes"
'-------------------------------------------------------------------------------------------'
Partial Class rTipos_Ajustes
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        	   
	   Try
	   
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()
		
			 loComandoSeleccionar.AppendLine("SELECT		UPPER(Cod_Tip) as Cod_Tip, " )
			 loComandoSeleccionar.AppendLine("				UPPER(Nom_Tip) as Nom_Tip, " )
			 loComandoSeleccionar.AppendLine("				Tipo, " )
			 loComandoSeleccionar.AppendLine("				Status, " )
			 loComandoSeleccionar.AppendLine("(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Tipos_Ajustes " )
			 loComandoSeleccionar.AppendLine("FROM			Tipos_Ajustes " )
			 loComandoSeleccionar.AppendLine("WHERE			Cod_Tip between " & lcParametro0Desde)
			 loComandoSeleccionar.AppendLine("				And " & lcParametro0Hasta)
			 loComandoSeleccionar.AppendLine("				And Status IN (" & lcParametro1Desde & ")")
            'loComandoSeleccionar.AppendLine(" ORDER BY		Cod_Tip, " )
            'loComandoSeleccionar.AppendLine("				Nom_Tip")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTipos_Ajustes", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrTipos_Ajustes.ReportSource = loObjetoReporte
			
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
' MJP : 07/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP :  11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP :  14/07/08 : Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' CMS:  06/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'