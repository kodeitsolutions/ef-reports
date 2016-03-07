'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCiudades"
'-------------------------------------------------------------------------------------------'
Partial Class rCiudades 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			 loComandoSeleccionar.AppendLine("SELECT		Ciudades.Cod_Ciu, " )
			 loComandoSeleccionar.AppendLine("				Ciudades.Nom_Ciu, " )
			 loComandoSeleccionar.AppendLine("				Ciudades.Cod_Pai, " )
			 loComandoSeleccionar.AppendLine("				Paises.Nom_Pai, " )
			 loComandoSeleccionar.AppendLine("				(Case when Ciudades.status = 'A' then 'Activo' Else 'Inactivo' End) As Status " )
			 loComandoSeleccionar.AppendLine("FROM			Ciudades,Paises " )
			 loComandoSeleccionar.AppendLine("WHERE			Ciudades.Cod_Ciu between " & lcParametro0Desde)
			 loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
			 loComandoSeleccionar.AppendLine(" 				AND Ciudades.status IN (" & lcParametro1Desde & ")")
			 loComandoSeleccionar.AppendLine(" 				AND Ciudades.Cod_Pai = Paises.Cod_pai" )
            'loComandoSeleccionar.AppendLine(" ORDER BY 	Ciudades.Cod_Ciu, Ciudades.Nom_Ciu")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

    


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=   cusAplicacion.goReportes.mCargarReporte("rCiudades", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrCiudades.ReportSource =	loObjetoReporte	


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
' MVP:  07/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  11/07/08: Adición de loObjetoReporte para eliminar los archivos temp en Uranus
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  10/03/09: Estandarización de codigo	 y ajustes al diseño.
'-------------------------------------------------------------------------------------------'
' CMS:  06/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'