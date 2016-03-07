'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLInventarios_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rLInventarios_Numeros
     Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden


	     Try
	     
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))


            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT		Libres_Inventarios.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Inventarios.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Inventarios.Fec_Fin, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Inventarios.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Inventarios.Tasa, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Inventarios.Status, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Inventarios.Comentario, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Inventarios.Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Inventarios.Mon_Imp, ")
            loComandoSeleccionar.AppendLine(" 				Libres_Inventarios.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 				Monedas.Nom_Mon ")
            loComandoSeleccionar.AppendLine(" From			Libres_Inventarios, ")
            loComandoSeleccionar.AppendLine("				Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE		Libres_Inventarios.Cod_Mon = Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" 				AND Libres_Inventarios.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Libres_Inventarios.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Libres_Inventarios.Cod_Mon between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Libres_Inventarios.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("               AND Libres_Inventarios.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 		        AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY		Libres_Inventarios.Documento ")
 
            Dim loServicios As New cusDatos.goDatos
 
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
 
            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rLInventarios_Numeros", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLInventarios_Numeros.ReportSource =	 loObjetoReporte	

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
' JJD: 25/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  23/03/09: Estandarización de código y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  02/07/09 : Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'