'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCortes_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rCortes_Numeros 
     Inherits vis2Formularios.frmReporte

    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))	
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()
	

            loComandoSeleccionar.AppendLine(" SELECT      Cortes.Documento, ")
            loComandoSeleccionar.AppendLine("             Cortes.status, ")
            loComandoSeleccionar.AppendLine("             Cortes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("             Cortes.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("             Cortes.Contacto, ")
            loComandoSeleccionar.AppendLine("             Cortes.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("             Cortes.Tasa, ")
            loComandoSeleccionar.AppendLine("             Cortes.Comentario, ")
            loComandoSeleccionar.AppendLine("             Monedas.Nom_Mon, ")
            loComandoSeleccionar.AppendLine("             Almacenes.Nom_Alm ")
            loComandoSeleccionar.AppendLine(" From        Cortes, ")
            loComandoSeleccionar.AppendLine("             Monedas, ")
            loComandoSeleccionar.AppendLine("             Almacenes ")
			loComandoSeleccionar.AppendLine(" WHERE Cortes.Cod_Mon = Monedas.Cod_Mon " )
            loComandoSeleccionar.AppendLine("             And Cortes.Cod_Alm = Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("             And Cortes.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("             And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("             And Cortes.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("             And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("             And Cortes.Cod_Alm between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("             And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("             And Cortes.Cod_Mon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("             And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("             And Cortes.status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("      	      AND Cortes.Cod_Rev between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		      AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("             And Cortes.Cod_Suc between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("             And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Cortes.Documento " )


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCortes_Numeros", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCortes_Numeros.ReportSource =	 loObjetoReporte	

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
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' MVP:  28/04/09: Agregar combo estatus y estandarizacion
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' AAP:  02/07/09 : Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'