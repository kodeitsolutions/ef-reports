'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTarjetas_Cobradas"
'-------------------------------------------------------------------------------------------'
Partial Class rTarjetas_Cobradas
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.Appendline("SELECT Cobros.Documento, " )
			loComandoSeleccionar.Appendline("		Cobros.Fec_Ini, " )
			loComandoSeleccionar.Appendline("		Cobros.Cod_Ven, " )
			loComandoSeleccionar.Appendline("		Detalles_Cobros.Cod_Tar, " )
			loComandoSeleccionar.Appendline("		Detalles_Cobros.Num_Doc, " )
			loComandoSeleccionar.Appendline("		Detalles_Cobros.Mon_Net, " )
			loComandoSeleccionar.Appendline("		Tarjetas.Por_Com, " )
			loComandoSeleccionar.Appendline("		(Detalles_Cobros.Mon_Net * (Tarjetas.Por_Com / 100)) AS  Mon_Com, " )
			loComandoSeleccionar.Appendline("		Tarjetas.Por_Ret, " )
			loComandoSeleccionar.Appendline("		(Detalles_Cobros.Mon_Net * (Tarjetas.Por_Ret / 100)) AS  Mon_Ret, " )
			loComandoSeleccionar.Appendline("		Tarjetas.Cod_Tip " )
			loComandoSeleccionar.Appendline("FROM   Cobros INNER JOIN Detalles_Cobros " )
			loComandoSeleccionar.Appendline("ON Cobros.Documento  =   Detalles_Cobros.Documento, " )
			loComandoSeleccionar.Appendline("		Tarjetas " )
			loComandoSeleccionar.Appendline("WHERE  Tarjetas.Cod_Tar     =   Detalles_Cobros.Cod_Tar " )
			loComandoSeleccionar.Appendline("		AND Cobros.Fec_Ini BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.Appendline("		AND " & lcParametro0Hasta )
			loComandoSeleccionar.Appendline("		AND Detalles_Cobros.Cod_Caj BETWEEN " & lcParametro1Desde  )
			loComandoSeleccionar.Appendline("		AND " & lcParametro1Hasta  )
			loComandoSeleccionar.Appendline("		AND Detalles_Cobros.Cod_Mon BETWEEN " & lcParametro2Desde  )
			loComandoSeleccionar.Appendline("		AND " & lcParametro2Hasta )
            loComandoSeleccionar.AppendLine("		AND Cobros.Status IN ( " & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("    	AND " & lcParametro4Hasta)
            'loComandoSeleccionar.Appendline("ORDER BY  Cobros.Fec_Ini " )
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString , "curReportes")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rTarjetas_Cobradas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTarjetas_Cobradas.ReportSource =	 loObjetoReporte	

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
' JJD: 06/09/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  15/04/09: Cambios Estandarización de codigo. 
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  06/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'