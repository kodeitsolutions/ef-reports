'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTAlmacenes_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rTAlmacenes_Articulos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	   
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
	
			Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Traslados.Documento, ")
            loConsulta.AppendLine("			Traslados.Fec_Ini, ")
            loConsulta.AppendLine("			Traslados.Fec_Fin, ")
            loConsulta.AppendLine("			Traslados.Alm_Ori, ")
            loConsulta.AppendLine("			Traslados.Alm_Des, ")
            loConsulta.AppendLine("			Traslados.Cod_Mon, ")
            loConsulta.AppendLine("			Traslados.Tasa, ")
            loConsulta.AppendLine("			Traslados.Status, ")
            loConsulta.AppendLine("			Traslados.Comentario, ")
            loConsulta.AppendLine("			Traslados.Cod_Tra, ")
            loConsulta.AppendLine("			Transportes.Nom_Tra, ")
            loConsulta.AppendLine("			Renglones_Traslados.Renglon, ")
            loConsulta.AppendLine("			Renglones_Traslados.Cod_Art, ")
            loConsulta.AppendLine("			Renglones_Traslados.Cod_Alm, ")
            loConsulta.AppendLine("			Renglones_Traslados.Can_Art1, ")
            loConsulta.AppendLine("			Renglones_Traslados.Cod_Uni, ")
            loConsulta.AppendLine("			Articulos.Nom_Art ")
            loConsulta.AppendLine(" FROM		Traslados, ")
            loConsulta.AppendLine("			Renglones_Traslados, ")
            loConsulta.AppendLine("			Transportes, ")
            loConsulta.AppendLine("			Articulos ")
            loConsulta.AppendLine(" WHERE		Traslados.Documento				=	Renglones_Traslados.Documento ")
            loConsulta.AppendLine("			AND Renglones_Traslados.Cod_Art		=	Articulos.Cod_Art ")
            loConsulta.AppendLine("			AND Traslados.Cod_Tra				=	Transportes.Cod_Tra ")
            loConsulta.AppendLine("			AND Renglones_Traslados.Cod_Art		BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)
            loConsulta.AppendLine("			AND Traslados.Documento				BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("			AND " & lcParametro1Hasta)
            loConsulta.AppendLine("			AND Traslados.Fec_Ini				BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("			AND " & lcParametro2Hasta)
            loConsulta.AppendLine("			AND Traslados.Status				IN (" & lcParametro3Desde & ")")
            loConsulta.AppendLine("			AND Articulos.Cod_Dep				BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("			AND Articulos.Cod_Sec				BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("			AND " & lcParametro5Hasta)
            loConsulta.AppendLine("			AND Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("			AND " & lcParametro6Hasta)
            loConsulta.AppendLine("			AND Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine("			AND " & lcParametro7Hasta)
            loConsulta.AppendLine("			AND Articulos.Cod_Mar				BETWEEN " & lcParametro8Desde)
            loConsulta.AppendLine("			AND " & lcParametro8Hasta)
            loConsulta.AppendLine("      	AND Traslados.Cod_Rev BETWEEN " & lcParametro9Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro9Hasta)
            loConsulta.AppendLine("      	AND Traslados.Cod_Suc BETWEEN " & lcParametro10Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro10Hasta)
            loConsulta.AppendLine("         AND Articulos.Cod_Ubi    BETWEEN " & lcParametro11Desde)
            loConsulta.AppendLine("         AND " & lcParametro11Hasta)
            loConsulta.AppendLine("ORDER BY   Renglones_Traslados.Cod_Art, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rTAlmacenes_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTAlmacenes_Articulos.ReportSource =	 loObjetoReporte	

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
' JJD: 05/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 11/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  15/07/09: Metodo de Ordenamiento, Verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Se agregaro el filtro_ Ubicación
'-------------------------------------------------------------------------------------------'