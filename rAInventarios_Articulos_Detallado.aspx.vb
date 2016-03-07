'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAInventarios_Articulos_Detallado"
'-------------------------------------------------------------------------------------------'
Partial Class rAInventarios_Articulos_Detallado
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
										

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            
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
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(14))
            Dim lcParametro14Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(14))
            Dim lcParametro15Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(15))
            Dim lcParametro15Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(15))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT		Articulos.Nom_art, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Status, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Ajustes.Comentario, ")
            loComandoSeleccionar.AppendLine(" 				Departamentos.Nom_Dep, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Ajustes.Renglon, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Ajustes.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Ajustes.Cod_Alm, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Ajustes.Cod_Uni, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Ajustes.Cos_Pro1, ")
            loComandoSeleccionar.AppendLine(" 				Renglones_Ajustes.Cos_Ult1, ")
            loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Ajustes.Tipo = 'Entrada' THEN ((1) * Renglones_Ajustes.Can_Art1)")
			loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Ajustes.Tipo = 'Salida' THEN ((-1) * Renglones_Ajustes.Can_Art1)")
			loComandoSeleccionar.AppendLine(" 			END AS Can_Art1,")
            loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Ajustes.Tipo = 'Entrada' THEN (Renglones_Ajustes.Cos_Pro1 * Renglones_Ajustes.Can_Art1)")
			loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Ajustes.Tipo = 'Salida' THEN (Renglones_Ajustes.Cos_Pro1 * (-1) * Renglones_Ajustes.Can_Art1)")
			loComandoSeleccionar.AppendLine(" 			END AS Tot_Pro,")
			loComandoSeleccionar.AppendLine(" 			CASE")
			loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Ajustes.Tipo = 'Entrada' THEN (Renglones_Ajustes.Cos_Ult1 * Renglones_Ajustes.Can_Art1)")
			loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Ajustes.Tipo = 'Salida' THEN (Renglones_Ajustes.Cos_Ult1 * (-1) * Renglones_Ajustes.Can_Art1)")
			loComandoSeleccionar.AppendLine(" 			END AS Tot_Ult,")
            loComandoSeleccionar.AppendLine(" 				Renglones_Ajustes.Tipo As Nom_Tip ")
            loComandoSeleccionar.AppendLine(" FROM			Articulos ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Ajustes ON  Articulos.Cod_Art	=	Renglones_Ajustes.Cod_Art")

            loComandoSeleccionar.AppendLine(" JOIN Ajustes ON  Ajustes.Documento    =	Renglones_Ajustes.Documento	")
            loComandoSeleccionar.AppendLine("				AND Ajustes.Tipo		IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Ajustes.Status		IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Ajustes.Fec_Ini				Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Ajustes.Documento			Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("      			And Ajustes.Cod_Rev between " & lcParametro14Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro14Hasta)
            loComandoSeleccionar.AppendLine(" JOIN Departamentos ON  Departamentos.Cod_Dep	=	Articulos.Cod_Dep")
            loComandoSeleccionar.AppendLine(" JOIN Tipos_Ajustes ON Tipos_Ajustes.Cod_Tip =   Renglones_Ajustes.Cod_Tip")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON Articulos.Cod_Pro  =	Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine(" WHERE			Renglones_Ajustes.Cod_Art	Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Ajustes.Cod_Alm	Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Ajustes.Cod_Tip	Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Dep	Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Sec	Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Cla	Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Tip	Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("      			And Articulos.Cod_Mar between " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("      			And Articulos.Cod_Pro between " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("      			And Articulos.Cod_Mon between " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine("      			And Articulos.Cod_Ubi between " & lcParametro15Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro15Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY		" & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAInventarios_Articulos_Detallado", laDatosReporte)
            
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrAInventarios_Articulos_Detallado.ReportSource = loObjetoReporte

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
' MAT: 31/01/11 Codigo inicial
'-------------------------------------------------------------------------------------------'
