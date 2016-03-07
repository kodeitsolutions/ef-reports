'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "reArticulos_cBarra"
'-------------------------------------------------------------------------------------------'
Partial Class reArticulos_cBarra
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lnParametro1Desde As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            If (lnParametro1Desde < 1) Then
                lnParametro1Desde = 1
            End If

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT   Articulos.Cod_Art			        AS Cod_Art, ")
            loConsulta.AppendLine("         Articulos.Nom_Art			        AS Nom_Art, ")
            loConsulta.AppendLine("         Articulos.Upc   		    	    AS Upc, ")
            loConsulta.AppendLine("        (CASE WHEN Articulos.Upc <> ''		 ")
            loConsulta.AppendLine("            THEN Articulos.Upc")
            loConsulta.AppendLine("            ELSE Articulos.Cod_Art")
            loConsulta.AppendLine("         END)		                        AS Barras, ")
            loConsulta.AppendLine("        (CASE Articulos.Status ")
            loConsulta.AppendLine("            WHEN 'A' THEN 'Activo'")
            loConsulta.AppendLine("            WHEN 'I' THEN 'Inactivo'")
            loConsulta.AppendLine("            ELSE 'Suspendido' ")
            loConsulta.AppendLine("         END)						        AS Status ")
            loConsulta.AppendLine("FROM     Articulos ")

            loConsulta.AppendLine("CROSS JOIN   (")
            loConsulta.AppendLine("                 SELECT '' AS X ")
            For n As Integer = 1 To lnParametro1Desde - 1
                loConsulta.AppendLine("             UNION ALL")
                loConsulta.AppendLine("                 SELECT ''")
            Next n

            loConsulta.AppendLine("             ) AS Y")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("WHERE    Articulos.Status IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("		AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Articulos.Cod_Dep BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("			AND " & lcParametro3Hasta)
            loConsulta.AppendLine("		AND Articulos.Cod_Sec BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("		AND Articulos.Cod_Mar BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("			AND " & lcParametro5Hasta)
            loConsulta.AppendLine("		AND Articulos.Cod_Pro BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("			AND " & lcParametro6Hasta)
            loConsulta.AppendLine("		AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine("			AND " & lcParametro7Hasta)
            loConsulta.AppendLine("		AND Articulos.Cod_Tip BETWEEN " & lcParametro8Desde)
            loConsulta.AppendLine("			AND " & lcParametro8Hasta)
            loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento)


            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("reArticulos_cBarra", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvreArticulos_cBarra.ReportSource = loObjetoReporte

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
' Bitacora
'-------------------------------------------------------------------------------------------'
' RJG:  22/04/13: Código Inicial.								                            '
'-------------------------------------------------------------------------------------------'
