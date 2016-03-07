'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Articulos.Cod_Art			AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art			AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Modelo			AS Modelo, ")
            loComandoSeleccionar.AppendLine("           Articulos.Talla				AS Talla, ")
            loComandoSeleccionar.AppendLine("           (CASE Articulos.Status ")
            loComandoSeleccionar.AppendLine("				WHEN 'A' THEN 'Activo'")
            loComandoSeleccionar.AppendLine("				WHEN 'I' THEN 'Inactivo'")
            loComandoSeleccionar.AppendLine("				ELSE 'Suspendido' ")
            loComandoSeleccionar.AppendLine("           END)						AS Status ")
            loComandoSeleccionar.AppendLine("FROM      Articulos ")
            loComandoSeleccionar.AppendLine("WHERE		Articulos.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Ubi BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos.ReportSource = loObjetoReporte

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
' MVP: 01/08/08: Cambios para multi idioma, mensaje de error y clase padre.					'
'-------------------------------------------------------------------------------------------'
' JJD: 02/02/09: Normalizacion de codigo. Inclusion de filtrado del Status.					'
'-------------------------------------------------------------------------------------------'
' CMS: 05/05/09: Ordenamiento																'
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Verificacion de registros y se agregaro el filtro_ Ubicación				'
'-------------------------------------------------------------------------------------------'
' RJG:  09/02/12: Normalización de Layout (tamaño de página).								'
'-------------------------------------------------------------------------------------------'
