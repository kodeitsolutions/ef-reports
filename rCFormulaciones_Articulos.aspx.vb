'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCFormulaciones_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rCFormulaciones_Articulos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Ajustes.Status, ")
            loComandoSeleccionar.AppendLine("           Ajustes.Documento, ")
            loComandoSeleccionar.AppendLine("           Ajustes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes.Cos_Pro1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes.Cos_Ult1, ")
            loComandoSeleccionar.AppendLine("           (Renglones_Ajustes.Cos_Pro1 * Renglones_Ajustes.Can_Art1) AS Tot_Pro, ")
            loComandoSeleccionar.AppendLine("           (Renglones_Ajustes.Cos_Ult1 * Renglones_Ajustes.Can_Art1) AS Tot_Ult ")
            loComandoSeleccionar.AppendLine(" FROM      Articulos, ")
            loComandoSeleccionar.AppendLine("           Ajustes, ")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes ")
            loComandoSeleccionar.AppendLine(" WHERE     Ajustes.Documento		        =	Renglones_Ajustes.Documento	")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art			=	Renglones_Ajustes.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Ajustes.Tip_Ori             =   'formulas' ")
            loComandoSeleccionar.AppendLine("           And Renglones_Ajustes.Cod_Art	Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Ajustes.Documento			Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Ajustes.Fec_Ini			    Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Ajustes.Status			    IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           And Renglones_Ajustes.Cod_Alm	Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Ajustes.Cod_Rev between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro5Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Renglones_Ajustes.Cod_Art, ")
            'loComandoSeleccionar.AppendLine("           Ajustes.Fec_Ini, ")
            'loComandoSeleccionar.AppendLine("           Ajustes.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY    Renglones_Ajustes.Cod_Art, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(),"curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCFormulaciones_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCFormulaciones_Articulos.ReportSource = loObjetoReporte

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
' JJD: 17/01/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  20/07/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'