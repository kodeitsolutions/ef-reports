'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "OnyxTD_rSeriales_Asignar"
'-------------------------------------------------------------------------------------------'
Partial Class OnyxTD_rSeriales_Asignar
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

        Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

        'Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        'Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
        Dim lcFiltroEstatus As String = ""

        If lcParametro2Desde.ToString() = "'Todos'" Then
            lcFiltroEstatus = "1=1"
        Else
            If lcParametro2Desde.ToString() = "'Disponibles'" Then
                lcFiltroEstatus = "Seriales.Doc_Sal = ''"
            Else
                lcFiltroEstatus = "Seriales.Doc_Sal <> ''"
            End If
        End If

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Seriales.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Articulos.Nom_Art,1,40)    AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Seriales.Serial, ")
            loComandoSeleccionar.AppendLine("           Seriales.Etapa, ")
            loComandoSeleccionar.AppendLine("           Seriales.Tip_Ent, ")
            loComandoSeleccionar.AppendLine("           Seriales.Doc_Ent, ")
            loComandoSeleccionar.AppendLine("           Seriales.Ren_Ent, ")
            loComandoSeleccionar.AppendLine("           Seriales.Alm_Ent, ")
            loComandoSeleccionar.AppendLine("           Seriales.Tip_Sal, ")
            loComandoSeleccionar.AppendLine("           Seriales.Doc_Sal, ")
            loComandoSeleccionar.AppendLine("           Seriales.Ren_Sal, ")
            loComandoSeleccionar.AppendLine("           Seriales.Alm_Sal ")
            loComandoSeleccionar.AppendLine(" FROM      Articulos, ")
            loComandoSeleccionar.AppendLine("           Seriales ")
            loComandoSeleccionar.AppendLine(" WHERE     Seriales.Cod_Art            =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           AND Seriales.Cod_Art        BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Seriales.Alm_Ent        BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND " & lcFiltroEstatus)
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes", 360)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("OnyxTD_rSeriales_Asignar", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvOnyxTD_rSeriales_Asignar.ReportSource = loObjetoReporte

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
' JJD: 14/08/14: Codigo inicial
'-------------------------------------------------------------------------------------------'
