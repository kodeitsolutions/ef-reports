Imports System.Data
Imports cusAplicacion

Partial Class rSMS_Auditorias_Usuarios
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine("SELECT	auditorias.Cod_Usu,")
            lcComandoSeleccionar.AppendLine("		Factory_Global.dbo.usuarios.nom_usu,")
            lcComandoSeleccionar.AppendLine("		COUNT(auditorias.Cod_Usu) AS SMS")
            lcComandoSeleccionar.AppendLine("FROM auditorias")
            lcComandoSeleccionar.AppendLine("	JOIN Factory_Global.dbo.usuarios ON auditorias.cod_usu collate database_default = Factory_Global.dbo.usuarios.cod_usu collate database_default")
            lcComandoSeleccionar.AppendLine("WHERE auditorias.accion = 'SMS'")
            lcComandoSeleccionar.AppendLine("       AND auditorias.cod_usu BETWEEN" & lcParametro0Desde & " AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("       AND auditorias.Registro BETWEEN" & lcParametro1Desde & " AND " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("GROUP BY auditorias.Cod_Usu, Factory_Global.dbo.usuarios.nom_usu")
            lcComandoSeleccionar.AppendLine("ORDER BY auditorias.Cod_Usu")



            'lcComandoSeleccionar.AppendLine(" ORDER BY  Cod_Usu, Fecha2 ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rSMS_Auditorias_Usuarios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSMS_Auditorias_Usuarios.ReportSource = loObjetoReporte

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
' EAG: 02/09/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
