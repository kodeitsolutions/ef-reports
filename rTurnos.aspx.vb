Imports System.Data
Partial Class rTurnos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	turnos.cod_tur, ")
            loComandoSeleccionar.AppendLine("		turnos.nom_tur, ")
            loComandoSeleccionar.AppendLine("		turnos.tipo, ")
            loComandoSeleccionar.AppendLine("		turnos.status, ")
            loComandoSeleccionar.AppendLine("Case When Status = 'A' Then 'Activo' Else 'Inactivo' End as Status_turno, ")
            loComandoSeleccionar.AppendLine("		turnos.hor_ini, ")
            loComandoSeleccionar.AppendLine("		turnos.hor_fin ")
            loComandoSeleccionar.AppendLine("FROM	turnos ")
            loComandoSeleccionar.AppendLine(" WHERE	turnos.cod_tur between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" AND 	turnos.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" ORDER BY turnos." & lcOrdenamiento)


            'loComandoSeleccionar.AppendLine(" ORDER BY turnos.cod_tur, turnos.nom_tur" )

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTurnos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTurnos.ReportSource = loObjetoReporte


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
' YJP:  27/04/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
