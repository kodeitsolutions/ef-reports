Imports System.Data
Partial Class rTipos_Containers
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT		Cod_Tip, ")
            loComandoSeleccionar.AppendLine("				Nom_Tip, ")
            loComandoSeleccionar.AppendLine("				Status ")
            loComandoSeleccionar.AppendLine(" FROM			Tipos_Containers ")
            loComandoSeleccionar.AppendLine(" WHERE			Cod_Tip between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				And Status IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY       " & lcOrdenamiento)

            Dim cad As String = loComandoSeleccionar.ToString()

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTipos_Containers", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTipos_Containers.ReportSource = loObjetoReporte

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
' RAC : 28/03/11 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' RAC : 29/03/11 : Modificacion en el Pie de Pagina del archivo rpt.
'-------------------------------------------------------------------------------------------'
