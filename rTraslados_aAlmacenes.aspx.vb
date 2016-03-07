Imports System.Data
Partial Class rTraslados_aAlmacenes
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect As String

        lcComandoSelect = "SELECT	Traslados.Documento, " _
     & "Renglones_Traslados.Cod_Art, " _
     & "Articulos.Nom_Art, " _
     & "Traslados.Fec_Ini, " _
     & "Traslados.Fec_Fin, " _
     & "Traslados.Alm_Ori, " _
     & "Traslados.Alm_Des, " _
     & "Renglones_Traslados.Can_Art1, " _
     & "Traslados.Status " _
     & "FROM	Traslados, Renglones_Traslados, Articulos   " _
     & "WHERE Traslados.Documento = Renglones_Traslados.Documento " _
     & " And Renglones_Traslados.Cod_Art = Articulos.Cod_Art  " _
     & " And Renglones_Traslados.Cod_Art between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
     & " And '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
     & " And Traslados.status between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
     & " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
     & " ORDER BY Traslados.Documento "
        '& " And Traslados.Alm_Ori between '" & cusAplicacion.goReportes.paParametrosIniciales(1) & "'" _
        '& " And '" & cusAplicacion.goReportes.paParametrosFinales(1) & "'" _
        '& " And Traslados.Alm_Des between '" & cusAplicacion.goReportes.paParametrosIniciales(2) & "'" _
        '& " And '" & cusAplicacion.goReportes.paParametrosFinales(2) & "'" _

        '& " And Traslados.fec_ini between '" & cusAplicacion.goReportes.paParametrosIniciales(4) & "'" _
        '& " And '" & cusAplicacion.goReportes.paParametrosFinales(4) & "'" _
        '& " And Traslados.fec_fin between '" & cusAplicacion.goReportes.paParametrosIniciales(5) & "'" _
        '& " And '" & cusAplicacion.goReportes.paParametrosFinales(5) & "'" _


        Try


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

            Me.crvrTraslados_aAlmacenes.ReportSource = cusAplicacion.goReportes.mCargarReporte("rTraslados_aAlmacenes", laDatosReporte)


        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub



End Class
