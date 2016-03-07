Imports System.Data
Partial Class rProveedores_pZona

    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect As String

        lcComandoSelect = "SELECT	Proveedores.Cod_Pro, " _
         & "Proveedores.Nom_Pro, " _
         & "Zonas.Cod_Zon " _
         & "FROM Proveedores, Zonas " _
         & "WHERE Proveedores.Cod_Zon = Zonas.Cod_Zon " _
         & " And Zonas.Cod_Zon between '" & cusAplicacion.goReportes.paParametrosIniciales(0) & "'" _
         & " And '" & cusAplicacion.goReportes.paParametrosFinales(0) & "'" _
         & " ORDER BY Cod_Pro"

        Try


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

            Me.crvrProveedores_pZona.ReportSource = cusAplicacion.goReportes.mCargarReporte("rProveedores_pZona", laDatosReporte)


        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub



End Class
