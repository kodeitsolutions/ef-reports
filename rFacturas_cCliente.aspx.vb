Imports System.Data
Partial Class rFacturas_cCliente
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcComandoSelect As String

        Try

            lcComandoSelect = "SELECT	Cuentas_Cobrar.Documento as Compras, " _
              & "Cuentas_Cobrar.Cod_Cli, " _
              & "Cuentas_Cobrar.Fec_Fin, " _
              & "Cuentas_Cobrar.Status, " _
              & "(case when Cuentas_Cobrar.Status = 'Pagado' or Cuentas_Cobrar.Status = 'Afectado' then Renglones_Cobros.Documento else '0' end) as Condicion, " _
              & "(case when Cuentas_Cobrar.Status = 'Pendiente' or Cuentas_Cobrar.Status = 'Confirmado' then DATEDIFF(day,getdate(), Cuentas_Cobrar.Fec_Fin) else 0 end) as Vencimiento, " _
              & "(case when Cuentas_Cobrar.Status = 'Pagado' or Cuentas_Cobrar.Status = 'Afectado' then Cuentas_Cobrar.Mon_Sal else Cuentas_Cobrar.Mon_Net end) as Monto " _
              & "FROM	Cuentas_Cobrar, Renglones_Cobros,Cobros  " _
              & "WHERE Cuentas_Cobrar.Cod_Cli = Cobros.Cod_Cli " _
              & " And Renglones_Cobros.Documento = Cobros.Documento " _
              & " And Renglones_Cobros.Cod_Tip = 'FACT'  " _
              & " ORDER BY Cuentas_Cobrar.Documento "


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

            Me.crvrFacturas_cCliente.ReportSource = cusAplicacion.goReportes.mCargarReporte("rFacturas_cCliente", laDatosReporte)


        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub



End Class
