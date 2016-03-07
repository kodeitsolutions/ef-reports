Imports System.Data
Partial Class rCCobrar_Clientes_FSV1
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcComandoSelect As String

            lcComandoSelect = "  SELECT Cuentas_Cobrar.Documento, " _
              & " Cuentas_Cobrar.Cod_Tip, " _
              & " Cuentas_Cobrar.Fec_Ini, " _
              & " Cuentas_Cobrar.Fec_Fin, " _
              & " DATEDIFF(D, Cuentas_Cobrar.Fec_Fin, GetDate())   As  Dias, " _
              & " Cuentas_Cobrar.Cod_Cli, " _
              & " Clientes.Nom_Cli, " _
              & " Cuentas_Cobrar.Cod_Ven, " _
              & " Cuentas_Cobrar.Cod_Tra, " _
              & " Cuentas_Cobrar.Cod_Mon, " _
              & " Cuentas_Cobrar.Control, " _
              & " Cuentas_Cobrar.Mon_Bru, " _
              & " Cuentas_Cobrar.Mon_Imp1, " _
              & " Cuentas_Cobrar.Mon_Net, " _
              & " Cuentas_Cobrar.Mon_Sal * (case when Cuentas_cobrar.Tip_Doc='Credito' then -1 else 1 end) As Mon_Sal  " _
          & " From Clientes, " _
              & " Cuentas_Cobrar, " _
              & " Vendedores, " _
              & " Transportes, " _
              & " Monedas " _
          & " WHERE Cuentas_Cobrar.Cod_Cli      =   Clientes.Cod_Cli " _
              & " And Cuentas_Cobrar.Cod_Ven    =   Vendedores.Cod_Ven " _
              & " And Cuentas_Cobrar.Cod_Tra    =   Transportes.Cod_Tra " _
              & " And Cuentas_Cobrar.Cod_Mon    =   Monedas.Cod_Mon " _
              & " And Cuentas_Cobrar.Mon_Sal    <>  0 " _
              & " And Cuentas_Cobrar.Status     <>  'Anulado' " _
              & " ORDER BY  Cuentas_Cobrar.Cod_Cli,  Cuentas_Cobrar.Cod_Tip, Cuentas_Cobrar.Documento "


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSelect, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("rCCobrar_Clientes_FSV1", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCCobrar_Clientes_FSV1.ReportSource = loObjetoReporte


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
' JJD: 06/09/08: Programacion inicial
'-------------------------------------------------------------------------------------------'