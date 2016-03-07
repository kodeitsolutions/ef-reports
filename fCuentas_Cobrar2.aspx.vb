Imports System.Data
Partial Class fCuentas_Cobrar2

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Clientes.Contacto, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Ven    =   Vendedores.Cod_Ven AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCuentas_Cobrar2", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCuentas_Cobrar2.ReportSource = loObjetoReporte

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
' JFP: 26/01/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
