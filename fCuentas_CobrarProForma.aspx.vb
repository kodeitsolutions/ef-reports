'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCuentas_CobrarProForma"
'-------------------------------------------------------------------------------------------'

Partial Class fCuentas_CobrarProForma

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
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Otr1 + Cuentas_Cobrar.Mon_Otr2 + Cuentas_Cobrar.Mon_Otr3)   AS  Mon_Otr, ")
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCuentas_CobrarProForma", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCuentas_CobrarProForma.ReportSource = loObjetoReporte

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
' CMS: 25/02/10: Codigo inicial
'-------------------------------------------------------------------------------------------'