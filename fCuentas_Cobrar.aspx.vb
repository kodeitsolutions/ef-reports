'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCuentas_Cobrar"
'-------------------------------------------------------------------------------------------'
Partial Class fCuentas_Cobrar
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loSeleccion As New StringBuilder()

            loSeleccion.AppendLine("SELECT	    Cuentas_Cobrar.Documento, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Cod_Tip, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Factura, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Control, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Cod_Cli, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Cod_Ven, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Status, ")
            loSeleccion.AppendLine("            CONVERT(NCHAR(10), Cuentas_Cobrar.Fec_Ini, 103)                                 AS	Fec_Ini, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Cod_Tra                                                          AS  Cod_Tra, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Cod_For                                                          AS  Cod_For, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Cod_Mon                                                          AS  Cod_Mon, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Tasa                                                             AS  Tasa, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Mon_Bru                                                          AS  Mon_Bru, ")
            loSeleccion.AppendLine("            (Cuentas_Cobrar.Mon_Imp1 + Cuentas_Cobrar.Mon_Imp2 + Cuentas_Cobrar.Mon_Imp3)   AS  Mon_Imp, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Mon_Net                                                          AS  Mon_Net, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Mon_Sal                                                          AS  Saldo, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Por_Rec                                                          AS  Por_Rec, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Mon_Rec                                                          AS  Mon_Rec, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Por_Des                                                          AS  Por_Des, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Mon_Des                                                          AS  Mon_Des, ")
            loSeleccion.AppendLine("            (Cuentas_Cobrar.Mon_Otr1 + Cuentas_Cobrar.Mon_Otr2 + Cuentas_Cobrar.Mon_Otr3)   AS  Mon_Otr, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Comentario                                                       AS  Comentario, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Tip_Ori, ")
            loSeleccion.AppendLine("            Cuentas_Cobrar.Doc_Ori, ")
            loSeleccion.AppendLine("            Clientes.Nom_Cli, ")
            loSeleccion.AppendLine("            Clientes.Rif, ")
            loSeleccion.AppendLine("            Clientes.Dir_Fis, ")
            loSeleccion.AppendLine("            Clientes.Telefonos                                                              AS  Telefonos, ")
            loSeleccion.AppendLine("            Clientes.Fax, ")
            loSeleccion.AppendLine("            Vendedores.Nom_Ven, ")
            loSeleccion.AppendLine("            Formas_Pagos.Nom_For, ")
            loSeleccion.AppendLine("            Transportes.Nom_Tra ")
            loSeleccion.AppendLine("FROM        Cuentas_Cobrar ")
            loSeleccion.AppendLine("    JOIN    Clientes ON Cuentas_Cobrar.Cod_Cli      = Clientes.Cod_Cli")
            loSeleccion.AppendLine("    JOIN    Vendedores ON Cuentas_Cobrar.Cod_Ven    = Vendedores.Cod_Ven ")
            loSeleccion.AppendLine("    JOIN    Formas_Pagos ON Cuentas_Cobrar.Cod_For  = Formas_Pagos.Cod_For")
            loSeleccion.AppendLine("    JOIN    Transportes ON Cuentas_Cobrar.Cod_Tra   = Transportes.Cod_Tra")
            loSeleccion.AppendLine("WHERE       " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loSeleccion.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCuentas_Cobrar", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCuentas_Cobrar.ReportSource = loObjetoReporte

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
' JJD: 24/01/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' RJG: 10/03/14: Comentarios y estandarizacion de código.
'-------------------------------------------------------------------------------------------'
