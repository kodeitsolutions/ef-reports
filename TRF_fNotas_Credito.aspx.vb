Imports System.Data
Partial Class TRF_fNotas_Credito
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT Cuentas_Cobrar.Factura              AS Factura,")
            loComandoSeleccionar.AppendLine("       Facturas.Control                    AS Control,")
            loComandoSeleccionar.AppendLine("       Clientes.Nom_Cli                    AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("       Clientes.Rif                        AS Rif,")
            loComandoSeleccionar.AppendLine("       Clientes.Dir_Fis                    AS Dir_Fis,")
            loComandoSeleccionar.AppendLine("       Clientes.Telefonos                  AS Telefonos,")
            loComandoSeleccionar.AppendLine("       Cuentas_Cobrar.Documento            AS Documento,")
            loComandoSeleccionar.AppendLine("       Cuentas_Cobrar.Fec_Ini              AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("       Cuentas_Cobrar.Mon_Bru              AS Mon_Bru,")
            loComandoSeleccionar.AppendLine("       Cuentas_Cobrar.Mon_Imp1             AS Mon_Imp1,")
            loComandoSeleccionar.AppendLine("       Cuentas_Cobrar.Mon_Net              AS Mon_Net,")
            loComandoSeleccionar.AppendLine("       Formas_Pagos.Nom_For                AS Nom_For,")
            loComandoSeleccionar.AppendLine("       Cuentas_Cobrar.Comentario           AS Comentario,")
            loComandoSeleccionar.AppendLine("       Renglones_Documentos.Cod_Art        AS Cod_Art,")
            loComandoSeleccionar.AppendLine("       Articulos.Nom_Art + SUBSTRING(Renglones_Documentos.Comentario,1,450)    AS  Nom_Art,")
            loComandoSeleccionar.AppendLine("       Renglones_Documentos.Can_Art        AS Can_Art1,")
            loComandoSeleccionar.AppendLine("       Renglones_Documentos.Cod_Uni        AS Cod_Uni,")
            loComandoSeleccionar.AppendLine("       Renglones_Documentos.Precio1        AS Precio1,")
            loComandoSeleccionar.AppendLine("       Renglones_Documentos.Mon_Net        As Neto,")
            loComandoSeleccionar.AppendLine("       Renglones_Documentos.Por_Imp        As Por_Imp")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Documentos ON Renglones_Documentos.Documento = Cuentas_Cobrar.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_Documentos.Cod_Tip =  Cuentas_Cobrar.Cod_Tip")
            loComandoSeleccionar.AppendLine("		AND Renglones_Documentos.Origen = 'Ventas'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Articulos  ON Articulos.Cod_Art   =   Renglones_Documentos.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("	JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Cuentas_Cobrar.Cod_For")
            loComandoSeleccionar.AppendLine("   JOIN Facturas ON Facturas.Documento = Cuentas_Cobrar.Factura")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("TRF_fNotas_Credito", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvTRF_fNotas_Credito.ReportSource = loObjetoReporte

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
' EAG: 15/08/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
