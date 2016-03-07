'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFacturas_Ventas_FRO"
'-------------------------------------------------------------------------------------------'
Partial Class fFacturas_Ventas_FRO
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("Select	Facturas.Documento							AS Invoice,")
            loConsulta.AppendLine("		Clientes.nom_cli							AS Bill_Client,")
            loConsulta.AppendLine("		Clientes.Dir_fis							AS Bill_Address,")
            loConsulta.AppendLine("		Clientes.telefonos							AS Bill_Telephone,")
            loConsulta.AppendLine("		Clientes.rif							    AS Bill_rif,")
            loConsulta.AppendLine("     Case Clientes.Dir_Ent")
            loConsulta.AppendLine("			WHEN '' THEN 'Same'")
            loConsulta.AppendLine("			ELSE Clientes.Dir_Ent	")
            loConsulta.AppendLine("		END											AS Ship_Address,")
            loConsulta.AppendLine("     Case Renglones_facturas.tip_ori")
            loConsulta.AppendLine("			WHEN 'Pedido' THEN Renglones_facturas.Doc_ori")
            loConsulta.AppendLine("			ELSE ''	")
            loConsulta.AppendLine("		END											AS PO_NUMBER,")
            loConsulta.AppendLine("		Formas_Pagos.nom_for						AS Terms,")
            loConsulta.AppendLine("		Clientes.contacto   						AS Rep,")
            loConsulta.AppendLine("		Transportes.nom_tra							AS Via,")
            loConsulta.AppendLine("		Renglones_Facturas.Can_art1					AS Quantity,")
            loConsulta.AppendLine("		Renglones_Facturas.cod_art					AS Item_code,")
            loConsulta.AppendLine("		Articulos.nom_art							AS Description,")
            loConsulta.AppendLine("		Renglones_facturas.precio1					AS Price_Each,")
            loConsulta.AppendLine("		Renglones_facturas.mon_net					AS Amount,")
            loConsulta.AppendLine("		Facturas.comentario     					AS Note")
            loConsulta.AppendLine("FROM       Facturas ")
            loConsulta.AppendLine("    JOIN   Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            loConsulta.AppendLine("    JOIN   Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Formas_Pagos.Cod_For = Facturas.Cod_For")
            loConsulta.AppendLine("    JOIN   Transportes ON Transportes.Cod_Tra = Facturas.Cod_Tra")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loConsulta.AppendLine("    JOIN   Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven")
            loConsulta.AppendLine("WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_Ventas_FRO", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFacturas_Ventas_FRO.ReportSource = loObjetoReporte

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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' EAG: 09/10/15: Código Inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' EAG: 16/10/15: Se agregó la información del campo Rep y se agregó el campo Note.          '
'-------------------------------------------------------------------------------------------'

