﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPresupuestos_Requisitos"
'-------------------------------------------------------------------------------------------'
Partial Class fPresupuestos_Requisitos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  Presupuestos.Documento                               AS Documento,")
            loComandoSeleccionar.AppendLine("        ''                                                   AS Factura, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Control                                 AS Control, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Status                                  AS Status, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Cod_Pro                                 AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("        (CASE WHEN Presupuestos.Nom_Pro = ''")
            loComandoSeleccionar.AppendLine("            THEN Proveedores.Nom_Pro ")
            loComandoSeleccionar.AppendLine("            ELSE Presupuestos.Nom_Pro")
            loComandoSeleccionar.AppendLine("        END)                                            AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("        (CASE WHEN Presupuestos.Rif = ''")
            loComandoSeleccionar.AppendLine("            THEN Proveedores.Rif ")
            loComandoSeleccionar.AppendLine("            ELSE Presupuestos.Rif")
            loComandoSeleccionar.AppendLine("        END)                                            AS Rif,")
            loComandoSeleccionar.AppendLine("        (CASE WHEN Presupuestos.Dir_Fis = ''")
            loComandoSeleccionar.AppendLine("            THEN Proveedores.Dir_Fis ")
            loComandoSeleccionar.AppendLine("            ELSE Presupuestos.Dir_Fis")
            loComandoSeleccionar.AppendLine("        END)                                            AS Dir_Fis,")
            loComandoSeleccionar.AppendLine("        (CASE WHEN Presupuestos.Telefonos = ''")
            loComandoSeleccionar.AppendLine("            THEN Proveedores.Telefonos ")
            loComandoSeleccionar.AppendLine("            ELSE Presupuestos.Telefonos")
            loComandoSeleccionar.AppendLine("        END)                                            AS Telefonos,")
            loComandoSeleccionar.AppendLine("        Proveedores.Fax                                 AS Fax, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Fec_Ini                                 AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Fec_Fin                                 AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Mon_Bru                                 AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Mon_Imp1                                AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Por_Des1                                AS Por_Des1, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Mon_Des1                                AS Mon_Des1, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Por_Rec1                                AS Por_Rec1, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Mon_Rec1                                AS Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Mon_Net                                 AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Mon_Sal                                 AS Mon_Sal, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Cod_For                                 AS Cod_For, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Cod_Mon                                 AS Cod_Mon, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Por_Imp1                                AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Comentario                              AS Comentario, ")
            loComandoSeleccionar.AppendLine("        Formas_Pagos.Nom_For                            AS Nom_For, ")
            loComandoSeleccionar.AppendLine("        Presupuestos.Cod_Ven                                 AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("        Vendedores.Nom_Ven                              AS Nom_Ven,")
            loComandoSeleccionar.AppendLine("        Requerimientos.Cod_Req                          AS Cod_Req,")
            loComandoSeleccionar.AppendLine("        Requerimientos.Nom_Req                          AS Nom_Req,")
            loComandoSeleccionar.AppendLine("        Requerimientos.Opcional                         AS Opcional,")
            loComandoSeleccionar.AppendLine("        Requerimientos.Peso                             AS Peso,")
            loComandoSeleccionar.AppendLine("        Requerimientos.Tip_Pro                          AS Tip_Pro,")
            loComandoSeleccionar.AppendLine("        COALESCE(Movimientos_Requerimientos.Val_Num, 0)                 AS Val_Num,")
            loComandoSeleccionar.AppendLine("        CAST(COALESCE(Movimientos_Requerimientos.Val_Log, 0) AS BIT)    AS Val_Log,")
            loComandoSeleccionar.AppendLine("        COALESCE(Movimientos_Requerimientos.Val_Car, '')                AS Val_Car,")
            loComandoSeleccionar.AppendLine("        COALESCE(Movimientos_Requerimientos.Val_Fec, ")
            loComandoSeleccionar.AppendLine("            CAST('19000101' AS DATE))                                   AS Val_Fec,")
            loComandoSeleccionar.AppendLine("        COALESCE(Movimientos_Requerimientos.Val_Mem, '')                AS Val_Mem,")
            loComandoSeleccionar.AppendLine("        COALESCE(Movimientos_Requerimientos.Comentario,'')              AS Comentario_Requerimientos")
            loComandoSeleccionar.AppendLine("FROM    Presupuestos")
            loComandoSeleccionar.AppendLine("    JOIN Proveedores ON Proveedores.Cod_Pro = Presupuestos.Cod_Pro ")
            loComandoSeleccionar.AppendLine("    JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Presupuestos.Cod_For ")
            loComandoSeleccionar.AppendLine("    JOIN Vendedores ON Vendedores.Cod_Ven = Presupuestos.Cod_Ven ")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Requerimientos ")
            loComandoSeleccionar.AppendLine("        ON Requerimientos.Opcion = 'Presupuestos' ")
            loComandoSeleccionar.AppendLine("        AND Requerimientos.Status = 'A'")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Movimientos_Requerimientos ")
            loComandoSeleccionar.AppendLine("        ON  Movimientos_Requerimientos.Origen = 'Presupuestos'")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Requerimientos.Adicional = ''")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Requerimientos.Cod_Reg = Presupuestos.Documento")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Requerimientos.Cod_Req = Requerimientos.Cod_Req")
            loComandoSeleccionar.AppendLine("WHERE  " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("ORDER BY Requerimientos.Orden, Requerimientos.Cod_Req")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPresupuestos_Requisitos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPresupuestos_Requisitos.ReportSource = loObjetoReporte

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
' RJG: 15/01/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
