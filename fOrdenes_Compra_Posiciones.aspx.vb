'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fOrdenes_Compra_Posiciones"
'-------------------------------------------------------------------------------------------'
Partial Class fOrdenes_Compra_Posiciones
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  Ordenes_Compras.Documento                               AS Documento,")
            loComandoSeleccionar.AppendLine("        ''                                 AS Factura, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Control                                 AS Control, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Status                                  AS Status, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Cod_Pro                                 AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("        (CASE WHEN Ordenes_Compras.Nom_Pro = ''")
            loComandoSeleccionar.AppendLine("            THEN Proveedores.Nom_Pro ")
            loComandoSeleccionar.AppendLine("            ELSE Ordenes_Compras.Nom_Pro")
            loComandoSeleccionar.AppendLine("        END)                                            AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("        (CASE WHEN Ordenes_Compras.Rif = ''")
            loComandoSeleccionar.AppendLine("            THEN Proveedores.Rif ")
            loComandoSeleccionar.AppendLine("            ELSE Ordenes_Compras.Rif")
            loComandoSeleccionar.AppendLine("        END)                                            AS Rif,")
            loComandoSeleccionar.AppendLine("        (CASE WHEN Ordenes_Compras.Dir_Fis = ''")
            loComandoSeleccionar.AppendLine("            THEN Proveedores.Dir_Fis ")
            loComandoSeleccionar.AppendLine("            ELSE Ordenes_Compras.Dir_Fis")
            loComandoSeleccionar.AppendLine("        END)                                            AS Dir_Fis,")
            loComandoSeleccionar.AppendLine("        (CASE WHEN Ordenes_Compras.Telefonos = ''")
            loComandoSeleccionar.AppendLine("            THEN Proveedores.Telefonos ")
            loComandoSeleccionar.AppendLine("            ELSE Ordenes_Compras.Telefonos")
            loComandoSeleccionar.AppendLine("        END)                                            AS Telefonos,")
            loComandoSeleccionar.AppendLine("        Proveedores.Fax                                 AS Fax, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Fec_Ini                                 AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Fec_Fin                                 AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Mon_Bru                                 AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Mon_Imp1                                AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Por_Des1                                AS Por_Des1, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Mon_Des1                                AS Mon_Des1, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Por_Rec1                                AS Por_Rec1, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Mon_Rec1                                AS Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Mon_Net                                 AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Mon_Sal                                 AS Mon_Sal, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Cod_For                                 AS Cod_For, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Cod_Mon                                 AS Cod_Mon, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Por_Imp1                                AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Comentario                              AS Comentario, ")
            loComandoSeleccionar.AppendLine("        Formas_Pagos.Nom_For                            AS Nom_For, ")
            loComandoSeleccionar.AppendLine("        Ordenes_Compras.Cod_Ven                                 AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("        Vendedores.Nom_Ven                              AS Nom_Ven,")
            loComandoSeleccionar.AppendLine("        Posiciones.Cod_Pos                              AS Cod_Pos,")
            loComandoSeleccionar.AppendLine("        Posiciones.Nom_Pos                              AS Nom_Pos,")
            loComandoSeleccionar.AppendLine("        Posiciones.Opcional                             AS Opcional,")
            loComandoSeleccionar.AppendLine("        Posiciones.Automatico                           AS Automatico,")
            loComandoSeleccionar.AppendLine("        CAST(COALESCE(Movimientos_Posiciones.Val_Log, 0) AS BIT) AS Autorizado,")
            loComandoSeleccionar.AppendLine("        COALESCE(Movimientos_Posiciones.Comentario,'')  AS Comentario_Posicion")
            loComandoSeleccionar.AppendLine("FROM    Ordenes_Compras")
            loComandoSeleccionar.AppendLine("    JOIN Proveedores ON Proveedores.Cod_Pro = Ordenes_Compras.Cod_Pro ")
            loComandoSeleccionar.AppendLine("    JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Ordenes_Compras.Cod_For ")
            loComandoSeleccionar.AppendLine("    JOIN Vendedores ON Vendedores.Cod_Ven = Ordenes_Compras.Cod_Ven ")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Posiciones ")
            loComandoSeleccionar.AppendLine("        ON Posiciones.Opcion = 'OrdenesCompra' ")
            loComandoSeleccionar.AppendLine("        AND Posiciones.Status = 'A'")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Movimientos_Posiciones ")
            loComandoSeleccionar.AppendLine("        ON  Movimientos_Posiciones.Origen = 'Ordenes_Compras'")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Posiciones.Adicional = ''")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Posiciones.Cod_Reg = Ordenes_Compras.Documento")
            loComandoSeleccionar.AppendLine("        AND Movimientos_Posiciones.Cod_Pos = Posiciones.Cod_Pos")
            loComandoSeleccionar.AppendLine("WHERE  " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("ORDER BY Posiciones.Orden, Posiciones.Cod_Pos")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Compra_Posiciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrdenes_Compra_Posiciones.ReportSource = loObjetoReporte

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
