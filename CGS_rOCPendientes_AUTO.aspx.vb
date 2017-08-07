'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rOCPendientes_AUTO"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rOCPendientes_AUTO

    Inherits vis2formularios.frmReporteAutomatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try

        Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT Ordenes_Compras.Documento 			AS Documento,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Fec_Ini				AS Fecha,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Comentario			AS Comentario,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Cod_Suc	    		AS Cod_Suc,")
            loComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Art			AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Alm			AS Cod_Alm,")
            loComandoSeleccionar.AppendLine("		Renglones_OCompras.Cod_Uni			AS Cod_Uni,")
            loComandoSeleccionar.AppendLine("		Renglones_OCompras.Can_Art1			AS Cantidad,")
            loComandoSeleccionar.AppendLine("		Renglones_OCompras.Precio1			AS Precio,")
            loComandoSeleccionar.AppendLine("		Renglones_OCompras.Mon_Net			AS Neto,")
            loComandoSeleccionar.AppendLine("		Renglones_OCompras.Mon_Imp1			AS Impuesto,")
            loComandoSeleccionar.AppendLine("		Renglones_OCompras.Notas			AS Notas,")
            loComandoSeleccionar.AppendLine("		Renglones_OCompras.Comentario		AS Com_Renglon,")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art					AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro					AS Nom_Pro")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Compras")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_OCompras.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON 	Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE Ordenes_Compras.Status = 'Pendiente'")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini >= '01/01/2017'")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------'
            ' Carga la imagen del logo en cusReportes                                                   '
            '-------------------------------------------------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------'
            ' Verificando si el select (tabla nº0) trae registros                                       '
            '-------------------------------------------------------------------------------------------'

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                          vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rOCPendientes_AUTO", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rOCPendientes_AUTO.ReportSource = loObjetoReporte

        'Catch loExcepcion As Exception

        '    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
        '              "No se pudo Completar el Proceso: " & loExcepcion.Message, _
        '               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
        '               "auto", _
        '               "auto")

        'End Try

    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                                 "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                                  vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                                  "auto", _
                                  "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 27/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/09: Ajuste al formato del impuesto IVA
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 11/06/10: Proveedor Genarico
'-------------------------------------------------------------------------------------------'
' JJD: 11/03/5: Ajustes al formato para el cliente CEGASA
'-------------------------------------------------------------------------------------------'