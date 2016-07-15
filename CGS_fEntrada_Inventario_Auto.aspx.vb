Imports System.Data
Partial Class CGS_fEntrada_Inventario_Auto
    Inherits vis2formularios.frmReporteAutomatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try

        Dim loComandoSeleccionar As New StringBuilder()

        loComandoSeleccionar.AppendLine("SELECT     Ordenes_Compras.Documento,")
        loComandoSeleccionar.AppendLine("		    Ordenes_Compras.Cod_Pro,")
        loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro,")
        loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
        loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
        loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
        loComandoSeleccionar.AppendLine("           Proveedores.Telefonos,")
        loComandoSeleccionar.AppendLine("           Proveedores.Correo, ")
        loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Uni,")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Ini, ")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Rec, ")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Comentario, ")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Status, ")
        loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Ven,")
        loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Art, ")
        loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
        loComandoSeleccionar.AppendLine("		    Articulos.Generico,")
        loComandoSeleccionar.AppendLine("		    Renglones_OCompras.Notas,")
        loComandoSeleccionar.AppendLine("           Renglones_OCompras.Renglon, ")
        loComandoSeleccionar.AppendLine("           Renglones_OCompras.Can_Art1, ")
        loComandoSeleccionar.AppendLine("           Renglones_OCompras.Comentario   As Comentario_Renglon ")
        loComandoSeleccionar.AppendLine(" FROM      Ordenes_Compras ")
        loComandoSeleccionar.AppendLine("       JOIN Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento ")
        loComandoSeleccionar.AppendLine("       JOIN Proveedores ON Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro ")
        loComandoSeleccionar.AppendLine("		JOIN Formas_Pagos ON Ordenes_Compras.Cod_For = Formas_Pagos.Cod_For")
        loComandoSeleccionar.AppendLine("       JOIN Articulos ON Renglones_OCompras.Cod_Art = Articulos.Cod_Art")
        loComandoSeleccionar.AppendLine(" WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)

        'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

        Dim loServicios As New cusDatos.goDatos

        Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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


        loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fEntrada_Inventario_Auto", laDatosReporte)

        Me.mTraducirReporte(loObjetoReporte)

        Me.mFormatearCamposReporte(loObjetoReporte)

        Me.crvCGS_fEntrada_Inventario_Auto.ReportSource = loObjetoReporte

        'Catch loExcepcion As Exception

        '    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
        '                  "No se pudo Completar el Proceso: " & loExcepcion.Message, _
        '                   vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
        '                   "auto", _
        '                   "auto")

        'End Try

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
' JJD: 08/11/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos gen.
'-------------------------------------------------------------------------------------------'
' JJD: 09/01/10: Se cambio para que leyera datos de genericos de la Cotizacion cuando aplique
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT: 02/09/11: Adición de Comentario en Renglones
'-------------------------------------------------------------------------------------------'
