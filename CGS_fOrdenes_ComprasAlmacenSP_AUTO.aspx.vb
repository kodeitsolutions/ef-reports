Imports System.Data
Partial Class CGS_fOrdenes_ComprasAlmacenSP_AUTO
    Inherits vis2formularios.frmReporteAutomatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try

        Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Ordenes_Compras.Cod_Pro                     AS Cod_Pro, ")
        loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro                         AS Nom_Pro, ")
        loComandoSeleccionar.AppendLine("           Proveedores.Rif                             AS Rif, ")
        loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis                         AS Dir_Fis, ")
        loComandoSeleccionar.AppendLine("           Proveedores.Telefonos                       AS Telefonos, ")
        loComandoSeleccionar.AppendLine("           Proveedores.Correo                          AS Correo, ")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Documento                   AS Documento, ")
        loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Uni                  AS Cod_Uni, ")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Ini                     AS Fec_Ini, ")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Rec                     AS Fec_Rec, ")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Comentario                  AS Comentario, ")
        loComandoSeleccionar.AppendLine("           Ordenes_Compras.Status                      AS Status, ")
        loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Art                  AS Cod_Art, ")
        loComandoSeleccionar.AppendLine("           Renglones_OCompras.Notas                    AS Notas, ")
        loComandoSeleccionar.AppendLine("           Articulos.Nom_Art                           AS Nom_Art, ")
        loComandoSeleccionar.AppendLine("           Articulos.Generico                          AS Generico, ")
        loComandoSeleccionar.AppendLine("           Renglones_OCompras.Can_Art1                 AS Can_Art1,   ")
            loComandoSeleccionar.AppendLine("       COALESCE((CONCAT(RTRIM(Requisiciones.Caracter1),CHAR(13),RTRIM(Requisiciones.Caracter2),CHAR(13),")
            loComandoSeleccionar.AppendLine("	    RTRIM(Requisiciones.Caracter3),CHAR(13),RTRIM(Requisiciones.Caracter4))),'') AS Solicitante ")
            loComandoSeleccionar.AppendLine(" FROM Ordenes_Compras ")
        loComandoSeleccionar.AppendLine("   JOIN Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Requisiciones ON Renglones_OCompras.Doc_Ori = Requisiciones.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_OCompras.Tip_Ori = 'Requisiciones'")
        loComandoSeleccionar.AppendLine("   JOIN Proveedores ON Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
        loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_OCompras.Cod_Art")
        loComandoSeleccionar.AppendLine("WHERE       " & cusAplicacion.goFormatos.pcCondicionPrincipal)
        loComandoSeleccionar.AppendLine("   AND Articulos.Cod_Dep NOT IN ('MP','PS','PC','AF','SR')")


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


        loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fOrdenes_ComprasAlmacenSP_AUTO", laDatosReporte)

        Me.mTraducirReporte(loObjetoReporte)

        Me.mFormatearCamposReporte(loObjetoReporte)

        Me.crvCGS_fOrdenes_ComprasAlmacenSP_AUTO.ReportSource = loObjetoReporte

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
