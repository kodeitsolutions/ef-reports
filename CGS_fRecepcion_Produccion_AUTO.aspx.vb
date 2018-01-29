Imports System.Data
Partial Class CGS_fRecepcion_Produccion_AUTO

    Inherits vis2formularios.frmReporteAutomatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try
        Dim Empresa As String = cusAplicacion.goEmpresa.pcNombre

        Dim loComandoSeleccionar As New StringBuilder()

        loComandoSeleccionar.AppendLine("SELECT Ajustes.Documento						AS Documento, ")
        loComandoSeleccionar.AppendLine("		Ajustes.Status							AS Estatus, ")
        loComandoSeleccionar.AppendLine("       Ajustes.Fec_Ini							AS Fec_Ini, ")
        loComandoSeleccionar.AppendLine("       Ajustes.Comentario						AS Comentario, ")
        loComandoSeleccionar.AppendLine("       Renglones_Ajustes.Renglon				AS Renglon, ")
        loComandoSeleccionar.AppendLine("       Renglones_Ajustes.Cod_Art				AS Cod_Art, ")
        loComandoSeleccionar.AppendLine("		Articulos.Nom_Art						AS Nom_Art,")
        loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Cod_Alm				AS Cod_Alm,")
        loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm						AS Nom_Alm,")
        loComandoSeleccionar.AppendLine("       Renglones_Ajustes.Cod_Uni				AS Cod_Uni, ")
        loComandoSeleccionar.AppendLine("       Renglones_Ajustes.Can_Art1				AS Can_Art1, ")
        loComandoSeleccionar.AppendLine("       Renglones_Ajustes.Notas					AS Notas,")
        loComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cod_Lot,'')	AS Lote,")
        loComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cantidad,0)	AS Cant_Lote,")
        loComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)				AS Piezas,")
        loComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num,0)		    AS Porc_Desperdicio,")
        loComandoSeleccionar.AppendLine("		COALESCE((Desperdicio.Res_Num * Operaciones_Lotes.Cantidad)/100, ")
        loComandoSeleccionar.AppendLine("		(Desperdicio.Res_Num * Renglones_Ajustes.Can_Art1)/100,0) AS Cant_Desperdicio")
        loComandoSeleccionar.AppendLine("FROM Ajustes ")
        loComandoSeleccionar.AppendLine("   JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
        loComandoSeleccionar.AppendLine("       AND Renglones_Ajustes.Cod_Tip = 'E02'")
        loComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art")
        loComandoSeleccionar.AppendLine("   JOIN Almacenes ON Renglones_Ajustes.Cod_Alm = Almacenes.Cod_Alm")
        loComandoSeleccionar.AppendLine("   LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones_Ajustes.Documento")
        loComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Tip_Doc = 'Ajustes_Inventarios' ")
        loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Ajustes.Renglon")
        loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
        loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Ajustes.Documento")
        loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Ajustes_Inventarios'")
        loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
        loComandoSeleccionar.AppendLine("		AND Renglones_Ajustes.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
        loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
        loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'AINV-NPIEZ'")
        loComandoSeleccionar.AppendLine("		AND Piezas.Res_Num > 0")
        loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Mediciones.Documento = Desperdicio.Documento")
        loComandoSeleccionar.AppendLine("		AND Desperdicio.Cod_Var = 'AINV-PDESP'")
        loComandoSeleccionar.AppendLine("		AND Desperdicio.Res_Num > 0")
        loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

        'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

        Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")



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


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fRecepcion_Produccion_AUTO", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fRecepcion_Produccion_AUTO.ReportSource = loObjetoReporte

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

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 27/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos genericos
'-------------------------------------------------------------------------------------------'
' CMS: 16/09/09: Se Agrego la distribucion de impuesto
'-------------------------------------------------------------------------------------------'
' MAT: 01/03/11: Ajuste de la vista de diseño												'
'-------------------------------------------------------------------------------------------'