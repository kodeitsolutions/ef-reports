Imports System.Data
Partial Class PAS_fConsumos_Produccion

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim Empresa As String = cusAplicacion.goEmpresa.pcNombre

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Encabezados.Documento					AS Documento,")
            loComandoSeleccionar.AppendLine("		Encabezados.Proyecto					AS OProduccion,")
            loComandoSeleccionar.AppendLine("		Encabezados.Status						AS Estatus, ")
            loComandoSeleccionar.AppendLine("       Encabezados.Fec_Ini						AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("       Encabezados.Comentario					AS Comentario,")
            loComandoSeleccionar.AppendLine("		Encabezados.Cod_Alm						AS Cod_Alm,")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm						AS Nom_Alm,")
            loComandoSeleccionar.AppendLine("		Renglones.Renglon						AS Renglon, ")
            loComandoSeleccionar.AppendLine("       Renglones.Cod_Art						AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("		Renglones.Nom_Art						AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		Renglones.Can_Art						AS Can_Art,")
            loComandoSeleccionar.AppendLine("		Renglones.Cod_Uni						AS Cod_Uni,")
            loComandoSeleccionar.AppendLine("		COALESCE(Operaciones_Lotes.Cod_Lot,'')	AS Lote,")
            loComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cantidad,0)	AS Cant_Lote,")
            loComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)				AS Piezas")
            loComandoSeleccionar.AppendLine("FROM Encabezados")
            loComandoSeleccionar.AppendLine("	JOIN Renglones ON Encabezados.Documento = Renglones.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Encabezados.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN Almacenes ON Encabezados.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones.Documento")
            loComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Tip_Doc = 'Encabezados'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Adicional = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones.Cod_Art ")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Salida'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Encabezados.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Encabezados'")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Cod_Art = Renglones.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Cod_Alm = Encabezados.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND Renglones.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'CPRO-NPIEZ'")
            loComandoSeleccionar.AppendLine("		AND Piezas.Res_Num > 0")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("	AND Encabezados.Origen = 'Consumos Produccion' ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")




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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("PAS_fConsumos_Produccion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvPAS_fConsumos_Produccion.ReportSource = loObjetoReporte

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
' JJD: 27/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos genericos
'-------------------------------------------------------------------------------------------'
' CMS: 16/09/09: Se Agrego la distribucion de impuesto
'-------------------------------------------------------------------------------------------'
' MAT: 01/03/11: Ajuste de la vista de diseño												'
'-------------------------------------------------------------------------------------------'