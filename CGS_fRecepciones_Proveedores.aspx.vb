Imports System.Data
Partial Class CGS_fRecepciones_Proveedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT Recepciones.Cod_Pro                     AS Cod_Pro, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Nom_Pro                     AS Nom_Pro, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Rif                         AS Rif, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Dir_Fis                     AS Dir_Fis, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Telefonos                   AS Telefonos, ")
            loComandoSeleccionar.AppendLine("	    Recepciones.Documento					AS Documento,")
            loComandoSeleccionar.AppendLine("	    Recepciones.Status					    AS Status,")
            loComandoSeleccionar.AppendLine("	    Recepciones.Fec_Ini						AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("	    Recepciones.Comentario					AS Comentario,")
            loComandoSeleccionar.AppendLine("	    Recepciones.Notas					    AS Datos,")
            'loComandoSeleccionar.AppendLine("	    Renglones_Recepciones.Renglon			AS Renglon,")
            loComandoSeleccionar.AppendLine("	    Renglones_Recepciones.Cod_Art			AS Cod_Art,")
            loComandoSeleccionar.AppendLine("	    Articulos.Nom_Art						AS Nom_Art,")
            loComandoSeleccionar.AppendLine("       Articulos.Generico						AS Generico,")
            loComandoSeleccionar.AppendLine("       Renglones_Recepciones.Notas				AS Notas, ")
            loComandoSeleccionar.AppendLine("		COALESCE((CASE WHEN Renglones_Recepciones.Comentario <> ''")
            loComandoSeleccionar.AppendLine("                      THEN CONVERT(NUMERIC(18,2), REPLACE(REPLACE(Renglones_Recepciones.Comentario,'.',''), ',','.' ) )")
            loComandoSeleccionar.AppendLine("                      ELSE 0")
            loComandoSeleccionar.AppendLine("                 END) ,0)                      AS Peso,")
            loComandoSeleccionar.AppendLine("	    Renglones_Recepciones.Can_Art1          AS Cantidad, ")
            loComandoSeleccionar.AppendLine("	    Renglones_Recepciones.Cod_Uni			AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("		Renglones_Recepciones.Caracter1			AS Adicional, ")
            loComandoSeleccionar.AppendLine("	    Renglones_Recepciones.Doc_Ori			As Doc_Ori, ")
            loComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cod_Lot,'')	AS Lote,")
            loComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cantidad,0)	AS Cantidad_Lote,")
            loComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)			    AS Piezas,")
            loComandoSeleccionar.AppendLine("		COALESCE(Desperdicio.Res_Num, 0)		AS Porc_Desperdicio,")
            loComandoSeleccionar.AppendLine("		COALESCE((Desperdicio.Res_Num*Operaciones_Lotes.Cantidad)/100, ")
            loComandoSeleccionar.AppendLine("		(Desperdicio.Res_Num*Renglones_Recepciones.Can_Art1)/100, 0)	AS Cant_Desperdicio")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INTO #tmpRecepcion")
            loComandoSeleccionar.AppendLine("FROM Recepciones")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Recepciones ON Recepciones.Documento = Renglones_Recepciones.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Recepciones.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones_Recepciones.Documento")
            loComandoSeleccionar.AppendLine("       AND Renglones_Recepciones.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Ren_Ori = Renglones_Recepciones.Renglon")
            loComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Tip_Doc = 'Recepciones' AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Origen = 'Recepciones'")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("       AND Renglones_Recepciones.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'NREC-NPIEZ'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Mediciones.Documento = Desperdicio.Documento")
            loComandoSeleccionar.AppendLine("		AND Desperdicio.Cod_Var = 'NREC-PDESP'")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal & ";")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("WITH Sumatorias (Cod_Art, Lote, Cantidad, Cantidad_Lote, Peso) AS")
            loComandoSeleccionar.AppendLine("(")
            loComandoSeleccionar.AppendLine("   SELECT Cod_Art, Lote,SUM(Cantidad) AS Cantidad, SUM(Cantidad_Lote) AS Cantidad_Lote, SUM(Peso) AS Peso")
            loComandoSeleccionar.AppendLine("   FROM #tmpRecepcion")
            loComandoSeleccionar.AppendLine("   GROUP BY Cod_Art, Lote")
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("UPDATE #tmpRecepcion ")
            loComandoSeleccionar.AppendLine("SET #tmpRecepcion.Cantidad = Sumatorias.Cantidad, ")
            loComandoSeleccionar.AppendLine("	#tmpRecepcion.Cantidad_Lote = Sumatorias.Cantidad_Lote,")
            loComandoSeleccionar.AppendLine("	#tmpRecepcion.Peso = Sumatorias.Peso")
            loComandoSeleccionar.AppendLine("FROM #tmpRecepcion")
            loComandoSeleccionar.AppendLine("	INNER JOIN Sumatorias ON Sumatorias.Cod_Art = #tmpRecepcion.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE #tmpRecepcion.Lote = Sumatorias.Lote")
            loComandoSeleccionar.AppendLine("   AND #tmpRecepcion.Cod_Art = Sumatorias.Cod_Art")
            loComandoSeleccionar.AppendLine("UPDATE #tmpRecepcion SET Cant_Desperdicio = (Cantidad_Lote * Porc_Desperdicio)/100 ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT * FROM #tmpRecepcion WHERE Adicional <> 'DECLARADO' ORDER BY Cod_Art, Lote")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpRecepcion")
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fRecepciones_Proveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fRecepciones_Proveedores.ReportSource = loObjetoReporte

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