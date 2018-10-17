Imports System.Data
Partial Class MCL_fAleacion_Colada

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT ROW_NUMBER() OVER (ORDER BY Registros.Lote) AS Num,*")
            loComandoSeleccionar.AppendLine("INTO #tmpDatos")
            loComandoSeleccionar.AppendLine("FROM (")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("   SELECT  Renglones_Ajustes.Cod_Art				            AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("   		Articulos.Nom_Art						            AS Nom_Art,")
            loComandoSeleccionar.AppendLine("           COALESCE(Operaciones_Lotes.Cod_Lot,'')	            AS Lote,")
            loComandoSeleccionar.AppendLine("           COALESCE(Operaciones_Lotes.Cantidad,Renglones_Ajustes.Can_Art1,0)   AS Cant_Lote,")
            loComandoSeleccionar.AppendLine("           COALESCE(R_Produccion_Consumido.Can_Art1, 0)        AS Cant_Consumido,")
            loComandoSeleccionar.AppendLine("           COALESCE(R_Produccion_Consumido.Cod_Art,'')         AS Art_Consumido,")
            loComandoSeleccionar.AppendLine("           COALESCE(Consumido.Nom_Art,'')                      AS NomArt_Consumido")
            loComandoSeleccionar.AppendLine("   FROM Ajustes ")
            loComandoSeleccionar.AppendLine("       JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            loComandoSeleccionar.AppendLine("   	JOIN Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art")
            loComandoSeleccionar.AppendLine("       JOIN Secciones ON Articulos.Cod_Sec = Secciones.Cod_Sec")
            loComandoSeleccionar.AppendLine("           AND Secciones.Cod_Dep = Articulos.Cod_Dep")
            loComandoSeleccionar.AppendLine("       LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones_Ajustes.Documento")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("           AND Renglones_Ajustes.Cod_Alm = Operaciones_Lotes.Cod_Alm")
            loComandoSeleccionar.AppendLine("           AND Operaciones_Lotes.Tip_Doc = 'Ajustes_Inventarios' ")
            loComandoSeleccionar.AppendLine("   		AND Operaciones_Lotes.Ren_Ori = Renglones_Ajustes.Renglon")
            loComandoSeleccionar.AppendLine("   		AND Operaciones_Lotes.Tip_Ope = 'Salida'")
            loComandoSeleccionar.AppendLine("   	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("   		AND Mediciones.Origen = 'Ajustes_Inventarios'")
            loComandoSeleccionar.AppendLine("   		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("   		AND Renglones_Ajustes.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("   	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("   		AND Piezas.Cod_Var = 'AINV-NPIEZ'")
            loComandoSeleccionar.AppendLine("   		AND Piezas.Res_Num > 0")
            loComandoSeleccionar.AppendLine("       LEFT JOIN Operaciones_Lotes AS Entrada ON Entrada.Cod_Lot = Operaciones_Lotes.Cod_Lot")
            loComandoSeleccionar.AppendLine("           AND Entrada.Tip_Doc = 'Ajustes_Inventarios' ")
            loComandoSeleccionar.AppendLine("   		AND Entrada.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("           AND Entrada.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("           AND Entrada.Cod_Alm = Operaciones_Lotes.Cod_Alm")
            loComandoSeleccionar.AppendLine("       LEFT JOIN Ajustes AS Entrada_Produccion On Entrada.Num_Doc = Entrada_Produccion.Documento")
            loComandoSeleccionar.AppendLine("       LEFT JOIN Ajustes AS Produccion_Consumido ON Produccion_Consumido.Documento = Entrada_Produccion.Caracter1")
            loComandoSeleccionar.AppendLine("       LEFT JOIN Renglones_Ajustes AS R_Produccion_Consumido ON Produccion_Consumido.Documento = R_Produccion_Consumido.Documento")
            loComandoSeleccionar.AppendLine("       LEFT JOIN Articulos AS Consumido ON R_Produccion_Consumido.Cod_Art = Consumido.Cod_Art")
            loComandoSeleccionar.AppendLine("   WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("   	AND Renglones_Ajustes.Cod_Tip = 'S10' ")
            loComandoSeleccionar.AppendLine("   ) Registros")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT ROW_NUMBER() OVER (ORDER BY Articulos.Cod_Art) AS Renglon,*")
            loComandoSeleccionar.AppendLine("INTO #tmpArticulos")
            loComandoSeleccionar.AppendLine("FROM (")
            loComandoSeleccionar.AppendLine("       SELECT DISTINCT")
            loComandoSeleccionar.AppendLine("          #tmpDatos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("          #tmpDatos.Nom_Art")
            loComandoSeleccionar.AppendLine("       FROM #tmpDatos")
            loComandoSeleccionar.AppendLine("       ) Articulos")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnFilas AS INT = (SELECT COUNT(*) FROM #tmpArticulos)")
            loComandoSeleccionar.AppendLine("DECLARE @lnRenglon AS INT = 1")
            loComandoSeleccionar.AppendLine("DECLARE @lcCod_Art AS VARCHAR(8) = ''")
            loComandoSeleccionar.AppendLine("DECLARE @lcNom_Art AS VARCHAR(MAX) = ''")
            loComandoSeleccionar.AppendLine("DECLARE @lcNomArticulos AS VARCHAR(MAX) = ''")
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArticulos AS VARCHAR(MAX) = ''")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("WHILE (@lnRenglon <= @lnFilas)")
            loComandoSeleccionar.AppendLine("BEGIN")
            loComandoSeleccionar.AppendLine("   SELECT  @lcCod_Art = Cod_Art,")
            loComandoSeleccionar.AppendLine("           @lcNom_Art = Nom_Art")
            loComandoSeleccionar.AppendLine("   FROM #tmpArticulos")
            loComandoSeleccionar.AppendLine("   WHERE Renglon =  @lnRenglon")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("   SET @lcNomArticulos = @lcNomArticulos + @lcNom_Art + CHAR(13)")
            loComandoSeleccionar.AppendLine("   SET @lcCodArticulos = @lcCodArticulos + @lcCod_Art + CHAR(13)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("   SET @lnRenglon = @lnRenglon + 1")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("END")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT #tmpDatos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("   	#tmpDatos.Nom_Art,")
            loComandoSeleccionar.AppendLine("       #tmpDatos.Lote,")
            loComandoSeleccionar.AppendLine("       #tmpDatos.Cant_Lote,")
            loComandoSeleccionar.AppendLine("       #tmpDatos.Cant_Consumido,")
            loComandoSeleccionar.AppendLine("       #tmpDatos.Art_Consumido,")
            loComandoSeleccionar.AppendLine("       #tmpDatos.NomArt_Consumido,")
            loComandoSeleccionar.AppendLine("       COALESCE((SELECT Lote FROM #tmpDatos WHERE Num = 1),'') AS First,")
            loComandoSeleccionar.AppendLine("       COALESCE((SELECT Lote FROM #tmpDatos WHERE Num = (SELECT MAX(Num) FROM #tmpDatos)),'')  AS Last,")
            loComandoSeleccionar.AppendLine("       @lcNomArticulos AS Nom_Articulos,")
            loComandoSeleccionar.AppendLine("       @lcCodArticulos AS Cod_Articulos")
            loComandoSeleccionar.AppendLine("FROM #tmpDatos")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpDatos")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpArticulos")



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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("MCL_fAleacion_Colada", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMCL_fAleacion_Colada.ReportSource = loObjetoReporte

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