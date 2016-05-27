Imports System.Data
Partial Class CGS_fTraslados_Almacenes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar1 As New StringBuilder()

            loComandoSeleccionar1.AppendLine("SELECT  Usu_Cre ")
            loComandoSeleccionar1.AppendLine("FROM      Traslados")
            loComandoSeleccionar1.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios1 As New cusDatos.goDatos

            Dim laDatosReporte1 As DataSet = loServicios1.mObtenerTodosSinEsquema(loComandoSeleccionar1.ToString, "usuario")

            Dim aString As String = laDatosReporte1.Tables(0).Rows(0).Item(0)
            aString = Trim(aString)

            Dim loComandoSeleccionar2 As New StringBuilder()

            loComandoSeleccionar2.AppendLine(" SELECT    Nom_Usu ")
            loComandoSeleccionar2.AppendLine(" FROM      Usuarios ")
            loComandoSeleccionar2.AppendLine(" WHERE     Cod_Usu = '" & aString & "'")
            loComandoSeleccionar2.AppendLine("		    AND Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))

            Dim loServicios2 As New cusDatos.goDatos

            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte2 As DataSet = loServicios2.mObtenerTodosSinEsquema(loComandoSeleccionar2.ToString, "nombreUsuario")
            Dim aString2 As String = laDatosReporte2.Tables("nombreUsuario").Rows(0).Item(0)
            aString2 = RTrim(aString2)

            Dim loComandoSeleccionar3 As New StringBuilder()

            loComandoSeleccionar3.AppendLine("SELECT  Traslados.Documento,")
            loComandoSeleccionar3.AppendLine("        Traslados.Status, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Fec_Ini, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Alm_Ori, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Alm_Des, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Cod_Tra,")
            loComandoSeleccionar3.AppendLine("	      Transportes.Nom_Tra, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Comentario, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Renglon, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Cod_Art, ")
            loComandoSeleccionar3.AppendLine("	      Articulos.Nom_Art,")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Cod_Uni, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Can_Art1, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Notas,")
            loComandoSeleccionar3.AppendLine("        Almacenes.Nom_Alm As Nom_Ori, ")
            loComandoSeleccionar3.AppendLine("	      (SELECT Nom_Alm ")
            loComandoSeleccionar3.AppendLine("	      FROM Almacenes ")
            loComandoSeleccionar3.AppendLine("	      WHERE Cod_Alm = Traslados.Alm_Des) AS Nom_Des,")
            loComandoSeleccionar3.AppendLine("        Operaciones_Lotes.Cod_Lot         AS Lote,")
            loComandoSeleccionar3.AppendLine("        Operaciones_Lotes.Cantidad         AS Cant_Lote,")
            loComandoSeleccionar3.AppendLine("        '" & aString2 & "' AS Usuario,")
            loComandoSeleccionar3.AppendLine("        Mediciones.Control,")
            loComandoSeleccionar3.AppendLine("        Variables.Nom_Var,")
            loComandoSeleccionar3.AppendLine("        Renglones_Mediciones.Cod_Uni    AS Uni_Med,")
            loComandoSeleccionar3.AppendLine("        Renglones_Mediciones.Res_Num")
            loComandoSeleccionar3.AppendLine("INTO #tmpLotes")
            loComandoSeleccionar3.AppendLine("FROM 	Traslados ")
            loComandoSeleccionar3.AppendLine("      JOIN Renglones_Traslados ON Traslados.Documento = Renglones_Traslados.Documento")
            loComandoSeleccionar3.AppendLine("	    JOIN Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar3.AppendLine("      JOIN Transportes ON Traslados.Cod_Tra = Transportes.Cod_Tra")
            loComandoSeleccionar3.AppendLine("      JOIN Almacenes ON Traslados.Alm_Ori   =   Almacenes.Cod_Alm")
            loComandoSeleccionar3.AppendLine("      JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones_Traslados.Documento")
            loComandoSeleccionar3.AppendLine("          AND Renglones_Traslados.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar3.AppendLine("          AND SUBSTRING(Operaciones_Lotes.Notas,2, LEN(Operaciones_Lotes.Notas)) = CAST(Renglones_Traslados.Renglon AS NVARCHAR(255))")
            loComandoSeleccionar3.AppendLine("          AND Operaciones_Lotes.Tip_Doc = 'Traslados' AND Operaciones_Lotes.Tip_Ope = 'Salida'")
            loComandoSeleccionar3.AppendLine("      JOIN Mediciones ON Traslados.Documento = Mediciones.Cod_Reg")
            loComandoSeleccionar3.AppendLine("          AND Mediciones.Origen = 'Traslados'")
            loComandoSeleccionar3.AppendLine("          AND Mediciones.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar3.AppendLine("      JOIN Renglones_Mediciones ON Mediciones.Documento = Renglones_Mediciones.Documento")
            loComandoSeleccionar3.AppendLine("      JOIN Variables ON Renglones_Mediciones.Cod_Var = Variables.Cod_Var")
            loComandoSeleccionar3.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar3.AppendLine("")
            loComandoSeleccionar3.AppendLine("")
            loComandoSeleccionar3.AppendLine("SELECT  Traslados.Documento,")
            loComandoSeleccionar3.AppendLine("        Traslados.Status, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Fec_Ini, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Alm_Ori, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Alm_Des, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Cod_Tra,")
            loComandoSeleccionar3.AppendLine("	      Transportes.Nom_Tra, ")
            loComandoSeleccionar3.AppendLine("        Traslados.Comentario, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Renglon, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Cod_Art, ")
            loComandoSeleccionar3.AppendLine("	      Articulos.Nom_Art,")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Cod_Uni, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Can_Art1, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Notas,")
            loComandoSeleccionar3.AppendLine("        Almacenes.Nom_Alm As Nom_Ori, ")
            loComandoSeleccionar3.AppendLine("	      (SELECT Nom_Alm ")
            loComandoSeleccionar3.AppendLine("	      FROM Almacenes ")
            loComandoSeleccionar3.AppendLine("	      WHERE Cod_Alm = Traslados.Alm_Des) AS Nom_Des,")
            loComandoSeleccionar3.AppendLine("        ''                                 AS Lote,")
            loComandoSeleccionar3.AppendLine("        0                                  AS Cant_Lote,")
            loComandoSeleccionar3.AppendLine("        '" & aString2 & "'                 AS Usuario,")
            loComandoSeleccionar3.AppendLine("        ''                                 AS Control,")
            loComandoSeleccionar3.AppendLine("        ''                                 AS Nom_Var,")
            loComandoSeleccionar3.AppendLine("        ''                                 AS Uni_Med,")
            loComandoSeleccionar3.AppendLine("        0                                  AS Res_Num")
            loComandoSeleccionar3.AppendLine("INTO #tmpNoLotes")
            loComandoSeleccionar3.AppendLine("FROM 	Traslados ")
            loComandoSeleccionar3.AppendLine("      JOIN Renglones_Traslados ON Traslados.Documento = Renglones_Traslados.Documento")
            loComandoSeleccionar3.AppendLine("	    JOIN Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar3.AppendLine("      JOIN Transportes ON Traslados.Cod_Tra = Transportes.Cod_Tra")
            loComandoSeleccionar3.AppendLine("      JOIN Almacenes ON Traslados.Alm_Ori   =   Almacenes.Cod_Alm")
            loComandoSeleccionar3.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar3.AppendLine("      AND Traslados.Documento NOT IN (SELECT Documento FROM #tmpLotes)")
            loComandoSeleccionar3.AppendLine("")
            loComandoSeleccionar3.AppendLine("SELECT * FROM #tmpLotes")
            loComandoSeleccionar3.AppendLine("UNION")
            loComandoSeleccionar3.AppendLine("SELECT * FROM #tmpNoLotes")
            loComandoSeleccionar3.AppendLine("")
            loComandoSeleccionar3.AppendLine("DROP TABLE #tmpLotes")
            loComandoSeleccionar3.AppendLine("DROP TABLE #tmpNoLotes")


            'Me.mEscribirConsulta(loComandoSeleccionar3.ToString)

            Dim loServicios3 As New cusDatos.goDatos
            Dim laDatosReporte3 As DataSet = loServicios3.mObtenerTodosSinEsquema(loComandoSeleccionar3.ToString, "curReportes")



            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte3.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                      "No se Encontraron Registros para los Parámetros Especificados. ", _
      vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                       "350px", _
                       "200px")
            End If


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte3.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fTraslados_Almacenes", laDatosReporte3)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fTraslados_Almacenes.ReportSource = loObjetoReporte

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