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
            loComandoSeleccionar3.AppendLine("        Traslados.Comentario, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Renglon, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Cod_Art, ")
            loComandoSeleccionar3.AppendLine("        Articulos.Nom_Art,")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Cod_Uni, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Can_Art1, ")
            loComandoSeleccionar3.AppendLine("        Renglones_Traslados.Notas,")
            loComandoSeleccionar3.AppendLine("        Origen.Nom_Alm						AS Nom_Ori, ")
            loComandoSeleccionar3.AppendLine("		  Destino.Nom_Alm						AS Nom_Des,")
            loComandoSeleccionar3.AppendLine("        COALESCE(Lotes_Traslados.Cod_Lot,'')	AS Lote,")
            loComandoSeleccionar3.AppendLine("        COALESCE(Lotes_Traslados.Cantidad,0)	AS Cant_Lote,")
            loComandoSeleccionar3.AppendLine("		  COALESCE(Desperdicio.Res_Num,0)		AS Porc_Desperdicio,")
            loComandoSeleccionar3.AppendLine("		  COALESCE((Desperdicio.Res_Num * Renglones_Traslados.Can_Art1)/100,0) AS Cant_Desperdicio,")
            loComandoSeleccionar3.AppendLine("		  COALESCE(Piezas.Res_Num, 0)			AS Piezas,")
            loComandoSeleccionar3.AppendLine("        '" & aString2 & "'                    AS Usuario")
            loComandoSeleccionar3.AppendLine("FROM Traslados ")
            loComandoSeleccionar3.AppendLine("  JOIN Renglones_Traslados ON Traslados.Documento = Renglones_Traslados.Documento")
            loComandoSeleccionar3.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar3.AppendLine("  JOIN Almacenes AS Origen ON Traslados.Alm_Ori = Origen.Cod_Alm")
            loComandoSeleccionar3.AppendLine("	JOIN Almacenes AS Destino ON Traslados.Alm_Des = Destino.Cod_Alm")
            loComandoSeleccionar3.AppendLine("  LEFT JOIN Operaciones_Lotes AS Lotes_Traslados ON Lotes_Traslados.Num_Doc = Renglones_Traslados.Documento")
            loComandoSeleccionar3.AppendLine("      AND Lotes_Traslados.Tip_Doc = 'Traslados' ")
            loComandoSeleccionar3.AppendLine("		AND Lotes_Traslados.Ren_Ori = Renglones_Traslados.Renglon")
            loComandoSeleccionar3.AppendLine("		AND Lotes_Traslados.Tip_Ope = 'Salida'")
            loComandoSeleccionar3.AppendLine("	LEFT JOIN Operaciones_Lotes AS Lotes_Entrada ON Lotes_Traslados.Cod_Lot = Lotes_Entrada.Cod_Lot")
            loComandoSeleccionar3.AppendLine("		AND Lotes_Entrada.Tip_Doc IN ('Ajustes_Inventarios','Recepciones')")
            loComandoSeleccionar3.AppendLine("		AND Lotes_Entrada.Tip_Ope = 'Entrada'")
            loComandoSeleccionar3.AppendLine("		AND Lotes_Entrada.Cod_Art = Lotes_Traslados.Cod_Art")
            loComandoSeleccionar3.AppendLine("	LEFT JOIN Mediciones ON Lotes_Entrada.Num_Doc = Mediciones.Cod_Reg")
            loComandoSeleccionar3.AppendLine("      AND Mediciones.Origen = Lotes_Entrada.Tip_Doc")
            loComandoSeleccionar3.AppendLine("      AND Mediciones.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar3.AppendLine("      AND Mediciones.Adicional LIKE ('%'+RTRIM(Lotes_Entrada.Cod_Lot)+'%')")
            loComandoSeleccionar3.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar3.AppendLine("		AND Piezas.Cod_Var  IN ('AINV-NPIEZ', 'NREC-NPIEZ')")
            loComandoSeleccionar3.AppendLine("		AND Piezas.Res_Num > 0")
            loComandoSeleccionar3.AppendLine("	LEFT JOIN Renglones_Mediciones AS Desperdicio ON Mediciones.Documento = Desperdicio.Documento")
            loComandoSeleccionar3.AppendLine("		AND Desperdicio.Cod_Var IN ('AINV-PDESP', 'NREC-PDESP')")
            loComandoSeleccionar3.AppendLine("		AND Desperdicio.Res_Num > 0")
            loComandoSeleccionar3.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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