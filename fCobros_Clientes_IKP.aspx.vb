Imports System.Data
Partial Class fCobros_Clientes_IKP
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine(" SELECT	Cobros.Cod_Cli, ")
            loConsulta.AppendLine("           Cobros.Status, ")
            loConsulta.AppendLine("           Cobros.Recibo, ")
            loConsulta.AppendLine("           Clientes.Nom_Cli, ")
            loConsulta.AppendLine("           Clientes.Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           Clientes.Dir_Fis, ")
            loConsulta.AppendLine("           Clientes.Telefonos, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Cobros.Documento, ")
            loConsulta.AppendLine("           Cobros.Fec_Ini, ")
            loConsulta.AppendLine("           Cobros.Fec_Fin, ")
            loConsulta.AppendLine("           Cobros.Mon_Bru, ")
            loConsulta.AppendLine("           Cobros.Mon_Imp, ")
            loConsulta.AppendLine("           Cobros.Mon_Net, ")
            loConsulta.AppendLine("           (Cobros.Mon_Des * -1)       AS  Mon_Des, ")
            loConsulta.AppendLine("           (Cobros.Mon_Ret * -1)       AS  Mon_Ret, ")
            loConsulta.AppendLine("           Cobros.Comentario           AS  Comentario, ")
            loConsulta.AppendLine("           Cobros.Cod_Suc              AS  Cod_Suc, ")
            loConsulta.AppendLine("           Renglones_Cobros.Renglon    AS  Ren_Doc, ")
            loConsulta.AppendLine("           Renglones_Cobros.Tip_Doc    AS  Tip_Doc, ")
            loConsulta.AppendLine("           Renglones_Cobros.Cod_Tip    AS  Cod_Tip, ")
            loConsulta.AppendLine("           Renglones_Cobros.Doc_Ori    AS  Doc_Ori, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cobros.Mon_Net    AS  Mon_NetD, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cobros.Mon_Abo    AS  Mon_Abo, ")
            loConsulta.AppendLine("           (CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' THEN Renglones_Cobros.Mon_Net ELSE (Renglones_Cobros.Mon_Net * -1) END)  AS  Mon_NetD, ")
            loConsulta.AppendLine("           (CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' THEN Renglones_Cobros.Mon_Abo ELSE (Renglones_Cobros.Mon_Abo * -1) END)  AS  Mon_Abo, ")
            loConsulta.AppendLine("			CONVERT(NCHAR(10), Cobros.Fec_Ini, 103)	AS	Fec_Che,	")
            loConsulta.AppendLine("           0.00                        AS  Ren_Tip_TMP, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Tip_Ope, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Doc_Des, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Num_Doc, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Cod_Caj, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Cod_Ban, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Cod_Cue, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Cod_Tar, ")
            loConsulta.AppendLine("           0.00                        AS  Mon_NetTP, ")
            loConsulta.AppendLine("           'Documentos'                AS  Tipo, ")
            loConsulta.AppendLine("           '2'                         AS  Orden, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Nom_Caj, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Nom_Ban, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Nom_Cue, ")
            loConsulta.AppendLine("           SPACE(10)                   AS  Nom_Tar ")
            loConsulta.AppendLine(" INTO      #tmpDocumentos ")
            loConsulta.AppendLine(" FROM      Cobros, ")
            loConsulta.AppendLine("           Renglones_Cobros, ")
            loConsulta.AppendLine("           Clientes ")
            loConsulta.AppendLine(" WHERE     Cobros.Documento    =   Renglones_Cobros.Documento AND ")
            loConsulta.AppendLine("           Cobros.Cod_Cli      =   Clientes.Cod_Cli AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loConsulta.AppendLine("")
            loConsulta.AppendLine(" SELECT	  Cobros.Cod_Cli, ")
            loConsulta.AppendLine("           Cobros.Status, ")
            loConsulta.AppendLine("           Cobros.Recibo, ")
            loConsulta.AppendLine("           Clientes.Nom_Cli, ")
            loConsulta.AppendLine("           Clientes.Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           Clientes.Dir_Fis, ")
            loConsulta.AppendLine("           Clientes.Telefonos, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Cobros.Documento, ")
            loConsulta.AppendLine("           Cobros.Fec_Ini, ")
            loConsulta.AppendLine("           Cobros.Fec_Fin, ")
            loConsulta.AppendLine("           Cobros.Mon_Bru, ")
            loConsulta.AppendLine("           Cobros.Mon_Imp, ")
            loConsulta.AppendLine("           Cobros.Mon_Net, ")
            loConsulta.AppendLine("           0.00                        AS  Mon_Des, ")
            loConsulta.AppendLine("           0.00                        AS  Mon_Ret, ")
            loConsulta.AppendLine("           Cobros.Comentario           AS  Comentario, ")
            loConsulta.AppendLine("           Cobros.Cod_Suc              AS  Cod_Suc, ")
            loConsulta.AppendLine("           0                           AS  Ren_Doc, ")
            loConsulta.AppendLine("           ''                          AS  Tip_Doc, ")
            loConsulta.AppendLine("           ''                          AS  Cod_Tip, ")
            loConsulta.AppendLine("           ''                          AS  Doc_Ori, ")
            loConsulta.AppendLine("           0.00                        AS  Mon_NetD, ")
            loConsulta.AppendLine("           0.00                        AS  Mon_Abo, ")
            loConsulta.AppendLine("			CONVERT(NCHAR(10), Detalles_Cobros.Fec_Ini, 103)	AS	Fec_Che,	")
            loConsulta.AppendLine("           Detalles_Cobros.Renglon     AS  Ren_Tip_TMP, ")
            loConsulta.AppendLine("           Detalles_Cobros.Tip_Ope     AS  Tip_Ope, ")
            loConsulta.AppendLine("           Detalles_Cobros.Doc_Des     AS  Doc_Des, ")
            loConsulta.AppendLine("           Detalles_Cobros.Num_Doc     AS  Num_Doc, ")
            loConsulta.AppendLine("           Detalles_Cobros.Cod_Caj     AS  Cod_Caj, ")
            loConsulta.AppendLine("           Detalles_Cobros.Cod_Ban     AS  Cod_Ban, ")
            loConsulta.AppendLine("           Detalles_Cobros.Cod_Cue     AS  Cod_Cue, ")
            loConsulta.AppendLine("           Detalles_Cobros.Cod_Tar     AS  Cod_Tar, ")
            loConsulta.AppendLine("           Detalles_Cobros.Mon_Net     AS  Mon_NetTP, ")
            loConsulta.AppendLine("           'TiposPagos'                AS  Tipo, ")
            loConsulta.AppendLine("           '1'                         AS  Orden ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos1 ")
            loConsulta.AppendLine(" FROM      Cobros, ")
            loConsulta.AppendLine("           Detalles_Cobros, ")
            loConsulta.AppendLine("           Clientes ")
            loConsulta.AppendLine(" WHERE     Cobros.Documento    =   Detalles_Cobros.Documento AND ")
            loConsulta.AppendLine("           Cobros.Cod_Cli      =   Clientes.Cod_Cli AND (" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")

            loConsulta.AppendLine("")
            loConsulta.AppendLine(" SELECT	#tmpTiposPagos1.*, ")
            loConsulta.AppendLine("           Cajas.Nom_Caj   AS  Nom_Caj ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos2 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos1 LEFT JOIN Cajas ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos1.Cod_Caj =   Cajas.Cod_Caj ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine(" SELECT	#tmpTiposPagos2.*, ")
            loConsulta.AppendLine("           Bancos.Nom_Ban   AS  Nom_Ban ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos3 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos2 LEFT JOIN Bancos ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos2.Cod_Ban =   Bancos.Cod_Ban ")


            loConsulta.AppendLine("")
            loConsulta.AppendLine(" SELECT	#tmpTiposPagos3.*, ")
            loConsulta.AppendLine("           Cuentas_Bancarias.Nom_Cue   AS  Nom_Cue ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos4 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos3 LEFT JOIN Cuentas_Bancarias ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos3.Cod_Cue =   Cuentas_Bancarias.Cod_Cue ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine(" SELECT	#tmpTiposPagos4.*, ")
            loConsulta.AppendLine("           Tarjetas.Nom_Tar    AS  Nom_Tar ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos5 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos4 LEFT JOIN Tarjetas ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos4.Cod_Tar =   Tarjetas.Cod_Tar ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  T.*, ")
            loConsulta.AppendLine("        Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,")
            loConsulta.AppendLine("        COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,")
            loConsulta.AppendLine("        Sucursales.direccion                     AS Direccion_Empresa_Cliente,")
            loConsulta.AppendLine("        Sucursales.telefonos                     AS Telefono_Empresa_Cliente,")
            loConsulta.AppendLine("        Sucursales.fax                           AS Fax_Empresa_Cliente")
            loConsulta.AppendLine("FROM(           ")
            loConsulta.AppendLine("        SELECT    ROW_NUMBER() OVER(PARTITION BY Orden ORDER BY Orden,Ren_Tip_TMP ASC) AS Ren_Tip, * ")
            loConsulta.AppendLine("        FROM      #tmpDocumentos ")
            loConsulta.AppendLine("        UNION ALL ")
            loConsulta.AppendLine("        SELECT    ROW_NUMBER() OVER(PARTITION BY Orden ORDER BY Orden,Ren_Tip_TMP ASC) AS Ren_Tip, * ")
            loConsulta.AppendLine("        FROM      #tmpTiposPagos5 ")
            loConsulta.AppendLine("    ) AS T")
            loConsulta.AppendLine("  JOIN     Sucursales ON Sucursales.cod_suc = T.cod_suc")
            loConsulta.AppendLine("  LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("        AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("        AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine("")
            loConsulta.AppendLine(" ORDER BY  Orden,Ren_Tip_TMP ")
            loConsulta.AppendLine("")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos1")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos2")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos3")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos4")
            'loConsulta.AppendLine("DROP TABLE #tmpDocumentos")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos5")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            
            Dim loServicios As New cusDatos.goDatos()

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCobros_Clientes_IKP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCobros_Clientes_IKP.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 01/08/13: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
