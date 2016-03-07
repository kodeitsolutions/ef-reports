'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPagos_Proveedores_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fPagos_Proveedores_IKP
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Pagos.Cod_Pro, ")
            loConsulta.AppendLine("           Pagos.Status, ")
            loConsulta.AppendLine("           Proveedores.Nom_Pro, ")
            loConsulta.AppendLine("           Proveedores.Rif, ")
            loConsulta.AppendLine("           Proveedores.Nit, ")
            loConsulta.AppendLine("           Proveedores.Dir_Fis, ")
            loConsulta.AppendLine("           Proveedores.Telefonos, ")
            loConsulta.AppendLine("           Proveedores.Fax, ")
            loConsulta.AppendLine("           Pagos.Documento, ")
            loConsulta.AppendLine("           Pagos.Fec_Ini, ")
            loConsulta.AppendLine("           Pagos.Fec_Fin, ")
            loConsulta.AppendLine("           Pagos.Mon_Bru, ")
            loConsulta.AppendLine("           Pagos.Mon_Imp, ")
            loConsulta.AppendLine("           Pagos.Mon_Net, ")
            loConsulta.AppendLine("           (Pagos.Mon_Des * -1)       AS  Mon_Des, ")
            loConsulta.AppendLine("           (Pagos.Mon_Ret * -1)       AS  Mon_Ret, ")
            loConsulta.AppendLine("           Pagos.Comentario           AS  Comentario, ")
            loConsulta.AppendLine("           Pagos.Cod_Suc              AS  Cod_Suc, ")
            loConsulta.AppendLine("           Renglones_Pagos.Renglon    AS  Ren_Doc, ")
            loConsulta.AppendLine("           Renglones_Pagos.Tip_Doc    AS  Tip_Doc, ")
            loConsulta.AppendLine("           Renglones_Pagos.Cod_Tip    AS  Cod_Tip, ")
            loConsulta.AppendLine("           Renglones_Pagos.Doc_Ori    AS  Doc_Ori, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Net    AS  Mon_NetD, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Abo  AS  Mon_Abo, ")
            loConsulta.AppendLine("           (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Net ELSE (Renglones_Pagos.Mon_Net * -1) END)  AS  Mon_NetD, ")
            loConsulta.AppendLine("           (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Abo ELSE (Renglones_Pagos.Mon_Abo * -1) END)  AS  Mon_Abo, ")
            loConsulta.AppendLine("			CONVERT(NCHAR(10), Pagos.Fec_Ini, 103)	AS	Fec_Che,	")
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
            loConsulta.AppendLine(" FROM      Pagos, ")
            loConsulta.AppendLine("           Renglones_Pagos, ")
            loConsulta.AppendLine("           Proveedores ")
            loConsulta.AppendLine(" WHERE     Pagos.Documento    =   Renglones_Pagos.Documento AND ")
            loConsulta.AppendLine("           Pagos.Cod_Pro      =   Proveedores.Cod_Pro AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            '((Case When (Renglones_Pagos.Tip_Doc = 'FACT' OR Renglones_Pagos.Tip_Doc = 'ATD') Then 1 Else -1 End) * 

            loConsulta.AppendLine(" SELECT	Pagos.Cod_Pro, ")
            loConsulta.AppendLine("           Pagos.Status, ")
            loConsulta.AppendLine("           Proveedores.Nom_Pro, ")
            loConsulta.AppendLine("           Proveedores.Rif, ")
            loConsulta.AppendLine("           Proveedores.Nit, ")
            loConsulta.AppendLine("           Proveedores.Dir_Fis, ")
            loConsulta.AppendLine("           Proveedores.Telefonos, ")
            loConsulta.AppendLine("           Proveedores.Fax, ")
            loConsulta.AppendLine("           Pagos.Documento, ")
            loConsulta.AppendLine("           Pagos.Fec_Ini, ")
            loConsulta.AppendLine("           Pagos.Fec_Fin, ")
            loConsulta.AppendLine("           Pagos.Mon_Bru, ")
            loConsulta.AppendLine("           Pagos.Mon_Imp, ")
            loConsulta.AppendLine("           Pagos.Mon_Net, ")
            loConsulta.AppendLine("           0.00                       AS  Mon_Des, ")
            loConsulta.AppendLine("           0.00                       AS  Mon_Ret, ")
            loConsulta.AppendLine("           Pagos.Comentario           AS  Comentario, ")
            loConsulta.AppendLine("           Pagos.Cod_Suc              AS  Cod_Suc, ")
            loConsulta.AppendLine("           0                          AS  Ren_Doc, ")
            loConsulta.AppendLine("           ''                         AS  Tip_Doc, ")
            loConsulta.AppendLine("           ''                         AS  Cod_Tip, ")
            loConsulta.AppendLine("           ''                         AS  Doc_Ori, ")
            loConsulta.AppendLine("           0.00                       AS  Mon_NetD, ")
            loConsulta.AppendLine("           0.00                       AS  Mon_Abo, ")
            loConsulta.AppendLine("			CONVERT(NCHAR(10), Detalles_Pagos.Fec_Ini, 103)	AS	Fec_Che,	")
            loConsulta.AppendLine("           Detalles_Pagos.Renglon    AS  Ren_Tip_TMP, ")
            loConsulta.AppendLine("           Detalles_Pagos.Tip_Ope    AS  Tip_Ope, ")
            loConsulta.AppendLine("           Detalles_Pagos.Doc_Des    AS  Doc_Des, ")
            loConsulta.AppendLine("           Detalles_Pagos.Num_Doc    AS  Num_Doc, ")
            loConsulta.AppendLine("           Detalles_Pagos.Cod_Caj    AS  Cod_Caj, ")
            loConsulta.AppendLine("           Detalles_Pagos.Cod_Ban    AS  Cod_Ban, ")
            loConsulta.AppendLine("           Detalles_Pagos.Cod_Cue    AS  Cod_Cue, ")
            loConsulta.AppendLine("           Detalles_Pagos.Cod_Tar    AS  Cod_Tar, ")
            loConsulta.AppendLine("           Detalles_Pagos.Mon_Net    AS  Mon_NetTP, ")
            loConsulta.AppendLine("           'TiposPagos'               AS  Tipo, ")
            loConsulta.AppendLine("           '1'                        AS  Orden ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos1 ")
            loConsulta.AppendLine(" FROM      Pagos, ")
            loConsulta.AppendLine("           Detalles_Pagos, ")
            loConsulta.AppendLine("           Proveedores ")
            loConsulta.AppendLine(" WHERE     Pagos.Documento    =   Detalles_Pagos.Documento AND ")
            loConsulta.AppendLine("           Pagos.Cod_Pro      =   Proveedores.Cod_Pro AND (" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")

            loConsulta.AppendLine(" SELECT	#tmpTiposPagos1.*, ")
            loConsulta.AppendLine("           SUBSTRING(Cajas.Nom_Caj,1,25)   AS  Nom_Caj ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos2 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos1 LEFT JOIN Cajas ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos1.Cod_Caj =   Cajas.Cod_Caj ")

            loConsulta.AppendLine(" SELECT	#tmpTiposPagos2.*, ")
            loConsulta.AppendLine("           SUBSTRING(Bancos.Nom_Ban,1,25)   AS  Nom_Ban ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos3 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos2 LEFT JOIN Bancos ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos2.Cod_Ban =   Bancos.Cod_Ban ")


            loConsulta.AppendLine(" SELECT	#tmpTiposPagos3.*, ")
            loConsulta.AppendLine("           SUBSTRING(Cuentas_Bancarias.Nom_Cue,1,25)   AS  Nom_Cue ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos4 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos3 LEFT JOIN Cuentas_Bancarias ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos3.Cod_Cue =   Cuentas_Bancarias.Cod_Cue ")

            loConsulta.AppendLine(" SELECT	#tmpTiposPagos4.*, ")
            loConsulta.AppendLine("           SUBSTRING(Tarjetas.Nom_Tar,1,25)   AS  Nom_Tar ")
            loConsulta.AppendLine(" INTO      #tmpTiposPagos5 ")
            loConsulta.AppendLine(" FROM      #tmpTiposPagos4 LEFT JOIN Tarjetas ")
            loConsulta.AppendLine("           ON  #tmpTiposPagos4.Cod_Tar =   Tarjetas.Cod_Tar ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  T.*, ")
            loConsulta.AppendLine("        Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,")
            loConsulta.AppendLine("        COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,")
            loConsulta.AppendLine("        Sucursales.direccion                     AS Direccion_Empresa_Cliente,")
            loConsulta.AppendLine("        Sucursales.telefonos                     AS Telefono_Empresa_Cliente")
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
            loConsulta.AppendLine("ORDER BY  Orden,Ren_Tip_TMP ")
            loConsulta.AppendLine("")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos1")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos2")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos3")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos4")
            'loConsulta.AppendLine("DROP TABLE #tmpDocumentos")
            'loConsulta.AppendLine("DROP TABLE #tmpTiposPagos5")
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

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
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPagos_Proveedores_IKP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPagos_Proveedores_IKP.ReportSource = loObjetoReporte

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
