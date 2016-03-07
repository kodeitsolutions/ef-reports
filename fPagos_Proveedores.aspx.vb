'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPagos_Proveedores"
'-------------------------------------------------------------------------------------------'
Partial Class fPagos_Proveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Pagos.Status, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("           Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pagos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Pagos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Pagos.Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Pagos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (Pagos.Mon_Des * -1)       AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           (Pagos.Mon_Ret * -1)       AS  Mon_Ret, ")
            loComandoSeleccionar.AppendLine("           Pagos.Comentario           AS  Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Renglon    AS  Ren_Doc, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Tip_Doc    AS  Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Cod_Tip    AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos.Doc_Ori    AS  Doc_Ori, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Net    AS  Mon_NetD, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Pagos.Mon_Abo  AS  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Net ELSE (Renglones_Pagos.Mon_Net * -1) END)  AS  Mon_NetD, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Renglones_Pagos.Tip_Doc = 'Debito' THEN Renglones_Pagos.Mon_Abo ELSE (Renglones_Pagos.Mon_Abo * -1) END)  AS  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Pagos.Fec_Ini, 103)	AS	Fec_Che,	")
            loComandoSeleccionar.AppendLine("           0.00                        AS  Ren_Tip_TMP, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Tip_Ope, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Doc_Des, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Num_Doc, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           0.00                        AS  Mon_NetTP, ")
            loComandoSeleccionar.AppendLine("           'Documentos'                AS  Tipo, ")
            loComandoSeleccionar.AppendLine("           '2'                         AS  Orden, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Nom_Caj, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Nom_Ban, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Nom_Cue, ")
            loComandoSeleccionar.AppendLine("           SPACE(10)                   AS  Nom_Tar ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpDocumentos ")
            loComandoSeleccionar.AppendLine(" FROM      Pagos, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pagos, ")
            loComandoSeleccionar.AppendLine("           Proveedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Pagos.Documento    =   Renglones_Pagos.Documento AND ")
            loComandoSeleccionar.AppendLine("           Pagos.Cod_Pro      =   Proveedores.Cod_Pro AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            '((Case When (Renglones_Pagos.Tip_Doc = 'FACT' OR Renglones_Pagos.Tip_Doc = 'ATD') Then 1 Else -1 End) * 

            loComandoSeleccionar.AppendLine(" SELECT	Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Pagos.Status, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("           Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pagos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Pagos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Pagos.Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Pagos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           0.00                       AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           0.00                       AS  Mon_Ret, ")
            loComandoSeleccionar.AppendLine("           Pagos.Comentario           AS  Comentario, ")
            loComandoSeleccionar.AppendLine("           0                          AS  Ren_Doc, ")
            loComandoSeleccionar.AppendLine("           ''                         AS  Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           ''                         AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           ''                         AS  Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           0.00                       AS  Mon_NetD, ")
            loComandoSeleccionar.AppendLine("           0.00                       AS  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Detalles_Pagos.Fec_Ini, 103)	AS	Fec_Che,	")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Renglon    AS  Ren_Tip_TMP, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Tip_Ope    AS  Tip_Ope, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Doc_Des    AS  Doc_Des, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Num_Doc    AS  Num_Doc, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Cod_Caj    AS  Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Cod_Ban    AS  Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Cod_Cue    AS  Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Cod_Tar    AS  Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos.Mon_Net    AS  Mon_NetTP, ")
            loComandoSeleccionar.AppendLine("           'TiposPagos'               AS  Tipo, ")
            loComandoSeleccionar.AppendLine("           '1'                        AS  Orden ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos1 ")
            loComandoSeleccionar.AppendLine(" FROM      Pagos, ")
            loComandoSeleccionar.AppendLine("           Detalles_Pagos, ")
            loComandoSeleccionar.AppendLine("           Proveedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Pagos.Documento    =   Detalles_Pagos.Documento AND ")
            loComandoSeleccionar.AppendLine("           Pagos.Cod_Pro      =   Proveedores.Cod_Pro AND (" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos1.*, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Cajas.Nom_Caj,1,25)   AS  Nom_Caj ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos2 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos1 LEFT JOIN Cajas ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos1.Cod_Caj =   Cajas.Cod_Caj ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos2.*, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Bancos.Nom_Ban,1,25)   AS  Nom_Ban ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos3 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos2 LEFT JOIN Bancos ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos2.Cod_Ban =   Bancos.Cod_Ban ")


            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos3.*, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Cuentas_Bancarias.Nom_Cue,1,25)   AS  Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos4 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos3 LEFT JOIN Cuentas_Bancarias ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos3.Cod_Cue =   Cuentas_Bancarias.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos4.*, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Tarjetas.Nom_Tar,1,25)   AS  Nom_Tar ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos5 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos4 LEFT JOIN Tarjetas ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos4.Cod_Tar =   Tarjetas.Cod_Tar ")

            loComandoSeleccionar.AppendLine(" SELECT    ROW_NUMBER() OVER(PARTITION BY Orden ORDER BY Orden,Ren_Tip_TMP ASC) AS Ren_Tip, * ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpDocumentos ")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine(" SELECT    ROW_NUMBER() OVER(PARTITION BY Orden ORDER BY Orden,Ren_Tip_TMP ASC) AS Ren_Tip, * ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos5 ")
            loComandoSeleccionar.AppendLine(" ORDER BY  Orden,Ren_Tip_TMP ")
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPagos_Proveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPagos_Proveedores.ReportSource = loObjetoReporte

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
' JJD: 29/11/08: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' CMS: 31/08/09: se multiplico -1 el monto neto y el abonado segun su naturaleza			'
'-------------------------------------------------------------------------------------------'
' RJG: 21/10/09: Reenuerados los regnlones para mostrar correctamente pagos mezclados por	'
'				 cajas y bancos.															'
'-------------------------------------------------------------------------------------------'
' JJD: 05/12/09: Se le incluyo los montos de descuentos y de retenciones que se encuentran 
'				 en el encabezado de la tabla de Pagos'
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaro el metodos carga de imagen 
'-------------------------------------------------------------------------------------------'
