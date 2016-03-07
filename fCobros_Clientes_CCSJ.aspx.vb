Imports System.Data
Partial Class fCobros_Clientes_CCSJ
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Cobros.Status, ")
            loComandoSeleccionar.AppendLine("           Cobros.Recibo, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Cobros.Documento, ")
            loComandoSeleccionar.AppendLine("           Cobros.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cobros.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (Cobros.Mon_Des * -1)       AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           (Cobros.Mon_Ret * -1)       AS  Mon_Ret, ")
            loComandoSeleccionar.AppendLine("           Cobros.Comentario           AS  Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros.Renglon    AS  Ren_Doc, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros.Tip_Doc    AS  Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros.Cod_Tip    AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros.Doc_Ori    AS  Doc_Ori, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cobros.Mon_Net    AS  Mon_NetD, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cobros.Mon_Abo    AS  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' THEN Renglones_Cobros.Mon_Net ELSE (Renglones_Cobros.Mon_Net * -1) END)  AS  Mon_NetD, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' THEN Renglones_Cobros.Mon_Abo ELSE (Renglones_Cobros.Mon_Abo * -1) END)  AS  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Cobros.Fec_Ini, 103)	AS	Fec_Che,	")
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
            loComandoSeleccionar.AppendLine(" FROM      Cobros, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cobros, ")
            loComandoSeleccionar.AppendLine("           Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cobros.Documento    =   Renglones_Cobros.Documento AND ")
            loComandoSeleccionar.AppendLine("           Cobros.Cod_Cli      =   Clientes.Cod_Cli AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loComandoSeleccionar.AppendLine(" SELECT	Cobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Cobros.Status, ")
            loComandoSeleccionar.AppendLine("           Cobros.Recibo, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Cobros.Documento, ")
            loComandoSeleccionar.AppendLine("           Cobros.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cobros.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Cobros.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           0.00                        AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           0.00                        AS  Mon_Ret, ")
            loComandoSeleccionar.AppendLine("           Cobros.Comentario           AS  Comentario, ")
            loComandoSeleccionar.AppendLine("           0                           AS  Ren_Doc, ")
            loComandoSeleccionar.AppendLine("           ''                          AS  Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           ''                          AS  Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           ''                          AS  Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           0.00                        AS  Mon_NetD, ")
            loComandoSeleccionar.AppendLine("           0.00                        AS  Mon_Abo, ")
            loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Detalles_Cobros.Fec_Ini, 103)	AS	Fec_Che,	")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Renglon     AS  Ren_Tip_TMP, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Tip_Ope     AS  Tip_Ope, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Doc_Des     AS  Doc_Des, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Num_Doc     AS  Num_Doc, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Cod_Caj     AS  Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Cod_Ban     AS  Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Cod_Cue     AS  Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Cod_Tar     AS  Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros.Mon_Net     AS  Mon_NetTP, ")
            loComandoSeleccionar.AppendLine("           'TiposPagos'                AS  Tipo, ")
            loComandoSeleccionar.AppendLine("           '1'                         AS  Orden ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos1 ")
            loComandoSeleccionar.AppendLine(" FROM      Cobros, ")
            loComandoSeleccionar.AppendLine("           Detalles_Cobros, ")
            loComandoSeleccionar.AppendLine("           Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cobros.Documento    =   Detalles_Cobros.Documento AND ")
            loComandoSeleccionar.AppendLine("           Cobros.Cod_Cli      =   Clientes.Cod_Cli AND (" & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos1.*, ")
            loComandoSeleccionar.AppendLine("           Cajas.Nom_Caj   AS  Nom_Caj ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos2 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos1 LEFT JOIN Cajas ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos1.Cod_Caj =   Cajas.Cod_Caj ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos2.*, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban   AS  Nom_Ban ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos3 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos2 LEFT JOIN Bancos ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos2.Cod_Ban =   Bancos.Cod_Ban ")


            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos3.*, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Nom_Cue   AS  Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTiposPagos4 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTiposPagos3 LEFT JOIN Cuentas_Bancarias ")
            loComandoSeleccionar.AppendLine("           ON  #tmpTiposPagos3.Cod_Cue =   Cuentas_Bancarias.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpTiposPagos4.*, ")
            loComandoSeleccionar.AppendLine("           Tarjetas.Nom_Tar    AS  Nom_Tar ")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCobros_Clientes_CCSJ", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCobros_Clientes_CCSJ.ReportSource = loObjetoReporte

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
' JJD: 08/11/08: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' JJD: 29/11/08: Continuacion de la programacion											'
'-------------------------------------------------------------------------------------------'
' CMS: 31/08/09: Se multiplico -1 el monto neto y el abonado segun su naturaleza			'
'-------------------------------------------------------------------------------------------'
' RJG: 08/10/09: Ordenado el detalle del pago por número de renglón.						'
'-------------------------------------------------------------------------------------------'
' RJG: 21/10/09: Reenuerados los renglones para mostrar correctamente pagos mezclados por	'
'				 cajas y bancos.															'
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'