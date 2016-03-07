'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fExpedientes"
'-------------------------------------------------------------------------------------------'
Partial Class fExpedientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" SELECT CAST(seguimientos as XML) AS seguimiento INTO #xmlData from expedientes")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" SELECT")	
			loComandoSeleccionar.AppendLine("  	expedientes.documento	AS documento,")	
			loComandoSeleccionar.AppendLine("  	expedientes.expediente	AS expediente,")	
			loComandoSeleccionar.AppendLine("  	expedientes.tipo			AS tipo,")	
			loComandoSeleccionar.AppendLine("  	expedientes.clase		AS clase,")	
			loComandoSeleccionar.AppendLine("  	expedientes.status		AS status,")	
			loComandoSeleccionar.AppendLine("  	expedientes.importacion	AS importacion,")	
			loComandoSeleccionar.AppendLine("  	expedientes.cod_log		AS cod_log,")	
			loComandoSeleccionar.AppendLine("  	expedientes.nom_log		AS nom_log,")	
			loComandoSeleccionar.AppendLine("  	expedientes.fec_ini		AS fec_ini,")	
			loComandoSeleccionar.AppendLine("  	expedientes.fec_fin		AS fec_fin,")	
			loComandoSeleccionar.AppendLine("  	expedientes.fec_rec		AS fec_rec,")	
			loComandoSeleccionar.AppendLine("  	expedientes.seguimientos AS seguimientos,")	
			loComandoSeleccionar.AppendLine("  	expedientes.comentario	AS comentario,")	
			loComandoSeleccionar.AppendLine("  	expedientes.prioridad	AS prioridad,")	
			loComandoSeleccionar.AppendLine("  	expedientes.nivel		AS nivel,")	
			loComandoSeleccionar.AppendLine("  ")	
			loComandoSeleccionar.AppendLine("  	re.renglon		AS re_renglon,")	
			loComandoSeleccionar.AppendLine("  	re.tip_doc		AS re_tip_doc,")	
			loComandoSeleccionar.AppendLine("  	re.referencia	AS re_referencia,")	
			loComandoSeleccionar.AppendLine("  	re.detalle		AS re_detalle,")	
			loComandoSeleccionar.AppendLine("  	re.status		AS re_status,")	
			loComandoSeleccionar.AppendLine("  	re.control		AS re_control,")	
			loComandoSeleccionar.AppendLine("  	re.completo		AS re_completo,")	
			loComandoSeleccionar.AppendLine("  	re.fec_ini		AS re_fec_ini,")	
			loComandoSeleccionar.AppendLine("  	re.fec_fin		AS re_fec_fin,")	
			loComandoSeleccionar.AppendLine("  	re.comentario	AS re_comentario,")	
			loComandoSeleccionar.AppendLine("  	re.prioridad	AS re_prioridad,")	
			loComandoSeleccionar.AppendLine("  ")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_renglon,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_tip_tra,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_factura,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_control,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_referencia,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_transporte,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_vehiculo,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_fec_emb,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_lug_emb,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_fec_lle,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_containers,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_bultos,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_cajas,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_unidades,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_peso,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_volumen,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_cod_uni,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_can_art,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_detalle,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_status,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_fec_ini,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_fec_fin,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_comentario,")	
			loComandoSeleccionar.AppendLine(" 	NULL as nombreUnidad,")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" 	NULL as se_renglon,")	
			loComandoSeleccionar.AppendLine(" 	NULL as se_fecha,")	
			loComandoSeleccionar.AppendLine(" 	NULL as se_seguimiento,")	
			loComandoSeleccionar.AppendLine(" 	NULL as se_status,")	
			loComandoSeleccionar.AppendLine("	1 as Tabla")	
			loComandoSeleccionar.AppendLine("  ")	
			loComandoSeleccionar.AppendLine(" FROM expedientes ")	
			loComandoSeleccionar.AppendLine(" JOIN renglones_expedientes AS re")	
			loComandoSeleccionar.AppendLine(" ON expedientes.documento = re.documento ")	
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" UNION ALL")	
			loComandoSeleccionar.AppendLine(" SELECT")	
			loComandoSeleccionar.AppendLine(" 	expedientes.documento	AS documento,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.expediente	AS expediente,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.tipo			AS tipo,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.clase		AS clase,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.status		AS status,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.importacion	AS importacion,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.cod_log		AS cod_log,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.nom_log		AS nom_log,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.fec_ini		AS fec_ini,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.fec_fin		AS fec_fin,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.fec_rec		AS fec_rec,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.seguimientos AS seguimientos,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.comentario	AS comentario,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.prioridad	AS prioridad,")	
			loComandoSeleccionar.AppendLine(" 	expedientes.nivel		AS nivel,")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_renglon,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_tip_doc,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_referencia,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_detalle,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_status,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_control,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_completo,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_fec_ini,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_fec_fin,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_comentario,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_prioridad,")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" 	de.renglon		AS de_renglon,")	
			loComandoSeleccionar.AppendLine(" 	de.tip_tra		AS de_tip_tra,")	
			loComandoSeleccionar.AppendLine(" 	de.factura		AS de_factura,")	
			loComandoSeleccionar.AppendLine(" 	de.control		AS de_control,")	
			loComandoSeleccionar.AppendLine(" 	de.referencia	AS de_referencia,")	
			loComandoSeleccionar.AppendLine(" 	de.transporte	AS de_transporte,")	
			loComandoSeleccionar.AppendLine(" 	de.vehiculo		AS de_vehiculo,")	
			loComandoSeleccionar.AppendLine(" 	de.fec_emb		AS de_fec_emb,")	
			loComandoSeleccionar.AppendLine(" 	de.lug_emb		AS de_lug_emb,")	
			loComandoSeleccionar.AppendLine(" 	de.fec_lle		AS de_fec_lle,")	
			loComandoSeleccionar.AppendLine(" 	de.containers	AS de_containers,")	
			loComandoSeleccionar.AppendLine(" 	de.bultos		AS de_bultos,")	
			loComandoSeleccionar.AppendLine(" 	de.cajas		AS de_cajas,")	
			loComandoSeleccionar.AppendLine(" 	de.unidades		AS de_unidades,")	
			loComandoSeleccionar.AppendLine(" 	de.peso			AS de_peso,")	
			loComandoSeleccionar.AppendLine(" 	de.volumen		AS de_volumen,")	
			loComandoSeleccionar.AppendLine(" 	de.cod_uni		AS de_cod_uni,")	
			loComandoSeleccionar.AppendLine(" 	de.can_art		AS de_can_art,")	
			loComandoSeleccionar.AppendLine(" 	de.detalle		AS de_detalle,")	
			loComandoSeleccionar.AppendLine(" 	de.status		AS de_status,")	
			loComandoSeleccionar.AppendLine(" 	de.fec_ini		AS de_fec_ini,")	
			loComandoSeleccionar.AppendLine(" 	de.fec_fin		AS de_fec_fin,")	
			loComandoSeleccionar.AppendLine(" 	de.comentario	AS de_comentario,")	
			loComandoSeleccionar.AppendLine(" 	unidades.nom_uni as nombreUnidad,")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" 	NULL as se_renglon,")	
			loComandoSeleccionar.AppendLine(" 	NULL as se_fecha,")	
			loComandoSeleccionar.AppendLine(" 	NULL as se_seguimiento,")	
			loComandoSeleccionar.AppendLine(" 	NULL as se_status,")	
			loComandoSeleccionar.AppendLine("	2 as Tabla")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" FROM expedientes ")	
			loComandoSeleccionar.AppendLine(" JOIN detalles_expedientes AS de")	
			loComandoSeleccionar.AppendLine(" ON expedientes.documento = de.documento")	
			loComandoSeleccionar.AppendLine(" LEFT JOIN unidades")	
			loComandoSeleccionar.AppendLine(" ON de.cod_uni = unidades.cod_uni")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" UNION ALL ")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" SELECT")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS documento,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS expediente,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS tipo,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS clase,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS status,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS importacion,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS cod_log,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS nom_log,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS fec_ini,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS fec_fin,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS fec_rec,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS seguimientos,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS comentario,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS prioridad,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS nivel,")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_renglon,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_tip_doc,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_referencia,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_detalle,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_status,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_control,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_completo,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_fec_ini,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_fec_fin,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_comentario,")	
			loComandoSeleccionar.AppendLine(" 	NULL	AS re_prioridad,")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_renglon,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_tip_tra,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_factura,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_control,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_referencia,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_transporte,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_vehiculo,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_fec_emb,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_lug_emb,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_fec_lle,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_containers,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_bultos,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_cajas,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_unidades,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_peso,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_volumen,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_cod_uni,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_can_art,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_detalle,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_status,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_fec_ini,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_fec_fin,")	
			loComandoSeleccionar.AppendLine("  	NULL	AS de_comentario,")	
			loComandoSeleccionar.AppendLine(" 	NULL as nombreUnidad,")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" 	ROW_NUMBER() OVER (ORDER BY D.C.value('@status', 'Varchar(15)') DESC) as se_renglon,")	
			loComandoSeleccionar.AppendLine(" 	D.C.value('@fecha', 'datetime') as se_fecha,")	
			loComandoSeleccionar.AppendLine(" 	D.C.value('@seguimiento', 'Varchar(5000)') as se_seguimiento,")	
			loComandoSeleccionar.AppendLine(" 	D.C.value('@status', 'Varchar(15)') as se_status,")	
			loComandoSeleccionar.AppendLine("	3 as Tabla")	
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" FROM #xmlData")	
			loComandoSeleccionar.AppendLine(" CROSS APPLY seguimiento.nodes('elementos/elemento') D(c) -- recuerda que seguimiento es el nombre del campo donde esta el XML")	
			loComandoSeleccionar.AppendLine(" ORDER BY expedientes.documento DESC ,tabla, re_renglon, de_renglon, se_renglon")
								  
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
					 
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fExpedientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfExpedientes.ReportSource = loObjetoReporte

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
' BAP: 05/08/10: Programacion inicial
'-------------------------------------------------------------------------------------------'