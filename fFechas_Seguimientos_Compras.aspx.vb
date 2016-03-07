'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFechas_Seguimientos_Compras"
'-------------------------------------------------------------------------------------------'
Partial Class fFechas_Seguimientos_Compras
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            
            loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" SELECT CAST(Compras.Seg_Adm as XML) AS seguimiento INTO #xmlData from Compras")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ")	
            
            loComandoSeleccionar.AppendLine(" SELECT    '1' AS Tabla, ")
            loComandoSeleccionar.AppendLine("		    Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Compras.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Compras.Rif = '') THEN Proveedores.Rif ELSE Compras.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Compras.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Compras.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Compras.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Compras.Documento, ")
            loComandoSeleccionar.AppendLine("           Compras.Factura, ")
            loComandoSeleccionar.AppendLine("           Compras.Control, ")
            loComandoSeleccionar.AppendLine("           Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Compras.Fec_Reg, ")
            loComandoSeleccionar.AppendLine("           Compras.Fec_Pag, ")
            loComandoSeleccionar.AppendLine("           Compras.Fec_Rec, ")
            loComandoSeleccionar.AppendLine("           Compras.Fec_Doc, ")
            loComandoSeleccionar.AppendLine("           Compras.Fecha1, ")
            loComandoSeleccionar.AppendLine("           Compras.Fecha2, ")
            loComandoSeleccionar.AppendLine("           Compras.Fecha3, ")
            loComandoSeleccionar.AppendLine("           Compras.Fecha4, ")
            loComandoSeleccionar.AppendLine("           Compras.Fecha5, ")
            loComandoSeleccionar.AppendLine("           Compras.Comentario, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
			loComandoSeleccionar.AppendLine(" 			NULL as se_renglon,")	
			loComandoSeleccionar.AppendLine(" 			NULL as se_fecha,")	
			loComandoSeleccionar.AppendLine(" 			NULL as se_contacto,")	
			loComandoSeleccionar.AppendLine(" 			NULL as se_accion,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_medio,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_comentario,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_prioridad,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_etapa,")
			loComandoSeleccionar.AppendLine(" 			NULL as se_usuario")
			loComandoSeleccionar.AppendLine(" FROM      Compras ")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON (Compras.Cod_Pro  =   Proveedores.Cod_Pro) ")
            loComandoSeleccionar.AppendLine(" JOIN Formas_Pagos ON (Compras.Cod_For =   Formas_Pagos.Cod_For) ")
            loComandoSeleccionar.AppendLine(" JOIN Vendedores ON (Compras.Cod_Ven   =   Vendedores.Cod_Ven) ")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine(" UNION ALL ")	
			loComandoSeleccionar.AppendLine(" ")	
			
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" SELECT    '2' AS Tabla, ")
			loComandoSeleccionar.AppendLine("		    NULL AS Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           NULL AS Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           NULL AS Rif, ")
            loComandoSeleccionar.AppendLine("           NULL AS Nit, ")
            loComandoSeleccionar.AppendLine("           NULL AS Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           NULL AS Telefonos, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fax, ")
            loComandoSeleccionar.AppendLine("           NULL AS Documento, ")
            loComandoSeleccionar.AppendLine("           NULL AS Factura, ")
            loComandoSeleccionar.AppendLine("           NULL AS Control, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fec_Reg, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fec_Pag, ")
            loComandoSeleccionar.AppendLine("			NULL AS Fec_Rec, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fec_Doc, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fecha1, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fecha2, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fecha3, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fecha4, ")
            loComandoSeleccionar.AppendLine("           NULL AS Fecha5, ")
            loComandoSeleccionar.AppendLine("           NULL AS Comentario, ")
            loComandoSeleccionar.AppendLine("           NULL AS Nom_For, ")
            loComandoSeleccionar.AppendLine("           NULL AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           NULL AS Nom_Ven, ")	
			loComandoSeleccionar.AppendLine(" 			ROW_NUMBER() OVER (ORDER BY D.C.value('@status', 'Varchar(15)') DESC) as se_renglon,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@fecha', 'datetime') as se_fecha,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@contacto', 'Varchar(300)') as se_contacto,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@accion', 'Varchar(300)') as se_accion,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@medio', 'Varchar(300)') as se_medio,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@comentario', 'Varchar(5000)') as se_comentario,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@prioridad', 'Varchar(300)') as se_prioridad,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@etapa', 'Varchar(300)') as se_etapa,")	
			loComandoSeleccionar.AppendLine(" 			D.C.value('@usuario', 'Varchar(100)') as se_usuario")	
			loComandoSeleccionar.AppendLine(" FROM #xmlData")	
			loComandoSeleccionar.AppendLine(" CROSS APPLY seguimiento.nodes('elementos/elemento') D(c) -- recuerda que seguimiento es el nombre del campo donde esta el XML")	
			loComandoSeleccionar.AppendLine(" ORDER BY tabla, compras.documento DESC ,se_renglon")
								  
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
					 
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFechas_Seguimientos_Compras", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFechas_Seguimientos_Compras.ReportSource = loObjetoReporte

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
' MAT: 28/07/11: Programacion inicial
'-------------------------------------------------------------------------------------------'