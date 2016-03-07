'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAyudas"
'-------------------------------------------------------------------------------------------'
Partial Class fAyudas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.documento,     ")
            loComandoSeleccionar.AppendLine(" 		CASE ")
            loComandoSeleccionar.AppendLine(" 				WHEN Ayudas.Status = 'I' THEN 'Inactivo' ")
            loComandoSeleccionar.AppendLine(" 				WHEN Ayudas.Status = 'A' THEN 'Activo' ")
            loComandoSeleccionar.AppendLine(" 				WHEN Ayudas.Status = 'S' THEN 'Suspendido' ")
            loComandoSeleccionar.AppendLine(" 		END AS Status, ")
            loComandoSeleccionar.AppendLine(" 		CAST(Ayudas.nom_ayu As Varchar(1000)) As Nom_Ayu_Ayuda,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.niv_ayu,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.orden1,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.orden2,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.orden3,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.encabezado,    ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.pie_pag,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.icono,         ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.imagen1,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.interna,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.raiz,          ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.libre,         ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.sistema,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.modulo,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.seccion,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.opcion,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.referencia,    ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.comentario,    ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.prioridad,     ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.nivel,         ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.importancia,   ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.cod_idi,")
            loComandoSeleccionar.AppendLine(" 		Ayudas.padre,   ")
            loComandoSeleccionar.AppendLine(" 		(Ayudas.video + '		' +(Select Nom_Vid From Videos Where Cod_Vid = Ayudas.video)) As Video,")            
            loComandoSeleccionar.AppendLine(" 		Renglones_Ayudas.Renglon As Renglon_Renglones,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Ayudas.Campo,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Ayudas.Ayuda As Ayuda_Renglones,")
            loComandoSeleccionar.AppendLine(" 		'' As Renglon_Referencias,")
            loComandoSeleccionar.AppendLine(" 		'' As Tipo, ")
            loComandoSeleccionar.AppendLine(" 		'' As Nom_Ayu_Referencias,")
            loComandoSeleccionar.AppendLine(" 		'' As Ayuda_Referencias,")
            loComandoSeleccionar.AppendLine(" 		'' As Renglon_Detalles,")
            loComandoSeleccionar.AppendLine(" 		'' As Nom_Ayu_Detalles,")
            loComandoSeleccionar.AppendLine(" 		'' As Web")
            loComandoSeleccionar.AppendLine(" FROM Ayudas")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Ayudas On Renglones_Ayudas.Documento = Ayudas.Documento")
            loComandoSeleccionar.AppendLine(" WHERE "  & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.documento,     ")
            loComandoSeleccionar.AppendLine(" 		CASE ")
            loComandoSeleccionar.AppendLine(" 				WHEN Ayudas.Status = 'I' THEN 'Inactivo' ")
            loComandoSeleccionar.AppendLine(" 				WHEN Ayudas.Status = 'A' THEN 'Activo' ")
            loComandoSeleccionar.AppendLine(" 				WHEN Ayudas.Status = 'S' THEN 'Suspendido' ")
            loComandoSeleccionar.AppendLine(" 		END AS Status, ")
            loComandoSeleccionar.AppendLine(" 		CAST(Ayudas.nom_ayu As Varchar(1000)) As Nom_Ayu_Ayuda,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.niv_ayu,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.orden1,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.orden2,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.orden3,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.encabezado,    ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.pie_pag,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.icono,         ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.imagen1,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.interna,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.raiz,          ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.libre,         ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.sistema,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.modulo,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.seccion,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.opcion,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.referencia,    ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.comentario,    ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.prioridad,     ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.nivel,         ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.importancia,   ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.cod_idi,")
            loComandoSeleccionar.AppendLine(" 		Ayudas.padre,   ")
            loComandoSeleccionar.AppendLine(" 		(Ayudas.video + '		' +(Select Nom_Vid From Videos Where Cod_Vid = Ayudas.video)) As Video,")            
            loComandoSeleccionar.AppendLine(" 		'' As Renglon_Renglones,")
            loComandoSeleccionar.AppendLine(" 		'' As Campo,")
            loComandoSeleccionar.AppendLine(" 		'' As Ayuda_Renglones,")
            loComandoSeleccionar.AppendLine(" 		Referencias_Ayudas.Renglon As Renglon_Referencias,")
            loComandoSeleccionar.AppendLine(" 		Referencias_Ayudas.Tipo, ")
            loComandoSeleccionar.AppendLine(" 		Referencias_Ayudas.Nom_Ayu As Nom_Ayu_Referencias,")
            loComandoSeleccionar.AppendLine(" 		Referencias_Ayudas.Ayuda As Ayuda_Referencias,")
            loComandoSeleccionar.AppendLine(" 		'' As Renglon_Detalles,")
            loComandoSeleccionar.AppendLine(" 		'' As Nom_Ayu_Detalles,")
            loComandoSeleccionar.AppendLine(" 		'' As Web")
            loComandoSeleccionar.AppendLine(" FROM Ayudas")
            loComandoSeleccionar.AppendLine(" JOIN Referencias_Ayudas On Referencias_Ayudas.Documento = Ayudas.Documento")
            loComandoSeleccionar.AppendLine(" WHERE "  & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.documento,     ")
            loComandoSeleccionar.AppendLine(" 		CASE ")
            loComandoSeleccionar.AppendLine(" 				WHEN Ayudas.Status = 'I' THEN 'Inactivo' ")
            loComandoSeleccionar.AppendLine(" 				WHEN Ayudas.Status = 'A' THEN 'Activo' ")
            loComandoSeleccionar.AppendLine(" 				WHEN Ayudas.Status = 'S' THEN 'Suspendido' ")
            loComandoSeleccionar.AppendLine(" 		END AS Status, ")
            loComandoSeleccionar.AppendLine(" 		CAST(Ayudas.nom_ayu As Varchar(1000)) As Nom_Ayu_Ayuda,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.niv_ayu,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.orden1,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.orden2,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.orden3,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.encabezado,    ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.pie_pag,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.icono,         ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.imagen1,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.interna,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.raiz,          ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.libre,         ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.sistema,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.modulo,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.seccion,       ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.opcion,        ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.referencia,    ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.comentario,    ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.prioridad,     ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.nivel,         ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.importancia,   ")
            loComandoSeleccionar.AppendLine(" 		Ayudas.cod_idi,")
            loComandoSeleccionar.AppendLine(" 		Ayudas.padre,   ")
            loComandoSeleccionar.AppendLine(" 		(Ayudas.video + '		' +(Select Nom_Vid From Videos Where Cod_Vid = Ayudas.video)) As Video,")            
            loComandoSeleccionar.AppendLine(" 		'' As Renglon_Renglones,")
            loComandoSeleccionar.AppendLine(" 		'' As Campo,")
            loComandoSeleccionar.AppendLine(" 		'' As Ayuda_Renglones,")
            loComandoSeleccionar.AppendLine(" 		'' As Renglon_Referencias,")
            loComandoSeleccionar.AppendLine(" 		'' As Tipo, ")
            loComandoSeleccionar.AppendLine(" 		'' As Nom_Ayu_Referencias,")
            loComandoSeleccionar.AppendLine(" 		'' As Ayuda_Referencias,")
            loComandoSeleccionar.AppendLine(" 		Detalles_Ayudas.Renglon As Renglon_Detalles,")
            loComandoSeleccionar.AppendLine(" 		Detalles_Ayudas.Nom_Ayu As Nom_Ayu_Detalles,")
            loComandoSeleccionar.AppendLine(" 		Detalles_Ayudas.Web")
            loComandoSeleccionar.AppendLine(" FROM Ayudas")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Ayudas On Detalles_Ayudas.Documento = Ayudas.Documento")
            loComandoSeleccionar.AppendLine(" WHERE "  & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ")

            Dim loServicios As New cusDatos.goDatos
			
			goDatos.pcNombreAplicativoExterno = "Framework"
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAyudas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfAyudas.ReportSource = loObjetoReporte

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
' CMS: 19/07/2010: Codigo inicial.
'-------------------------------------------------------------------------------------------'