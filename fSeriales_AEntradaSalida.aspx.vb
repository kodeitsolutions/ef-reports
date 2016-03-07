'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSeriales_AEntradaSalida"
'-------------------------------------------------------------------------------------------'
Partial Class fSeriales_AEntradaSalida
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try	
		
		Dim loComandoSeleccionar As New StringBuilder()
		
					loComandoSeleccionar.AppendLine(" SELECT	renglones_ajustes.documento,                             			")
					loComandoSeleccionar.AppendLine("          	renglones_ajustes.cod_tip           as Tipo,             			")
					loComandoSeleccionar.AppendLine("          	renglones_ajustes.cod_art           as Codigo,           			")
					loComandoSeleccionar.AppendLine("          	renglones_ajustes.cod_alm           as Almacen,          			")
					loComandoSeleccionar.AppendLine("          	ajustes.can_art1                    as Cantidad_Total,   			")
					loComandoSeleccionar.AppendLine("          	ajustes.Fec_Ini,")
					loComandoSeleccionar.AppendLine("          	ajustes.Comentario,")
					loComandoSeleccionar.AppendLine("  			Seriales.Renglon AS Renglon_Serial,")
					loComandoSeleccionar.AppendLine("  			Seriales.Cod_Art AS Cod_Art_Serial,")
					loComandoSeleccionar.AppendLine("  			Seriales.Nom_Art AS Nom_Art_Serial,")
					loComandoSeleccionar.AppendLine("  			Seriales.Serial,")
					loComandoSeleccionar.AppendLine("  			Seriales.Alm_Ent,")
					loComandoSeleccionar.AppendLine("  			Seriales.Tip_Ent,")
					loComandoSeleccionar.AppendLine("  			Seriales.Doc_Ent,")
					loComandoSeleccionar.AppendLine("  			Seriales.Ren_Ent,")
					loComandoSeleccionar.AppendLine("  			Seriales.Alm_Sal,")
					loComandoSeleccionar.AppendLine("  			Seriales.Tip_Sal,")
					loComandoSeleccionar.AppendLine("  			Seriales.Doc_Sal,")
					loComandoSeleccionar.AppendLine("  			Seriales.Ren_Sal")
					loComandoSeleccionar.AppendLine(" FROM 	Ajustes")
					loComandoSeleccionar.AppendLine(" JOIN Renglones_Ajustes ON Renglones_Ajustes.Documento	=	Ajustes.Documento")
					loComandoSeleccionar.AppendLine(" JOIN Seriales	ON	((Seriales.Doc_Ent	=	Ajustes.Documento")
					loComandoSeleccionar.AppendLine("					AND	 Seriales.tip_ent	=	'Ajustes'")
					loComandoSeleccionar.AppendLine(" 					AND	 Seriales.Ren_Ent	=   Renglones_Ajustes.Renglon")
					loComandoSeleccionar.AppendLine(" 					AND	 Seriales.Ren_Ent	=   Renglones_Ajustes.Renglon")
					loComandoSeleccionar.AppendLine("					AND  Seriales.Cod_Art   =   Renglones_Ajustes.Cod_Art)")
					loComandoSeleccionar.AppendLine("				OR (Seriales.Doc_Sal	=	Ajustes.Documento")
					loComandoSeleccionar.AppendLine("					AND	 Seriales.tip_sal	=	'Ajustes'")
					loComandoSeleccionar.AppendLine("					AND	 Seriales.Ren_Sal	=   Renglones_Ajustes.Renglon")
					loComandoSeleccionar.AppendLine("					AND  Seriales.Cod_Art   =   Renglones_Ajustes.Cod_Art))")
					loComandoSeleccionar.AppendLine(" WHERE	Ajustes.Tipo='Existencia'")					
					loComandoSeleccionar.AppendLine("       AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
					loComandoSeleccionar.AppendLine("ORDER BY Seriales.Cod_Art,renglones_ajustes.cod_tip,Seriales.Renglon ASC")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSeriales_AEntradaSalida", laDatosReporte)
	   
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfSeriales_AEntradaSalida.ReportSource = loObjetoReporte
			
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
		
			loObjetoReporte.Close ()
		
		Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

	End Sub
	
End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS:  03/03/10 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT:  24/03/11 : Ajuste del Select, mejora de la vista de diseño							'
'-------------------------------------------------------------------------------------------'