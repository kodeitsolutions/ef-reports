'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAjuste_CantidadesIngles"
'-------------------------------------------------------------------------------------------'
Partial Class fAjuste_CantidadesIngles
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try	
		
		
		Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("SELECT	  renglones_ajustes.documento,                            	")		
			loComandoSeleccionar.AppendLine("         renglones_ajustes.renglon           as Numero,          	")		
			loComandoSeleccionar.AppendLine("         renglones_ajustes.cod_tip           as Tipo,            	")	
			loComandoSeleccionar.AppendLine("         renglones_ajustes.Tipo			  as TipoAjuste,       	")		
			loComandoSeleccionar.AppendLine("         renglones_ajustes.cod_art           as Codigo,          	")		
			loComandoSeleccionar.AppendLine("         renglones_ajustes.cod_alm           as Almacen,         	")		
			loComandoSeleccionar.AppendLine("         renglones_ajustes.can_art1          as Cantidad,        	")		
			loComandoSeleccionar.AppendLine("         renglones_ajustes.cos_ult1          as Costo_Unitario,  	")		
			loComandoSeleccionar.AppendLine("         renglones_ajustes.mon_net           as Costo_Total,     	")		
			loComandoSeleccionar.AppendLine("         articulos.nom_art                   as Descripcion,     	")		
			loComandoSeleccionar.AppendLine("         articulos.modelo                    as Modelo,          	")		
			loComandoSeleccionar.AppendLine("         articulos.cod_uni1                  as Unidad,          	")		
			loComandoSeleccionar.AppendLine("         ajustes.can_art1                    as Cantidad_Total,  	")		
			loComandoSeleccionar.AppendLine("         ajustes.mon_net                     as Suma_Costo,      	")
			loComandoSeleccionar.AppendLine("         ajustes.comentario                  as Observacion,      	")		
			loComandoSeleccionar.AppendLine("         ajustes.Fec_Ini	                  as Fec_ini	      	")	
			loComandoSeleccionar.AppendLine("INTO #tmpTemporal")			
			loComandoSeleccionar.AppendLine("FROM 	renglones_ajustes")		
			loComandoSeleccionar.AppendLine("JOIN   articulos ON renglones_ajustes.cod_art = articulos.cod_art 	")		
			loComandoSeleccionar.AppendLine("JOIN   ajustes ON renglones_ajustes.documento = ajustes.documento	")		
			loComandoSeleccionar.AppendLine("WHERE	  " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			
			loComandoSeleccionar.AppendLine("SELECT	  #tmpTemporal.documento,")
			loComandoSeleccionar.AppendLine("		  #tmpTemporal.numero AS Renglon,")
			loComandoSeleccionar.AppendLine("		  CASE")
			loComandoSeleccionar.AppendLine("				WHEN #tmpTemporal.TipoAjuste = 'Entrada' THEN SUM(#tmpTemporal.Cantidad) ElSE 0.00 ")	
			loComandoSeleccionar.AppendLine("		  END AS Total_Cantidad_Entrada,")
			loComandoSeleccionar.AppendLine("		  CASE")
			loComandoSeleccionar.AppendLine("				WHEN #tmpTemporal.TipoAjuste = 'Entrada' THEN SUM(#tmpTemporal.Costo_Total) ElSE 0.00 ")	
			loComandoSeleccionar.AppendLine("		  END AS Total_Costo_Entrada,")
			loComandoSeleccionar.AppendLine("		  CASE")
			loComandoSeleccionar.AppendLine("				WHEN #tmpTemporal.TipoAjuste = 'Salida' THEN SUM(#tmpTemporal.Cantidad) ElSE 0.00 ")	
			loComandoSeleccionar.AppendLine("		  END AS Total_Cantidad_Salida,")
			loComandoSeleccionar.AppendLine("		  CASE")
			loComandoSeleccionar.AppendLine("				WHEN #tmpTemporal.TipoAjuste = 'Salida' THEN SUM(#tmpTemporal.Costo_Total) ElSE 0.00 ")	
			loComandoSeleccionar.AppendLine("		  END AS Total_Costo_Salida")
			loComandoSeleccionar.AppendLine("INTO 	#tmpTemporal1")			
			loComandoSeleccionar.AppendLine("FROM 	#tmpTemporal")	
			loComandoSeleccionar.AppendLine("GROUP BY #tmpTemporal.documento,#tmpTemporal.numero,#tmpTemporal.TipoAjuste")	
			
			loComandoSeleccionar.AppendLine("SELECT		#tmpTemporal.*,")	
			loComandoSeleccionar.AppendLine("			#tmpTemporal1.Total_Cantidad_Entrada,")	
            loComandoSeleccionar.AppendLine("			#tmpTemporal1.Total_Costo_Entrada,")	
            loComandoSeleccionar.AppendLine("			#tmpTemporal1.Total_Cantidad_Salida,")	
            loComandoSeleccionar.AppendLine("			#tmpTemporal1.Total_Costo_Salida")	                                    
            loComandoSeleccionar.AppendLine("FROM 	#tmpTemporal")	
            loComandoSeleccionar.AppendLine("JOIN  #tmpTemporal1 ON #tmpTemporal1.Renglon = #tmpTemporal.Numero")	
            
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAjuste_CantidadesIngles", laDatosReporte)
	   
			Me.mTraducirReporte(loObjetoReporte)            
			Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfAjuste_CantidadesIngles.ReportSource = loObjetoReporte
			
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
' CMS :  29/04/10 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT :  24/08/11 : Ajuste del Select, Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
' MAT :  15/09/11: Eliminación del Pie de Página del eFactory según requerimiento
'-------------------------------------------------------------------------------------------'