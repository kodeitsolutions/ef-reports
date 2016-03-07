'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAjustes_EntradaSalida"
'-------------------------------------------------------------------------------------------'
Partial Class fAjustes_EntradaSalida
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte AS CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try	
		
			Dim loComandoSeleccionar As New StringBuilder()
	
			loComandoSeleccionar.AppendLine("SELECT 		Renglones_Ajustes.Documento,") 
			loComandoSeleccionar.AppendLine("				Renglones_Ajustes.Renglon           AS Numero,          ") 
			loComandoSeleccionar.AppendLine("				Renglones_Ajustes.Cod_Tip           AS Tipo,            ")
			loComandoSeleccionar.AppendLine("				Renglones_Ajustes.Tipo				AS TipoAjuste,     	")	 
			loComandoSeleccionar.AppendLine("				Renglones_Ajustes.Cod_Art           AS Codigo,          ") 
			loComandoSeleccionar.AppendLine("				Renglones_Ajustes.Cod_Alm           AS Almacen,         ") 
			loComandoSeleccionar.AppendLine("				Renglones_Ajustes.Can_Art1          AS Cantidad,        ")
			loComandoSeleccionar.AppendLine("				Renglones_Ajustes.Cos_Ult1          AS Costo_Unitario,  ") 
			loComandoSeleccionar.AppendLine("				Renglones_Ajustes.Mon_Net           AS Costo_Total,     ") 
			loComandoSeleccionar.AppendLine("				Articulos.Nom_Art                   AS Descripcion,     ") 
			loComandoSeleccionar.AppendLine("				Articulos.Modelo                    AS Modelo,          ")
			loComandoSeleccionar.AppendLine("				Articulos.Cod_Uni1                  AS Unidad,          ") 
			loComandoSeleccionar.AppendLine("				Ajustes.Can_Art1                    AS Cantidad_Total,  ") 
			loComandoSeleccionar.AppendLine("				Ajustes.Mon_Net                     AS Suma_Costo,      ")
			loComandoSeleccionar.AppendLine("				Ajustes.Comentario                  AS Observacion,     ") 
			loComandoSeleccionar.AppendLine("				Ajustes.Fec_Ini		               AS Fec_ini		    ") 
			loComandoSeleccionar.AppendLine("INTO #tmpTemporal")			
			loComandoSeleccionar.AppendLine("FROM Renglones_Ajustes")		
			loComandoSeleccionar.AppendLine("JOIN   Articulos ON Renglones_Ajustes.cod_art = Articulos.cod_art 	")		
			loComandoSeleccionar.AppendLine("JOIN   Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento	")		
			loComandoSeleccionar.AppendLine("WHERE	  " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			
			loComandoSeleccionar.AppendLine("SELECT	  #tmpTemporal.Documento,")
			loComandoSeleccionar.AppendLine("		  #tmpTemporal.Numero AS Renglon,")
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

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            
            'Me.mEscribirConsulta(lcComandoSelect.ToString())

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
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAjustes_EntradaSalida", laDatosReporte)
	   
			Me.mTraducirReporte(loObjetoReporte)            
			Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfAjustes_EntradaSalida.ReportSource = loObjetoReporte
			
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
' MJP :  07/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP :  11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP :  14/07/08 : Agregacion de filtro Status
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 30/04/10: Se Agrego el campo fecha
'-------------------------------------------------------------------------------------------'
' MAT :  24/08/11 : Ajuste del Select y la vista de diseño
'-------------------------------------------------------------------------------------------'