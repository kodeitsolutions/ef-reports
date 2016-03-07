'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAjustes_EntradaSalida_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fAjustes_EntradaSalida_IKP
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte AS CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try	
		
			Dim loConsulta As New StringBuilder()
	
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT 		Renglones_Ajustes.Documento         AS Documento,") 
			loConsulta.AppendLine("				Renglones_Ajustes.Renglon           AS Numero,          ") 
			loConsulta.AppendLine("				Renglones_Ajustes.Cod_Tip           AS Tipo,            ")
			loConsulta.AppendLine("				Renglones_Ajustes.Tipo				AS TipoAjuste,     	")	 
			loConsulta.AppendLine("				Renglones_Ajustes.Cod_Art           AS Codigo,          ") 
			loConsulta.AppendLine("				Renglones_Ajustes.Cod_Alm           AS Almacen,         ") 
			loConsulta.AppendLine("				Renglones_Ajustes.Can_Art1          AS Cantidad,        ")
			loConsulta.AppendLine("				Renglones_Ajustes.Cos_Ult1          AS Costo_Unitario,  ") 
			loConsulta.AppendLine("				Renglones_Ajustes.Mon_Net           AS Costo_Total,     ") 
			loConsulta.AppendLine("				Articulos.Nom_Art                   AS Descripcion,     ") 
			loConsulta.AppendLine("				Articulos.Modelo                    AS Modelo,          ")
			loConsulta.AppendLine("				Articulos.Cod_Uni1                  AS Unidad,          ") 
			loConsulta.AppendLine("				Ajustes.Can_Art1                    AS Cantidad_Total,  ") 
			loConsulta.AppendLine("				Ajustes.Mon_Net                     AS Suma_Costo,      ")
			loConsulta.AppendLine("				Ajustes.Comentario                  AS Observacion,     ") 
			loConsulta.AppendLine("				Ajustes.Fec_Ini		                AS Fec_ini,         ") 
			loConsulta.AppendLine("				Ajustes.Cod_Suc		                AS Cod_Suc		    ") 
			loConsulta.AppendLine("INTO #tmpTemporal")			
			loConsulta.AppendLine("FROM Renglones_Ajustes")		
			loConsulta.AppendLine("   JOIN   Articulos ON Renglones_Ajustes.cod_art = Articulos.cod_art 	")		
			loConsulta.AppendLine("   JOIN   Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento	")		
			loConsulta.AppendLine("WHERE	  " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT	  #tmpTemporal.Documento,")
			loConsulta.AppendLine("		  #tmpTemporal.Numero AS Renglon,")
			loConsulta.AppendLine("		  CASE")
			loConsulta.AppendLine("				WHEN #tmpTemporal.TipoAjuste = 'Entrada' THEN SUM(#tmpTemporal.Cantidad) ElSE 0.00 ")	
			loConsulta.AppendLine("		  END AS Total_Cantidad_Entrada,")
			loConsulta.AppendLine("		  CASE")
			loConsulta.AppendLine("				WHEN #tmpTemporal.TipoAjuste = 'Entrada' THEN SUM(#tmpTemporal.Costo_Total) ElSE 0.00 ")	
			loConsulta.AppendLine("		  END AS Total_Costo_Entrada,")
			loConsulta.AppendLine("		  CASE")
			loConsulta.AppendLine("				WHEN #tmpTemporal.TipoAjuste = 'Salida' THEN SUM(#tmpTemporal.Cantidad) ElSE 0.00 ")	
			loConsulta.AppendLine("		  END AS Total_Cantidad_Salida,")
			loConsulta.AppendLine("		  CASE")
			loConsulta.AppendLine("				WHEN #tmpTemporal.TipoAjuste = 'Salida' THEN SUM(#tmpTemporal.Costo_Total) ElSE 0.00 ")	
			loConsulta.AppendLine("		  END AS Total_Costo_Salida")
			loConsulta.AppendLine("INTO 	#tmpTemporal1")			
			loConsulta.AppendLine("FROM 	#tmpTemporal")	
			loConsulta.AppendLine("GROUP BY #tmpTemporal.documento,#tmpTemporal.numero,#tmpTemporal.TipoAjuste")	
			
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT	#tmpTemporal.*,")	
			loConsulta.AppendLine("		    #tmpTemporal1.Total_Cantidad_Entrada,")	
            loConsulta.AppendLine("		    #tmpTemporal1.Total_Costo_Entrada,")	
            loConsulta.AppendLine("		    #tmpTemporal1.Total_Cantidad_Salida,")	
            loConsulta.AppendLine("		    #tmpTemporal1.Total_Costo_Salida, ")
            loConsulta.AppendLine("         Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("         COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("         Sucursales.direccion                     AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("         Sucursales.telefonos                     AS Telefono_Empresa_Cliente   ")                                    
            loConsulta.AppendLine("FROM 	#tmpTemporal")	
            loConsulta.AppendLine("    JOIN #tmpTemporal1 ON #tmpTemporal1.Renglon = #tmpTemporal.Numero")
            loConsulta.AppendLine("    JOIN Sucursales ON Sucursales.cod_suc  = #tmpTemporal.cod_suc")
            loConsulta.AppendLine("    LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("            AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("            AND campos_propiedades.cod_pro = 'SUC-RIF'")
			loConsulta.AppendLine("")
            
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            
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
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAjustes_EntradaSalida_IKP", laDatosReporte)
	   
			Me.mTraducirReporte(loObjetoReporte)            
			Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfAjustes_EntradaSalida_IKP.ReportSource = loObjetoReporte
			
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
' RJG: 02/08/13: Codigo inicial
'-------------------------------------------------------------------------------------------'
