'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAjustes_EntradaSalida_Imagenes"
'-------------------------------------------------------------------------------------------'
Partial Class fAjustes_EntradaSalida_Imagenes
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

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
			loComandoSeleccionar.AppendLine("				Articulos.Foto						AS Foto,     ") 
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
            
            laDatosReporte.Tables(0).Columns.Add("Foto2", getType(String))
			laDatosReporte.Tables(0).Columns.Add("FotoImagen", getType(Byte()))

            Dim lcXml As String = "<foto></foto>"
            Dim lcFoto As String = ""
            Dim lnNumeroImagenes As Integer = 0
            Dim loFotos As New System.Xml.XmlDocument()

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("foto")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loFotos.LoadXml(lcXml)
				lcFoto = "*"
				lnNumeroImagenes = 0
								
                'En cada renglón lee el contenido de cada imagen
                For Each loFoto As System.Xml.XmlNode In loFotos.SelectNodes("fotos/foto")
                        lcFoto = lcFoto & ", " & loFoto.SelectSingleNode("nombre").InnerText  
                        lnNumeroImagenes = lnNumeroImagenes + 1
                Next loFoto
                
                lcFoto = lcFoto.Replace("*,", "")
                laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Foto2")= lnNumeroImagenes.ToString & lcFoto.ToString
                
            Next lnNumeroFila
            
            

            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			Me.mCargarFoto(laDatosReporte.Tables(0))
			
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
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAjustes_EntradaSalida_Imagenes", laDatosReporte)
	   
			Me.mTraducirReporte(loObjetoReporte)            
			Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfAjustes_EntradaSalida_Imagenes.ReportSource = loObjetoReporte
			
        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub
    
    Protected Sub mCargarFoto(ByRef loTabla As DataTable)

		If loTabla.Rows.Count <= 0 Then Return 
		
		'Dim loFila AS DataRow = loTabla.NewRow()
		
		For j As Integer = 0 To (loTabla.Rows.Count - 1)
		Dim s as String = "w"	
			'loFila.ItemArray() = loTabla.Rows(j).ItemArray()
			
			If loTabla.Rows(j).Item("Foto2").ToString <> "" Then
		
					Dim LcNombreImagen As String = loTabla.Rows(j).Item("Foto2").ToString.Substring(1)
					Dim LnNumeroImagenes As Integer = cInt(loTabla.Rows(j).Item("Foto2").ToString.Substring(0,1))

					Dim lcMatrizNombres AS NEW Arraylist() 
					
					lcMatrizNombres.AddRange(Split(LcNombreImagen,","))
					If LnNumeroImagenes > 0 Then

						For i As Integer = 0 To (lcMatrizNombres.Count - 1)

							'If i >= 1 Then
							'	loTabla.Rows.Add(loFila.ItemArray)
							'	'loTabla.Rows.InsertAt(loFila,j+1)
							'End If
							
								lcMatrizNombres(i) = lcMatrizNombres(i).ToString.ToUpper.Trim
																					   
								If IO.File.Exists(MapPath("../../Administrativo/Complementos/"& goCliente.pcCodigo &"/"& goEmpresa.pcCodigo &"/" & lcMatrizNombres(i).ToString)) Then
								
									Dim loArchivoImagen As IO.FileStream = New IO.FileStream(MapPath("../../Administrativo/Complementos/"& goCliente.pcCodigo &"/"& goEmpresa.pcCodigo &"/" & lcMatrizNombres(i).ToString), IO.FileMode.Open, Io.FileAccess.Read)
									Dim loImagenBinaria As Byte()
									ReDim loImagenBinaria(loArchivoImagen.Length)

									loArchivoImagen.Read(loImagenBinaria, 0, CInt(loArchivoImagen.Length))
									loArchivoImagen.Close()
									
									'If i >= 1 Then
										
									'	loTabla.Rows(j+1).Item("FotoImagen") = loImagenBinaria
									'Else
									
										loTabla.Rows(j).Item("FotoImagen") = loImagenBinaria
										
									'End If

								End If

						Next
					
					End If
			Else
			
				Dim loArchivoImagen As IO.FileStream = New IO.FileStream(MapPath(Me.pcLogoEmpresa), IO.FileMode.Open, Io.FileAccess.Read)
				Dim loImagenBinaria As Byte()
				ReDim loImagenBinaria(loArchivoImagen.Length)

				loArchivoImagen.Read(loImagenBinaria, 0, CInt(loArchivoImagen.Length))
				loArchivoImagen.Close()
				loTabla.Rows(j).Item("FotoImagen") = loImagenBinaria
				'loTabla.Rows(0).Item("FotoImagen") = loImagenBinaria
							
			End If
		Next

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
' MAT :  31/01/11 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT :  24/08/11 : Adición de Totales
'-------------------------------------------------------------------------------------------'
