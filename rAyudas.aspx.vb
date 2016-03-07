'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAyudas"
'-------------------------------------------------------------------------------------------'
Partial Class rAyudas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Documento As						Ayuda_Documento,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Nom_Ayu As						Ayuda_Nom_Ayu,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Encabezado As					Ayuda_Encabezado,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Imagen1 As						Ayuda_Imagen1,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Pie_Pag As						Ayuda_Pie_Pag,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Ayudas.Documento As			Campo1,		")
			loComandoSeleccionar.AppendLine(" 			Renglones_Ayudas.Renglon As				Campo2,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Ayudas.Campo As 				Campo3,")
			loComandoSeleccionar.AppendLine(" 			Renglones_Ayudas.Ayuda As 				Campo4,")
			loComandoSeleccionar.AppendLine(" 			'' As									Campo5,")
			loComandoSeleccionar.AppendLine(" 			'1Renglones' As							Tabla")
			loComandoSeleccionar.AppendLine(" FROM Ayudas")
			loComandoSeleccionar.AppendLine(" JOIN Renglones_Ayudas ON Renglones_Ayudas.Documento = Ayudas.Documento")
			'loComandoSeleccionar.AppendLine(" where ayudas.documento IN ('00007779', '00007778') ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Documento As						Ayuda_Documento,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Nom_Ayu As						Ayuda_Nom_Ayu,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Encabezado As					Ayuda_Encabezado,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Imagen1 As						Ayuda_Imagen1,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Pie_Pag As						Ayuda_Pie_Pag,")
			loComandoSeleccionar.AppendLine(" 			Referencias_Ayudas.Documento As			Campo1,")
			loComandoSeleccionar.AppendLine(" 			Referencias_Ayudas.Renglon As			Campo2,")
			loComandoSeleccionar.AppendLine(" 			Referencias_Ayudas.Tipo As				Campo3,")
			loComandoSeleccionar.AppendLine(" 			Referencias_Ayudas.Nom_Ayu As			Campo4,")
			loComandoSeleccionar.AppendLine(" 			Referencias_Ayudas.Ayuda As				Campo5,")
			loComandoSeleccionar.AppendLine(" 			'2Referencias' As							Tabla")
			loComandoSeleccionar.AppendLine(" FROM Ayudas")
			loComandoSeleccionar.AppendLine(" JOIN Referencias_Ayudas ON Referencias_Ayudas.Documento = Ayudas.Documento")
			'loComandoSeleccionar.AppendLine(" where ayudas.documento IN ('00007779', '00007778') ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Documento As						Ayuda_Documento,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Nom_Ayu As						Ayuda_Nom_Ayu,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Encabezado As					Ayuda_Encabezado,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Imagen1 As						Ayuda_Imagen1,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Pie_Pag As						Ayuda_Pie_Pag,")
			loComandoSeleccionar.AppendLine(" 			Detalles_Ayudas.Documento As			Campo1,")
			loComandoSeleccionar.AppendLine(" 			Detalles_Ayudas.Renglon As				Campo2,")
			loComandoSeleccionar.AppendLine(" 			Detalles_Ayudas.Nom_Ayu As				Campo3,")
			loComandoSeleccionar.AppendLine(" 			'' As									Campo4,")
			loComandoSeleccionar.AppendLine(" 			'' As									Campo5,")
			loComandoSeleccionar.AppendLine(" 			'3Detalles' As							Tabla")
			loComandoSeleccionar.AppendLine(" FROM Ayudas")
			loComandoSeleccionar.AppendLine(" JOIN Detalles_Ayudas ON Detalles_Ayudas.Documento = Ayudas.Documento")
			'loComandoSeleccionar.AppendLine(" where ayudas.documento IN ('00007779', '00007778') ")
			loComandoSeleccionar.AppendLine(" ")
			loComandoSeleccionar.AppendLine(" UNION ALL")
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Documento As						Ayuda_Documento,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Nom_Ayu As						Ayuda_Nom_Ayu,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Encabezado As					Ayuda_Encabezado,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Imagen1 As						Ayuda_Imagen1,")
			loComandoSeleccionar.AppendLine(" 			Ayudas.Pie_Pag As						Ayuda_Pie_Pag,")
			loComandoSeleccionar.AppendLine(" 			Ejemplos_Ayudas.Documento As			Campo1,")
			loComandoSeleccionar.AppendLine(" 			Ejemplos_Ayudas.Renglon As				Campo2,")
			loComandoSeleccionar.AppendLine(" 			Ejemplos_Ayudas.Nom_Ayu As				Campo3,")
			loComandoSeleccionar.AppendLine(" 			Ejemplos_Ayudas.Ayuda As				Campo4,")
			loComandoSeleccionar.AppendLine(" 			'' As									Campo5,")
			loComandoSeleccionar.AppendLine(" 			'4Ejemplos' As							Tabla")
			loComandoSeleccionar.AppendLine(" FROM Ayudas")
			loComandoSeleccionar.AppendLine(" JOIN Ejemplos_Ayudas ON Ejemplos_Ayudas.Documento = Ayudas.Documento")
			'loComandoSeleccionar.AppendLine(" where ayudas.documento IN ('00007779', '00007778') ")
			loComandoSeleccionar.AppendLine(" Order BY Ayudas.Documento, Tabla, Campo2")
			
			
			
            Dim loServicios As New cusDatos.goDatos
			
			cusDatos.goDatos.pcNombreAplicativoExterno = "Framework"
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            
			' Se agrega la columna donde se guarda la imagen
			laDatosReporte.Tables(0).Columns.Add("FotoImagen", getType(Byte()))
            

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAyudas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrAyudas.ReportSource = loObjetoReporte

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
    
    
      Protected Sub mCargarFoto(ByRef loTabla As DataTable)

		If loTabla.Rows.Count <= 0 Then Return 
		

		For j As Integer = 0 To (loTabla.Rows.Count - 1)
		Dim s as String = "w"	
			
			If loTabla.Rows(j).Item("Ayuda_Imagen1").ToString <> "" And loTabla.Rows(j).Item("Ayuda_Imagen1").ToString <> "[no disponible]" Then
		
					Dim LcNombreImagen As String = loTabla.Rows(j).Item("Ayuda_Imagen1").ToString
					Dim LnNumeroImagenes As Integer = 1 'cInt(loTabla.Rows(j).Item("FotoImagen").ToString.Substring(0,1))

					Dim lcMatrizNombres AS NEW Arraylist() 
					
					lcMatrizNombres.AddRange(Split(LcNombreImagen,","))
					If LnNumeroImagenes > 0 Then

						For i As Integer = 0 To (lcMatrizNombres.Count - 1)
							
								lcMatrizNombres(i) = lcMatrizNombres(i).ToString.ToUpper.Trim
																					   
								If IO.File.Exists(MapPath("../../Administrativo/Complementos/"& goCliente.pcCodigo &"/"& goEmpresa.pcCodigo &"/" & lcMatrizNombres(i).ToString)) Then
								 
									Dim loArchivoImagen As IO.FileStream = New IO.FileStream(MapPath("../../Administrativo/Complementos/"& goCliente.pcCodigo &"/"& goEmpresa.pcCodigo &"/" & lcMatrizNombres(i).ToString), IO.FileMode.Open, Io.FileAccess.Read)
									Dim loImagenBinaria As Byte()
									ReDim loImagenBinaria(loArchivoImagen.Length)

									loArchivoImagen.Read(loImagenBinaria, 0, CInt(loArchivoImagen.Length))
									loArchivoImagen.Close()
									
									loTabla.Rows(j).Item("FotoImagen") = loImagenBinaria
									
								End If

						Next
					
					End If
			Else
			
				Dim loArchivoImagen As IO.FileStream = New IO.FileStream(MapPath(Me.pcLogoEmpresa), IO.FileMode.Open, Io.FileAccess.Read)
				Dim loImagenBinaria As Byte()
				ReDim loImagenBinaria(loArchivoImagen.Length)

				loArchivoImagen.Read(loImagenBinaria, 0, CInt(loArchivoImagen.Length))
				loArchivoImagen.Close()
				loTabla.Rows(j).Item("Ayuda_Imagen1") = "x"
				'loTabla.Rows(0).Item("FotoImagen") = loImagenBinaria
							
			End If
		Next

      End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS: 27/07/2010: Codigo inicial.
'-------------------------------------------------------------------------------------------'