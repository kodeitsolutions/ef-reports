'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAyuda_Impresa"
'-------------------------------------------------------------------------------------------'
Partial Class fAyuda_Impresa

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("SELECT ")
			loComandoSeleccionar.AppendLine("			Ayudas.Documento						AS	Ayuda_Documento,")
			loComandoSeleccionar.AppendLine("			Ayudas.Nom_Ayu							AS	Ayuda_Nom_Ayu,")
			loComandoSeleccionar.AppendLine("			Ayudas.Encabezado						AS	Ayuda_Encabezado,")
			loComandoSeleccionar.AppendLine("			Ayudas.Imagen1							AS	Ayuda_Imagen1,")
			loComandoSeleccionar.AppendLine("			Ayudas.Pie_Pag							AS	Ayuda_Pie_Pag,")
			loComandoSeleccionar.AppendLine("			Renglones_Ayudas.Documento				AS	Campo1,		")
			loComandoSeleccionar.AppendLine("			Renglones_Ayudas.Renglon				AS	Campo2,")
			loComandoSeleccionar.AppendLine("			Renglones_Ayudas.Campo					AS 	Campo3,")
			loComandoSeleccionar.AppendLine("			Renglones_Ayudas.Ayuda					AS 	Campo4,")
			loComandoSeleccionar.AppendLine("			''										AS	Campo5,")
			loComandoSeleccionar.AppendLine("			'1Renglones'							AS	Tabla")
			loComandoSeleccionar.AppendLine("FROM		Ayudas")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Ayudas ON Renglones_Ayudas.Documento = Ayudas.Documento")
			loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("UNION ALL")
			loComandoSeleccionar.AppendLine("SELECT ")
			loComandoSeleccionar.AppendLine("			Ayudas.Documento					AS	Ayuda_Documento,")
			loComandoSeleccionar.AppendLine("			Ayudas.Nom_Ayu						AS	Ayuda_Nom_Ayu,")
			loComandoSeleccionar.AppendLine("			Ayudas.Encabezado					AS	Ayuda_Encabezado,")
			loComandoSeleccionar.AppendLine("			Ayudas.Imagen1 						AS	Ayuda_Imagen1,")
			loComandoSeleccionar.AppendLine("			Ayudas.Pie_Pag 						AS	Ayuda_Pie_Pag,")
			loComandoSeleccionar.AppendLine("			Referencias_Ayudas.Documento		AS	Campo1,")
			loComandoSeleccionar.AppendLine("			Referencias_Ayudas.Renglon			AS	Campo2,")
			loComandoSeleccionar.AppendLine("			Referencias_Ayudas.Tipo				AS	Campo3,")
			loComandoSeleccionar.AppendLine("			Referencias_Ayudas.Nom_Ayu			AS	Campo4,")
			loComandoSeleccionar.AppendLine("			Referencias_Ayudas.Ayuda			AS	Campo5,")
			loComandoSeleccionar.AppendLine("			'2Referencias'						AS	Tabla")
			loComandoSeleccionar.AppendLine("FROM		Ayudas")
			loComandoSeleccionar.AppendLine("	JOIN	Referencias_Ayudas ON Referencias_Ayudas.Documento = Ayudas.Documento")
			loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("UNION ALL")
			loComandoSeleccionar.AppendLine("SELECT ")
			loComandoSeleccionar.AppendLine("			Ayudas.Documento					AS	Ayuda_Documento,")
			loComandoSeleccionar.AppendLine("			Ayudas.Nom_Ayu						AS	Ayuda_Nom_Ayu,")
			loComandoSeleccionar.AppendLine("			Ayudas.Encabezado					AS	Ayuda_Encabezado,")
			loComandoSeleccionar.AppendLine("			Ayudas.Imagen1						AS	Ayuda_Imagen1,")
			loComandoSeleccionar.AppendLine("			Ayudas.Pie_Pag						AS	Ayuda_Pie_Pag,")
			loComandoSeleccionar.AppendLine("			Detalles_Ayudas.Documento			AS	Campo1,")
			loComandoSeleccionar.AppendLine("			Detalles_Ayudas.Renglon				AS	Campo2,")
			loComandoSeleccionar.AppendLine("			Detalles_Ayudas.Nom_Ayu				AS	Campo3,")
			loComandoSeleccionar.AppendLine("			''									AS	Campo4,")
			loComandoSeleccionar.AppendLine("			''									AS	Campo5,")
			loComandoSeleccionar.AppendLine("			'3Detalles'							AS	Tabla")
			loComandoSeleccionar.AppendLine("FROM		Ayudas")
			loComandoSeleccionar.AppendLine("	JOIN	Detalles_Ayudas ON Detalles_Ayudas.Documento = Ayudas.Documento")
			loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("UNION ALL")
			loComandoSeleccionar.AppendLine("SELECT ")
			loComandoSeleccionar.AppendLine("			Ayudas.Documento					AS	Ayuda_Documento,")
			loComandoSeleccionar.AppendLine("			Ayudas.Nom_Ayu						AS	Ayuda_Nom_Ayu,")
			loComandoSeleccionar.AppendLine("			Ayudas.Encabezado					AS	Ayuda_Encabezado,")
			loComandoSeleccionar.AppendLine("			Ayudas.Imagen1						AS	Ayuda_Imagen1,")
			loComandoSeleccionar.AppendLine("			Ayudas.Pie_Pag						AS	Ayuda_Pie_Pag,")
			loComandoSeleccionar.AppendLine("			Ejemplos_Ayudas.Documento			AS	Campo1,")
			loComandoSeleccionar.AppendLine("			Ejemplos_Ayudas.Renglon				AS	Campo2,")
			loComandoSeleccionar.AppendLine("			Ejemplos_Ayudas.Nom_Ayu				AS	Campo3,")
			loComandoSeleccionar.AppendLine("			Ejemplos_Ayudas.Ayuda				AS	Campo4,")
			loComandoSeleccionar.AppendLine("			''									AS	Campo5,")
			loComandoSeleccionar.AppendLine("			'4Ejemplos'							AS	Tabla")
			loComandoSeleccionar.AppendLine("FROM		Ayudas")
			loComandoSeleccionar.AppendLine("	JOIN	Ejemplos_Ayudas ON Ejemplos_Ayudas.Documento = Ayudas.Documento")
			loComandoSeleccionar.AppendLine("WHERE		" & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("ORDER BY	Ayudas.Documento, Tabla, Campo2")
			
			
			
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAyuda_Impresa", laDatosReporte)

	'Dim LcAnchoLinea As Integer = me.loObjetoReporte.ReportDefinition.ReportObjects.Item("line1").Width
 '           Dim LcAnchoImagen As Integer = me.loObjetoReporte.ReportDefinition.ReportObjects.Item("FotoImagen2").Width
 '           Dim LcPosicion As Decimal = (LcAnchoLinea - LcAnchoImagen)/2
            
 '           me.loObjetoReporte.ReportDefinition.ReportObjects.Item("FotoImagen2").Left =  LcPosicion
				
	'			Me.mEscribirConsulta(LcAnchoLinea.ToString & " - " &LcAnchoImagen)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfAyuda_Impresa.ReportSource = loObjetoReporte
			
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
																					   
								If IO.File.Exists(MapPath("../../FrameWork/Ayuda/"& lcMatrizNombres(i).ToString)) Then
								 
									Dim loArchivoImagen As IO.FileStream = New IO.FileStream(MapPath("../../FrameWork/Ayuda/"& lcMatrizNombres(i).ToString), IO.FileMode.Open, Io.FileAccess.Read)
													   
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
							
			End If
		Next

      End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' CMS: 27/07/2010: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
' RJG: 21/09/2010: Cambiada la imagen por defecto (logo de eFactory). Corrección de unión	'
'				   con renglones: lso renglones son opcionales.								'
'-------------------------------------------------------------------------------------------'
