'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPrecios_Almacenes"
'-------------------------------------------------------------------------------------------'
Partial Class rPrecios_Almacenes
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro12Desde As String = cusAplicacion.goReportes.paParametrosIniciales(12)
             Dim lcParametro13Desde As String = cusAplicacion.goReportes.paParametrosIniciales(13) 
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()
            
            loComandoSeleccionar.AppendLine("SELECT     Articulos.Cod_Dep				AS Cod_Dep, ")
            loComandoSeleccionar.AppendLine("			Departamentos.Nom_Dep			AS Nom_Dep, ")
            loComandoSeleccionar.AppendLine("			Renglones_Almacenes.Cod_Alm		AS Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Almacenes.Nom_Alm				AS Nom_Alm, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Art				AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art				AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Almacenes.Exi_Act1	AS Exi_Act1, ")
            loComandoSeleccionar.AppendLine("			Renglones_Almacenes.Exi_Ped1	AS Exi_Ped1, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Foto					AS Foto, ")
            loComandoSeleccionar.AppendLine("		    Articulos.Web					AS Web, ")
            
            Select Case lcParametro12Desde
				Case "Si"
					loComandoSeleccionar.AppendLine("			'Si' 						AS Mostrar,")
				Case "No"
                    loComandoSeleccionar.AppendLine("			'No' 						AS Mostrar,")
				
			End Select

            Select Case lcParametro8Desde
                Case "Si"
                    loComandoSeleccionar.AppendLine("			'Si'						AS Disponible,")
                Case "No"
                    loComandoSeleccionar.AppendLine("			'No'						AS Disponible,")
            End Select

            Select Case lcParametro9Desde
                Case "Precio1"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio1 				AS Precio")
                Case "Precio2"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio2 				AS Precio")
                Case "Precio3"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio3 				AS Precio")
                Case "Precio4"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio4 				AS Precio")
                Case "Precio5"
                    loComandoSeleccionar.AppendLine("			Articulos.Precio5				AS Precio")
            End Select

            loComandoSeleccionar.AppendLine("FROM		Articulos ")
            loComandoSeleccionar.AppendLine("	JOIN 	Departamentos ON (Articulos.Cod_Dep = Departamentos.Cod_Dep )")
            loComandoSeleccionar.AppendLine("	JOIN 	Renglones_Almacenes ON Renglones_Almacenes.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN 	Almacenes  ")
            loComandoSeleccionar.AppendLine("		ON	Almacenes.Cod_Alm = Renglones_Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND Renglones_Almacenes.Cod_Alm	BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("WHERE		Articulos.Cod_Art			BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Dep       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Sec       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Mar       BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Tip       BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Cla       BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Ubi		BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro7Hasta)
            
            If lcParametro13Desde = "Si"	 Then
				loComandoSeleccionar.AppendLine("			AND CAST(Articulos.Foto As VARCHAR) <> ''") 
			End If
			
            If lcParametro10Desde = "Si"	 Then
				loComandoSeleccionar.AppendLine("			AND (Renglones_Almacenes.Exi_Act1 - Renglones_Almacenes.Exi_Ped1) > 0")	
			End If
						
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

		   'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
		   
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
            
            If lcParametro12Desde = "Si"	 Then
            
				   laDatosReporte.Tables(0).Columns.Add("Foto2", GetType(String))
					laDatosReporte.Tables(0).Columns.Add("FotoImagen", GetType(Byte()))

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
						laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Foto2") = lnNumeroImagenes.ToString & lcFoto.ToString

					Next lnNumeroFila

					Me.mCargarFoto(laDatosReporte.Tables(0))
				
				
			End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPrecios_Almacenes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPrecios_Almacenes.ReportSource = loObjetoReporte


        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub
    
    Protected Sub mCargarFoto(ByRef loTabla As DataTable)
    
		' Si la tabla no tiene registros
        If loTabla.Rows.Count <= 0 Then Return

        ' Se redimensiona la imagen 
        Dim loImage As Bitmap = Me.mRedimensionarImagen(MapPath(Me.pcLogoEmpresa), 50, 50)
        ' se carga en memoria
        Dim loMemory As MemoryStream = New MemoryStream()
        loImage.Save(loMemory, Imaging.ImageFormat.Jpeg)
        ' se guarda la imagen en un arreglo de byte
        Dim loImageByteEmpresa As Byte() = loMemory.GetBuffer()

        ' Recorriendo los registros de la tabla
        For j As Integer = 0 To (loTabla.Rows.Count - 1)

            ' Si el registro tiene imagen asociada
            If loTabla.Rows(j).Item("Foto2").ToString <> "" Then

                ' se extrae los nombres de archivo de imagen del registro
                Dim LcNombreImagen As String = loTabla.Rows(j).Item("Foto2").ToString.Substring(1)
                Dim LnNumeroImagenes As Integer = CInt(loTabla.Rows(j).Item("Foto2").ToString.Substring(0, 1))

                Dim lcMatrizNombres As New ArrayList()
                lcMatrizNombres.AddRange(Split(LcNombreImagen, ","))

                ' Si existe archivos de imagen asociado
                If LnNumeroImagenes > 0 Then

						' Recorriendo la lista de archivos de imagenes
						For i As Integer = 0 To (lcMatrizNombres.Count - 1)

							' se eliminan los espacios en blanco
							lcMatrizNombres(i) = lcMatrizNombres(i).ToString.ToUpper.Trim

							' Si existe el archivo de imagen
							If IO.File.Exists(MapPath("../../Administrativo/Complementos/" & goCliente.pcCodigo & "/" & goEmpresa.pcCodigo & "/" & lcMatrizNombres(i).ToString)) Then

								' Se redimensiona la imagen
								loImage = Me.mRedimensionarImagen(MapPath("../../Administrativo/Complementos/" & goCliente.pcCodigo & "/" & goEmpresa.pcCodigo & "/" & lcMatrizNombres(i).ToString),70,70)
								' se carga en memoria
								loMemory = New MemoryStream()
								loImage.Save(loMemory, Imaging.ImageFormat.Jpeg)
								' se guarda la imagen en un arreglo de byte
								Dim loImageByte As Byte() = loMemory.GetBuffer()
								' se escribe en la tabla de registro
								loTabla.Rows(j).Item("FotoImagen") = loImageByte
								
							Else
								

								loTabla.Rows(j).Item("FotoImagen") = loImageByteEmpresa
	                            
							End If

						Next

                Else
                	
							loTabla.Rows(j).Item("FotoImagen") = loImageByteEmpresa
						
                End If
            Else

                ' se escribe en la tabla de registro
                loTabla.Rows(j).Item("FotoImagen") = loImageByteEmpresa

            End If
        Next

    End Sub
    
    Protected Function mRedimensionarImagen(ByVal lcFilename As String, ByVal lnWidth As Integer, ByVal lnHeight As Integer) As Bitmap

        ' Se lee el archivo de la imagen
        Dim loArchivoImagen As IO.FileStream = New IO.FileStream(lcFilename, IO.FileMode.Open, IO.FileAccess.Read)
        ' Se carga la imagen
        Dim loBMP As Bitmap = New Bitmap(loArchivoImagen)
        ' Variable donde se guardar la imagen redimensionada
        Dim bmpOut As Bitmap = New Bitmap(lnWidth, lnHeight)
        Try

            Dim lnRatio As Decimal
            Dim lnNewWidth As Integer = 0
            Dim lnNewHeight As Integer = 0

            ' Si el tamaño de la imagen es menor a la que se quiere redimensionar
            If (loBMP.Width < lnWidth And loBMP.Height < lnHeight) Then
                ' se retorna la imagen original
                Return loBMP
            End If

            ' Si el ancho de la imagen original es mayo que la altura de la imagen original
            If (loBMP.Width > loBMP.Height) Then
                ' se calcula la relacion de anchura para redimensionar
                lnRatio = lnWidth / loBMP.Width 
                ' ancho de la nueva imagen
                lnNewWidth = lnWidth
                ' se calcula la altura de la nueva imagen
                Dim lnTemp As Decimal = loBMP.Height * lnRatio
                lnNewHeight = lnTemp
            Else
                ' se calcula la relacion de altura para redimensionar
                lnRatio = lnHeight / loBMP.Height
                ' altura de la nueva imagen
                lnNewHeight = lnHeight
                ' se calcula la anchura de la nueva imagen
                Dim lnTemp As Decimal = loBMP.Width * lnRatio
                lnNewWidth = lnTemp
            End If

            ' se crea la imagen nueva para redimensionar
            bmpOut = New Bitmap(lnNewWidth, lnNewHeight, loBMP.PixelFormat)
            ' se carga la manipulacion de la imagen
            Dim g As Graphics = Graphics.FromImage(bmpOut)
            ' se estable el modo de interpolacion de la imagen para redimensionar
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            ' se carga el tamaño al que se redimensionara
            g.FillRectangle(Brushes.White, 0, 0, lnNewWidth, lnNewHeight)
            ' se dibuja la imagen redimensionandola
            g.DrawImage(loBMP, 0, 0, lnNewWidth, lnNewHeight)

            loBMP.Dispose()
        Catch
            ' si ocurre un error, retorna la imagen original
            Return loBMP

        End Try
        ' retorna la imagen redimensionada
        Return bmpOut

    End Function

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

			 loObjetoReporte.Close()

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
' MAT: 22/06/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 29/09/11: Ajuste del Select con Filtro Solo con Existencia (Exi_Act1 - Exi_Ped1)
'-------------------------------------------------------------------------------------------'
' RJG: 26/01/12: Ajuste en filtro de artículos con inventario (inventario por almacen).		'
'-------------------------------------------------------------------------------------------'
