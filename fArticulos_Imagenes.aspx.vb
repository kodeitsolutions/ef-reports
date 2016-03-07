'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.IO
'Imports System.Data
Imports System.Drawing
Imports System.Drawing.Imaging
'Imports System.Runtime.InteropServices

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fArticulos_Imagenes"
'-------------------------------------------------------------------------------------------'
Partial Class fArticulos_Imagenes
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
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT  TOP 0")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Status, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Upc, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Modelo, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Talla, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Informacion, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Ancho, ")
            loComandoSeleccionar.AppendLine("			Articulos.Web,")
            loComandoSeleccionar.AppendLine(" 			Articulos.Alto, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Fondo, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Peso, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Volumen, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Garantia, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Uni1, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Color, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Precio1, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Precio2, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Precio3, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Precio4, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Precio5, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Por_Imp, ")
            loComandoSeleccionar.AppendLine(" 			CAST(0 AS INTEGER)	AS Mon_Imp,")
            loComandoSeleccionar.AppendLine(" 			Articulos.Exi_Act1, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Imp As Cod_Imp, ")
	        loComandoSeleccionar.AppendLine(" 			Articulos.Comentario, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.notas, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Foto ")
            loComandoSeleccionar.AppendLine(" FROM Articulos ")
            loComandoSeleccionar.AppendLine(" JOIN Departamentos ON Departamentos.Cod_Dep = Articulos.Cod_Dep ")
            loComandoSeleccionar.AppendLine(" JOIN Secciones ON Departamentos.Cod_Dep = Secciones.Cod_Dep AND Secciones.Cod_Sec = Articulos.Cod_Sec ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Marcas ON Marcas.Cod_Mar = Articulos.Cod_Mar ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Proveedores ON Proveedores.Cod_Pro = Articulos.Cod_Pro ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Impuestos ON Impuestos.Cod_Imp = Articulos.Cod_Imp ")
            loComandoSeleccionar.AppendLine(" WHERE     Articulos.Cod_Art           Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Sec       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar       Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Tip       Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Cla       Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("      	    And Articulos.Cod_Ubi between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		    And " & lcParametro7Hasta)
            If lcParametro8Desde = "Si"	 Then
				loComandoSeleccionar.AppendLine(" 		And Cast(Articulos.Foto As VARCHAR) <> ''") 
			End If
			loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)
            
         
			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
			
			laDatosReporte.Tables(0).Columns.Add("Foto2", getType(String))
			laDatosReporte.Tables(0).Columns.Add("FotoImagen", getType(Byte()))

            Dim lcXml As String = "<foto></foto>"
            Dim lcFoto As String = ""
            Dim lnNumeroImagenes As Integer = 0
            Dim loFotos As New System.Xml.XmlDocument()
            Dim lcTipo As String = "Porcentaje"
            Dim lcFecha As Date = Now()

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("foto")
                
                laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Mon_Imp")= cusAdministrativo.goImpuestos.mObtenerPorcentaje(laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Cod_Imp"),lcFecha ,10,lcTipo)

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("fArticulos_Imagenes", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfArticulos_Imagenes.ReportSource = loObjetoReporte

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
								loImage = Me.mRedimensionarImagen(MapPath("../../Administrativo/Complementos/" & goCliente.pcCodigo & "/" & goEmpresa.pcCodigo & "/" & lcMatrizNombres(i).ToString),100,100)
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

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' CMS: 25/03/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RJG: 30/03/10: Modificación de la imágen predeterminada (ya no es un subreporte, sino que '
'				 es asignada dinámicamente porque al exportar a excel se mostraban ambas).	'
'-------------------------------------------------------------------------------------------'
' MAT: 14/02/11: Modificación del Select. Ajuste en la vista de Diseño. 
'-------------------------------------------------------------------------------------------'
' MAT: 01/03/11: Hipervinculo para la pág. Web de cada artículo
'-------------------------------------------------------------------------------------------'
' MAT: 05/04/11: Agregado método Redimensionar Imágenes
'-------------------------------------------------------------------------------------------'
