'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices


'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPedidos_IHC"
'-------------------------------------------------------------------------------------------'
Partial Class fPedidos_IHC

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Pedidos.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Pedidos.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Pedidos.Rif = '') THEN Clientes.Rif ELSE Pedidos.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Pedidos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Pedidos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Pedidos.Telefonos = '') THEN Clientes.Telefonos ELSE Pedidos.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Documento, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Notas AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Comentario As Comentario_Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Mon_Net As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Pedidos, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Pedidos.Documento  =   Renglones_Pedidos.Documento AND ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art  =   Renglones_Pedidos.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
           
			laDatosReporte.Tables(0).Columns.Add("Firma", GetType(Byte()))
			laDatosReporte.Tables(0).Columns.Add("Vacio", GetType(String))
			
			'--------------------------------------------------'
			' Carga la Firma del Cliente					   '
			'--------------------------------------------------'
			Me.mCargarFoto(laDatosReporte.Tables(0))
	  
            
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            
            
            '--------------------------------------------------'
			' Carga la distribución de Impuestos			   '
			'--------------------------------------------------'
            
			Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpuesto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpuesto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
						If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
								If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)<> 0 Then
									lcPorcentajesImpuesto = lcPorcentajesImpuesto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
								End If
						End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpuesto = lcPorcentajesImpuesto & ")"
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace("(,", "(")
            
           
		
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

			
			
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPedidos_IHC", laDatosReporte)
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace(".",",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpuesto.ToString

            
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPedidos_Clientes.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub
    
	Protected Sub mCargarFoto(ByRef loTabla As DataTable)

			Dim lcRuta	As String
			Dim lcFirma	As String
	
			lcRuta = ("~/Administrativo/Complementos/" & Trim(LCase(goCliente.pcCodigo)) & "/" & Trim(goEmpresa.pcCodigo) & "/" )

        Dim lcDocumento As String = cusAplicacion.goFormatos.pcCondicionPrincipal

        lcDocumento = lcDocumento.ToLower()
        lcDocumento = lcDocumento.Replace("(pedidos.documento='", "")
        lcDocumento = lcDocumento.Replace("')", "")
			
		
			Dim lcNombreArchivo = "firma_pedidos_" + lcDocumento + ".png"					
			
			If My.Computer.FileSystem.FileExists(HttpContext.Current.Server.MapPath(lcRuta +  lcNombreArchivo)) Then

				lcFirma = "../../Administrativo/Complementos/" & goCliente.pcCodigo() & "/" & Trim(goEmpresa.pcCodigo) & "/" & lcNombreArchivo
				
			Else
			
				lcFirma = ""
				
			End If		
	
			
			 ' Se redimensiona la imagen 
			Dim loImage As Bitmap = Me.mRedimensionarImagen(MapPath(Me.pcLogoEmpresa), 100, 100)
			' se carga en memoria
			Dim loMemory As MemoryStream = New MemoryStream()
			loImage.Save(loMemory, Imaging.ImageFormat.Jpeg)
			' se guarda la imagen en un arreglo de byte
			Dim loImageByteEmpresa As Byte() = loMemory.GetBuffer()
			' se inicializa la imagen de producto
			Dim loImageByte As Byte() = loImageByteEmpresa
			
			For	lnFila AS Integer = 0 To loTabla.Rows.Count - 1
			
				'Si no se ha guardado firma asociada
				If lcFirma = "" Then 

					loTabla.Rows(lnFila).Item("Vacio") = "Si"
					
					Exit For
					
				Else
					loTabla.Rows(lnFila).Item("Vacio") = "No"
					
				End If
				
				
				' Se redimensiona la imagen
				loImage = Me.mRedimensionarImagen(MapPath(lcFirma),200, 200)
				
				' se carga en memoria
				loMemory = New MemoryStream()
				
				loImage.Save(loMemory, Imaging.ImageFormat.Jpeg)
				
				' se guarda la imagen en un arreglo de byte
				loImageByte = loMemory.GetBuffer()
				
				' se escribe en la tabla de registro
				loTabla.Rows(lnFila).Item("Firma") = loImageByte

			Next lnFila
			
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
                Dim lnTemp As Decimal = loBMP.Height * 2 * lnRatio
                lnNewHeight = lnTemp
            Else
                ' se calcula la relacion de altura para redimensionar
                lnRatio = lnHeight / loBMP.Height
                ' altura de la nueva imagen
                lnNewHeight = lnHeight
                ' se calcula la anchura de la nueva imagen
                Dim lnTemp As Decimal = loBMP.Width * 2 * lnRatio
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
' MAT: 08/06/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 01/09/11: Inclusión de la firma del cliente en el Formato
'-------------------------------------------------------------------------------------------'
