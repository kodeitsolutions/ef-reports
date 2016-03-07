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
' Inicio de clase "fPresupuestos_IHP"
'-------------------------------------------------------------------------------------------'
Partial Class fPresupuestos_IHP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Presupuestos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Presupuestos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Presupuestos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Presupuestos.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Presupuestos.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Presupuestos.Rif = '') THEN Proveedores.Rif ELSE Presupuestos.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Presupuestos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Presupuestos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Presupuestos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Presupuestos.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Presupuestos.Telefonos = '') THEN Proveedores.Telefonos ELSE Presupuestos.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Nom_Pro                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Documento, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Presupuestos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Cod_Art, ")
            'loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("		CASE")
			loComandoSeleccionar.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
			loComandoSeleccionar.AppendLine("			ELSE Renglones_Presupuestos.Notas")
			loComandoSeleccionar.AppendLine("		END														AS Nom_Art,  ")            
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Comentario As Comentario_Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuestos.Mon_Imp1         As  Impuesto ")  
			loComandoSeleccionar.AppendLine(" FROM      Presupuestos ")
            loComandoSeleccionar.AppendLine("           JOIN Renglones_Presupuestos on Presupuestos.Documento  =   Renglones_Presupuestos.Documento")
            loComandoSeleccionar.AppendLine("           JOIN Proveedores ON Presupuestos.Cod_Pro    =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           JOIN Formas_Pagos ON Presupuestos.Cod_For    =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Vendedores ON Presupuestos.Cod_Ven    =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           JOIN Articulos ON Articulos.Cod_Art       =   Renglones_Presupuestos.Cod_Art")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
            laDatosReporte.Tables(0).Columns.Add("Firma", GetType(Byte()))
            laDatosReporte.Tables(0).Columns.Add("Vacio", GetType(String))

            '--------------------------------------------------'
            ' Carga la Firma del proveedor					   '
            '--------------------------------------------------'
            Me.mCargarFoto(laDatosReporte.Tables(0))


            '--------------------------------------------------'
            ' Carga la distribución de Impuestos			   '
            '--------------------------------------------------'

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim loImpuestos As New System.Xml.XmlDocument()
       

            lcPorcentajesImpueto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribución de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                    'Verifica si el impuesto es igual a Cero
 						if CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) <> 0 Then
							lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
						End If
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,","(")
            
            if lcPorcentajesImpueto = "()" Then
					lcPorcentajesImpueto = " "
			End If

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

            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPresupuestos_IHP", laDatosReporte)
            
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".",",")
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPresupuestos_IHP.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                      "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                       vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                       "auto", _
                       "auto")

        End Try

    End Sub

    Protected Sub mCargarFoto(ByRef loTabla As DataTable)

        Dim lcRuta As String
        Dim lcFirma As String

        lcRuta = ("~/Administrativo/Complementos/" & Trim(LCase(goCliente.pcCodigo)) & "/" & Trim(goEmpresa.pcCodigo) & "/")

        Dim lcDocumento As String = cusAplicacion.goFormatos.pcCondicionPrincipal

        lcDocumento = lcDocumento.ToLower()
        lcDocumento = lcDocumento.Replace("(presupuestos.documento='", "")
        lcDocumento = lcDocumento.Replace("')", "")


        Dim lcNombreArchivo = "firma_presupuestos_" + lcDocumento + ".png"

        If My.Computer.FileSystem.FileExists(HttpContext.Current.Server.MapPath(lcRuta + lcNombreArchivo)) Then

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

        For lnFila As Integer = 0 To loTabla.Rows.Count - 1

            'Si no se ha guardado firma asociada
            If lcFirma = "" Then

                loTabla.Rows(lnFila).Item("Vacio") = "Si"

                Exit For

            Else
                loTabla.Rows(lnFila).Item("Vacio") = "No"

            End If


            ' Se redimensiona la imagen
            loImage = Me.mRedimensionarImagen(MapPath(lcFirma), 200, 200)

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
' MAT: 19/08/11: Código Inicial
'-------------------------------------------------------------------------------------------'