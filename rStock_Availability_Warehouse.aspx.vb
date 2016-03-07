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
' Inicio de clase "rStock_Availability_Warehouse"
'-------------------------------------------------------------------------------------------'
Partial Class rStock_Availability_Warehouse
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)
            Dim lcParametro12Desde As String = cusAplicacion.goReportes.paParametrosIniciales(12)
            Dim lcExisiencia As String = ""

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Art,")
            loComandoSeleccionar.AppendLine("			Almacenes.Cod_Alm,")
            loComandoSeleccionar.AppendLine("			SUM(")
            loComandoSeleccionar.AppendLine("				CASE")
            loComandoSeleccionar.AppendLine("					WHEN DATEDIFF(day, GETDATE(), Ordenes_Compras.Fec_Fin ) <= 30 THEN Renglones_oCompras.can_pen1")
            loComandoSeleccionar.AppendLine("					ELSE 0")
            loComandoSeleccionar.AppendLine("				END")
            loComandoSeleccionar.AppendLine("			) AS Tran_Less,")
            loComandoSeleccionar.AppendLine("			SUM(")
            loComandoSeleccionar.AppendLine("				CASE")
            loComandoSeleccionar.AppendLine("					WHEN DATEDIFF(day,  GETDATE(), Ordenes_Compras.Fec_Fin ) > 30 THEN Renglones_oCompras.can_pen1")
            loComandoSeleccionar.AppendLine("					ELSE 0")
            loComandoSeleccionar.AppendLine("				END")
            loComandoSeleccionar.AppendLine("			) AS Tran_More")
            loComandoSeleccionar.AppendLine("INTO #tablaTransito")
            loComandoSeleccionar.AppendLine("FROM	Articulos ")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_oCompras ON Articulos.Cod_Art = Renglones_oCompras.Cod_Art")
            loComandoSeleccionar.AppendLine("JOIN	Ordenes_Compras ON Renglones_oCompras.Documento = Ordenes_Compras.Documento")
            loComandoSeleccionar.AppendLine("JOIN 	Departamentos ON Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("JOIN 	Secciones ON Articulos.Cod_Sec = Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("JOIN 	Marcas ON Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("JOIN 	Tipos_Articulos ON Articulos.Cod_Tip = Tipos_Articulos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("JOIN 	Clases_Articulos ON Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine("JOIN 	Almacenes ON Renglones_oCompras.Cod_Alm = Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("WHERE	Secciones.cod_dep = Departamentos.cod_dep ")
            loComandoSeleccionar.AppendLine(" 		AND Renglones_oCompras.Cod_Alm = Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("       AND Ordenes_Compras.Status NOT IN ('Anulado','Procesado')")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Sec = Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine(" 		AND Secciones.cod_dep = Departamentos.cod_dep ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Tip = Tipos_Articulos.Cod_Tip ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Articulos.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 		AND Renglones_oCompras.Cod_Alm BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Departamentos.Cod_Dep BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Secciones.Cod_Sec BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Marcas.Cod_Mar BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Tipos_Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Clases_Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("      	And Articulos.Cod_Ubi between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 		And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("      	And Articulos.Cod_Pro between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 		And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Articulos.Cod_Art,Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("		Articulos.Cod_Uni1, ")

            Select Case lcParametro11Desde
                Case "Actual"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Act1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Act1"
                Case "Comprometida"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Ped1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Ped1"
                Case "Cotizada"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Cot1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Cot1"
                Case "En_Produccion"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Pro1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Pro1"
                Case "Por_Llegar"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Por1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Por1"
                Case "Por_Despachar"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Des1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Des1"
                Case "Por_Distribuir"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Dis1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Dis1"
            End Select

            loComandoSeleccionar.AppendLine("		Renglones_Almacenes.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("		Articulos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine("		Articulos.Cod_Sec, ")
            loComandoSeleccionar.AppendLine("		Articulos.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("		Articulos.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("		Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine("		Articulos.Web, ")
            loComandoSeleccionar.AppendLine("		Articulos.Promocion,")
            loComandoSeleccionar.AppendLine("		Articulos.Foto,")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm,  ")
            loComandoSeleccionar.AppendLine("		ISNULL(#tablaTransito.Tran_Less,0) AS Tran_Less,  ")
            loComandoSeleccionar.AppendLine("		ISNULL(#tablaTransito.Tran_More,0) AS Tran_More,  ")
            loComandoSeleccionar.AppendLine("       ISNULL(Unidades_Articulos.Can_Uni,Articulos.Can_Uni) AS Can_Uni")
            loComandoSeleccionar.AppendLine("FROM	Articulos")
            loComandoSeleccionar.AppendLine("JOIN  	Renglones_Almacenes ON Articulos.Cod_Art = Renglones_Almacenes.Cod_Art ")
            loComandoSeleccionar.AppendLine("JOIN 	Departamentos ON Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("JOIN 	Secciones ON Articulos.Cod_Sec = Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("JOIN 	Marcas ON Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("JOIN 	Tipos_Articulos ON Articulos.Cod_Tip = Tipos_Articulos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("JOIN 	Clases_Articulos ON Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine("JOIN 	Almacenes ON Renglones_Almacenes.Cod_Alm = Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("LEFT OUTER JOIN Unidades_Articulos ON Articulos.Cod_Art = Unidades_Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("LEFT OUTER JOIN #tablaTransito ON (Articulos.Cod_Art = #tablaTransito.Cod_Art AND Almacenes.Cod_Alm = #tablaTransito.Cod_Alm)")
            loComandoSeleccionar.AppendLine("WHERE	Secciones.cod_dep = Departamentos.cod_dep ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Articulos.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Departamentos.Cod_Dep BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Secciones.Cod_Sec BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Marcas.Cod_Mar BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Tipos_Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Clases_Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("      	And Articulos.Cod_Ubi between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 		And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("      	And Articulos.Cod_Pro between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 		And " & lcParametro9Hasta)
            If lcParametro12Desde = "Si" Then
                loComandoSeleccionar.AppendLine(" 		    And Cast(Articulos.Foto As VARCHAR) <> ''")
            End If

            Select Case lcParametro10Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("      ")
                Case "Igual"
                    loComandoSeleccionar.AppendLine("     AND " & lcExisiencia & "          =   0  ")
                Case "Mayor"
                    loComandoSeleccionar.AppendLine("     AND " & lcExisiencia & "          >   0  ")
                Case "Menor"
                    loComandoSeleccionar.AppendLine("     AND " & lcExisiencia & "          <   0  ")
                Case "Maximo"
                    loComandoSeleccionar.AppendLine("     AND Articulos.Exi_Max           =   " & lcExisiencia & "  ")
                Case "Minimo"
                    loComandoSeleccionar.AppendLine("     And Articulos.Exi_Min           =   " & lcExisiencia & "  ")
                Case "Pedido"
                    loComandoSeleccionar.AppendLine("     And Articulos.Exi_pto           =   " & lcExisiencia & "  ")
            End Select

            loComandoSeleccionar.AppendLine("ORDER BY   Articulos.Cod_Art, " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rStock_Availability_Warehouse", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrStock_Availability_Warehouse.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message & " StackTrace: " & loExcepcion.StackTrace, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "600px", _
                           "500px")

        End Try
    End Sub
    
    Protected Sub mCargarFoto(ByRef loTabla As DataTable)

        ' Si la tabla no tiene registros
        If loTabla.Rows.Count <= 0 Then Return

        'Codigo del articulo
        Dim lcCod_Art As String = ""

        ' Se redimensiona la imagen 
        Dim loImage As Bitmap = Me.mRedimensionarImagen(MapPath(Me.pcLogoEmpresa), 50, 50)
        ' se carga en memoria
        Dim loMemory As MemoryStream = New MemoryStream()
        loImage.Save(loMemory, Imaging.ImageFormat.Jpeg)
        ' se guarda la imagen en un arreglo de byte
        Dim loImageByteEmpresa As Byte() = loMemory.GetBuffer()
        ' se inicializa la imagen de producto
        Dim loImageByte As Byte() = loImageByteEmpresa

        ' Recorriendo los registros de la tabla
        For j As Integer = 0 To (loTabla.Rows.Count - 1)

            'si el codigo del articulo cambia
            If loTabla.Rows(j).Item("Cod_Art").ToString <> lcCod_Art Then

                lcCod_Art = loTabla.Rows(j).Item("Cod_Art").ToString

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
                                loImage = Me.mRedimensionarImagen(MapPath("../../Administrativo/Complementos/" & goCliente.pcCodigo & "/" & goEmpresa.pcCodigo & "/" & lcMatrizNombres(i).ToString), 50, 50)
                                ' se carga en memoria
                                loMemory = New MemoryStream()
                                loImage.Save(loMemory, Imaging.ImageFormat.Jpeg)
                                ' se guarda la imagen en un arreglo de byte
                                loImageByte = loMemory.GetBuffer()
                                ' se escribe en la tabla de registro
                                loTabla.Rows(j).Item("FotoImagen") = loImageByte

                            End If

                        Next

                    End If
                Else

                    ' se escribe en la tabla de registro
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
' DLC:  16/07/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 23/09/2010: - Se ajusto la selección por rango de fecha para determinar si esta 
'                   en el rango de menos de 30 días o mayores a 30 dias.
'                   - Se ajusto la visualización, correspondiendo a los datos enviados.
'                   (Si las columnas no aparece nada es porque no hay en stock ni en ordenes de compra)
'                   - La selección de los articulos en las ordenes de compra se basa en 
'                   el estatu de la orden, si esta anulada o procesada se descarta.
'-------------------------------------------------------------------------------------------' 
' MAT: 04/02/2011: Correción y mantenimiento del Reporte
'-------------------------------------------------------------------------------------------'