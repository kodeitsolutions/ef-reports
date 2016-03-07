Imports System.Data
Imports System.Drawing.Printing 

Partial Class fEtiquetas_fCompras
    Inherits vis1Controles.frmBase

''' <summary>
''' Almacena la estructura simplificada de una imagen monocromática de mapa de bits.
''' </summary>
''' <remarks></remarks>
	Private Structure strImagen
	
		Dim lnAncho		As Integer 
		Dim lnAlto		As Integer 
		Dim laBytes()	As Byte
		
	''' <summary>
	''' Devuelve una representación de cadena de la imagen.
	''' </summary>
	''' <returns></returns>
	''' <remarks></remarks>
		Overrides Function ToString() As String
			
			If	laBytes IsNot Nothing	AndAlso _
				laBytes.Length > 0		Then 
				Return Encoding.GetEncoding(1252).GetString(laBytes)
			Else 
				Return ""
			End If		 			
			
		End Function
		
	End Structure

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loComandoSeleccionar As New StringBuilder()

		loComandoSeleccionar.AppendLine("SELECT		Renglones_Compras.Cod_Art		AS Cod_Art,")
		loComandoSeleccionar.AppendLine("			Renglones_Compras.Can_Art1		AS Can_Art1,")
		loComandoSeleccionar.AppendLine("			Articulos.Nom_Art				AS Nom_Art,")
		loComandoSeleccionar.AppendLine("			Articulos.Modelo				AS Modelo")
		loComandoSeleccionar.AppendLine("FROM		Compras")
		loComandoSeleccionar.AppendLine("	JOIN	Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
		loComandoSeleccionar.AppendLine("	JOIN	Articulos ON Articulos.Cod_Art = Renglones_Compras.Cod_Art")
		loComandoSeleccionar.AppendLine("WHERE		")
		loComandoSeleccionar.AppendLine(			cusAplicacion.goFormatos.pcCondicionPrincipal)
		loComandoSeleccionar.AppendLine("ORDER BY	Renglones_Compras.Cod_Art")


        Try

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")
			
		'Carga la imagen en un objeto strImagen
			Dim lcArchivoImagen As String = Strings.Trim(goOpciones.mObtener("ARCLOGMONP","")) '"Logo_Guarida_BN.bmp"
			lcArchivoImagen = "~/Administrativo/Complementos/" & goCliente.pcCodigo & "/" & goEmpresa.pcCodigo & "/" & lcArchivoImagen
			lcArchivoImagen = Me.Server.MapPath(lcArchivoImagen)
			Dim loImagen As strImagen = Me.mGenerarImagen(lcArchivoImagen)			
			
		'Prepara los textos para imprimir
			Dim loCadena As New StringBuilder()
			Dim lcImpresora As String = "\\rjg\LP2824Plus"
			
			For Each loRenglon As DataRow In laDatosReporte.Tables(0).Rows
				
				Dim lcCodigo As String =  CStr(loRenglon.Item("cod_art")).Trim()
				Dim lcNombre As String =  Strings.Left(CStr(loRenglon.Item("nom_art")).Trim(), 18)
				Dim lcModelo As String =  Strings.Left(CStr(loRenglon.Item("modelo")).Trim(), 18)
				'lcModelo = "PRM-01"
				Dim lnIzq1	As Integer = (448 - lcCodigo.Length * 1.7*8)/2		'Centrado
				Dim lnIzq2	As Integer = (448 - lcNombre.Length * 1.7*8) - 32	'Alineado a la izquierda
				Dim lnIzq3	As Integer = (448 - lcModelo.Length * 1.7*8) - 32	'Alineado a la izquierda
				Dim lnCantidad	As Integer = Math.Ceiling(CDec(loRenglon.Item("Can_Art1")))

				If (lnIzq1 < 5) Then lnIzq1 = 5			  
				If (lnIzq2 < 5) Then lnIzq2 = 5
				If (lnIzq3 < 5) Then lnIzq3 = 5
								
				
				loCadena.AppendLine("N")
				loCadena.AppendLine(Me.mImprimirBarra(10, 10, 0, "1", 2, 8, 50, False, lcCodigo))
				loCadena.AppendLine(Me.mImprimirImagen(10, 80, loImagen))			
				loCadena.AppendLine(Me.mImprimirTexto(lnIzq1, 65,	0, "3", 1, 1, False, lcCodigo))
				loCadena.AppendLine(Me.mImprimirTexto(lnIzq2, 120,	0, "3", 1, 1, False, lcNombre))
				loCadena.AppendLine(Me.mImprimirTexto(lnIzq3, 90,	0, "4", 1, 1, False, lcModelo))
				'loCadena.AppendLine("P1")
				loCadena.AppendLine("P" & lnCantidad.ToString())
				
				'Exit For 
				
			Next loRenglon
			
			loCadena.AppendLine("N")
			
			Dim lcArchivo As String = "~/Administrativo/Temporales/" & System.Guid.NewGuid().ToString("D") & ".txt"
			lcArchivo = Me.Server.MapPath(lcArchivo)
			
			System.IO.File.WriteAllText(lcArchivo, loCadena.ToString(), System.Text.Encoding.Default)  
			shell("PRINT " & lcArchivo & " /D:" & lcImpresora)
		
		
			ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), "CerrarPagina", "window.top.close();", True)
			
        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

	Private Function mImprimirBarra(ByVal lnTop As Integer, ByVal lnLeft As Integer, ByVal lnRotacion As Integer, ByVal lcTipoCodigo As String, _
									ByVal lnAncho1 As Integer, ByVal lnAncho2 As Integer, ByVal lnAlto As Integer, _
									ByVal llImprimirTexto As Boolean, ByVal lcCodigo As String) As String
		
		Dim loComando As New StringBuilder()
		
		loComando.Append("B")
		loComando.Append(lnTop.ToString())
		loComando.Append(",")
		loComando.Append(lnLeft.ToString())
		loComando.Append(",")
		loComando.Append(lnRotacion.ToString())
		loComando.Append(",")
		loComando.Append(lcTipoCodigo.ToString())
		loComando.Append(",")
		loComando.Append(lnAncho1.ToString())
		loComando.Append(",")
		loComando.Append(lnAncho2.ToString())
		loComando.Append(",")
		loComando.Append(lnAlto.ToString())
		loComando.Append(",")
		loComando.Append(IIf(llImprimirTexto, "B", "N"))
		loComando.Append(",")
		loComando.Append("""" & lcCodigo.Replace("\", "\\").Replace("""", "\""") & """")
		
		Return loComando.ToString()
		
	End Function
	
	Private Function mImprimirTexto(ByVal lnLeft As Integer, ByVal lnTop As Integer, ByVal lnRotacion As Integer, ByVal lcFuente As String, _
									ByVal lnEscalaHorizontal As Integer, ByVal lnEscalaVertical As Integer, _
									ByVal llInvertirImagen As Boolean, ByVal lcTexto As String) As String
		
		Dim loComando As New StringBuilder()
		
		loComando.Append("A")
		loComando.Append(lnLeft.ToString())
		loComando.Append(",")
		loComando.Append(lnTop.ToString())
		loComando.Append(",")
		loComando.Append(lnRotacion.ToString())
		loComando.Append(",")
		loComando.Append(lcFuente.ToString())
		loComando.Append(",")
		loComando.Append(lnEscalaHorizontal.ToString())
		loComando.Append(",")
		loComando.Append(lnEscalaVertical.ToString())
		loComando.Append(",")
		loComando.Append(IIf(llInvertirImagen, "R", "N"))
		loComando.Append(",")
		loComando.Append("""" & lcTexto.Replace("\", "\\").Replace("""", "\""") & """")
		
		Return loComando.ToString()
		
	End Function
	
''' <summary>
''' Devuelve la cadena de comando para imprimir la imagen indicada.  
''' </summary>
''' <param name="lnLeft">Distancia, en puntos, desde el borde izquierdo del papel.</param>
''' <param name="lnTop">Distancia, en puntos, desde el borde superior del papel.</param>
''' <param name="loImagen">Objeto strImagen con los bytes y el tamaño de la imagen a imprimir.</param>
''' <returns></returns>
''' <remarks></remarks>
	Private Function mImprimirImagen(lnLeft AS Integer, lnTop As Integer, loImagen As strImagen) As String
		
		Return String.Format("GW{0},{1},{2},{3},{4}" & vbLf, lnLeft.ToString(), lnTop.ToString(), _
			loImagen.lnAncho.ToString(), _
			loImagen.lnAlto.ToString(), _
			Encoding.GetEncoding(1252).GetString(loImagen.laBytes))

	End Function
	
''' <summary>
''' Devuelve un objeto strImagen para ser usado por mImprimirImagen. 
''' </summary>
''' <param name="lcArchivo">ruta completa al archivo de imagen a cargar.</param>
''' <returns></returns>
''' <remarks></remarks>
	Private Function mGenerarImagen(lcArchivo As String) As strImagen
				
		Dim loBitMap As Bitmap   = Drawing.Image.FromFile(lcArchivo)			
		Dim loConversor As New Drawing.ImageConverter()
		
		Dim laBits As Imaging.BitmapData = loBitMap.LockBits(New Rectangle(0, 0, loBitMap.Width, loBitMap.Height), Imaging.ImageLockMode.[ReadOnly], loBitMap.PixelFormat)

		Dim imageBytes As Byte() = New Byte(laBits.Height * laBits.Stride - 1) {}
		System.Runtime.InteropServices.Marshal.Copy(laBits.Scan0, imageBytes, 0, laBits.Stride * laBits.Height)


		'Dim realWidth As Integer = laBits.Width / 8
		''only works for fixed size 1bpp images where width % 8 == 0
		'If realWidth <> laBits.Stride Then
		'	Dim bytesToClear As Integer = laBits.Stride - realWidth
		'	For i As Integer = 0 To laBits.Height - 1
		'		Dim pos As Integer = realWidth + laBits.Stride * i
		'		Dim counter As Integer = 0
		'		While counter < bytesToClear
		'			imageBytes(pos) = 255
		'			counter += 1
		'			pos += 1
		'		End While
		'	Next
		'End If
		
		'Dim realWidth As Integer = laBits.Width / 8
		If laBits.Width <> laBits.Stride * 8 Then
			Dim bitsToClear As Integer = laBits.Stride*8 - laBits.Width
			For i As Integer = 0 To laBits.Height - 1
				Dim counter As Integer = bitsToClear
				Dim pos As Integer = (laBits.Width / 8) + laBits.Stride * i
				While (counter > 0)
					If (counter>=8)
						imageBytes(pos) = 255
					Else
						imageBytes(pos) = ( imageBytes(pos) Or (CInt(Math.Pow(2, counter))-1) ) 
					End If
					
					counter -= 8					
					pos -= 1
				End While
			Next
		End If
		
		Dim loImagen As strImagen
		
		loImagen.lnAncho = laBits.Stride
		loImagen.lnAlto = laBits.Height
		loImagen.laBytes = imageBytes
		loBitMap.UnlockBits(laBits)
		
		Return loImagen

	End Function

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 18/04/2011: Programacion inicial. 
'-------------------------------------------------------------------------------------------'
