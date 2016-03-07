Imports System.Data
Imports System.Drawing.Printing 

Partial Class rEtiquetas_Articulos
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
        Dim loComandoSeleccionar As New StringBuilder()

		loComandoSeleccionar.AppendLine("SELECT		articulos.cod_art AS cod_art,")
		loComandoSeleccionar.AppendLine("			articulos.nom_art AS nom_art")
		loComandoSeleccionar.AppendLine("FROM		articulos")
        loComandoSeleccionar.AppendLine("WHERE		articulos.cod_art BETWEEN ")
        loComandoSeleccionar.AppendLine(lcParametro0Desde)
        loComandoSeleccionar.AppendLine(" AND ")
        loComandoSeleccionar.AppendLine(lcParametro0Hasta)
        'loComandoSeleccionar.AppendLine("ORDER BY	articulos.Cod_Art")
        loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


        Try


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			Dim lcCadena As New StringBuilder()
			Dim lnCantidad As Integer
			If (cusAplicacion.goReportes.paParametrosIniciales(1) IsNot Nothing)
			
				lnCantidad = Math.Max(CInt(cusAplicacion.goReportes.paParametrosIniciales(1)), 1)
				
			Else
			
				lnCantidad = 1
				
			End If
			
			For Each loRenglon As DataRow In laDatosReporte.Tables(0).Rows
				
				lcCadena.AppendLine("N")
				lcCadena.AppendLine("B10,20,0,1,3,5,125,B,""" & CStr(loRenglon.Item("cod_art")).Trim() & """")
				lcCadena.AppendLine("A10,183,0,3,1,1,N,""" & Left(CStr(loRenglon.Item("nom_art")).Trim(),30)  & """")
				lcCadena.AppendLine("P" & CStr(lnCantidad))
				
			Next loREnglon
			
			lcCadena.AppendLine("N")
			
			
			Dim loConfiguracionImpresora	As New System.Xml.XmlDocument()
			Dim lcImpresoraEtiquetas		As String 
			
			loConfiguracionImpresora.Load(Request.MapPath("~/Administrativo/Xml/xmlImpresoraFiscalAdministrativo.xml"))
			lcImpresoraEtiquetas = loConfiguracionImpresora.SelectSingleNode("impresoraFiscal/otroHardware").Attributes("puerto").InnerText 
			
			
			IO.File.WriteAllText("C:\Zebra.txt",lcCadena.ToString()) 
			shell("PRINT C:\Zebra.txt /D:" & lcImpresoraEtiquetas)
			

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")
			Return
			
        End Try
        
        Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso Terminado", _
                      "Las etiquetas seleccionadas fueron enviadas a la impresora satisfactoriamente."  , _
                       vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                       "auto", _
                       "auto")

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 29/09/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' RJG: 03/10/08: Ajuste para trabajar con impresora de etiquetas.
'-------------------------------------------------------------------------------------------'
' CMS: 14/07/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'