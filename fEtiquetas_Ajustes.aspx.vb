Imports System.Data
Imports System.Drawing.Printing 

Partial Class fEtiquetas_Ajustes
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loComandoSeleccionar As New StringBuilder()

		loComandoSeleccionar.AppendLine("SELECT		renglones_ajustes.cod_art AS cod_art,")
		loComandoSeleccionar.AppendLine("			renglones_ajustes.can_art1 AS can_art1,")
		loComandoSeleccionar.AppendLine("			articulos.nom_art AS nom_art")
		loComandoSeleccionar.AppendLine("FROM		ajustes,")
		loComandoSeleccionar.AppendLine("			renglones_ajustes,")
		loComandoSeleccionar.AppendLine("			articulos")
		loComandoSeleccionar.AppendLine("WHERE		ajustes.documento = renglones_ajustes.documento")
		loComandoSeleccionar.AppendLine("		AND	renglones_ajustes.cod_art = articulos.cod_art")
		loComandoSeleccionar.AppendLine("		AND ")
		loComandoSeleccionar.AppendLine(			cusAplicacion.goFormatos.pcCondicionPrincipal)
		loComandoSeleccionar.AppendLine("ORDER BY	Renglones_Ajustes.Cod_Art")


        Try


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")
			Dim lcCadena As New StringBuilder()
			
			For Each loRenglon As DataRow In laDatosReporte.Tables(0).Rows
				
				lcCadena.AppendLine("N")
				lcCadena.AppendLine("B10,20,0,1,3,5,125,B,""" & CStr(loRenglon.Item("cod_art")).Trim() & """")
				lcCadena.AppendLine("A10,183,0,3,1,1,N,""" & Left(CStr(loRenglon.Item("nom_art")).Trim(),30)  & """")
				lcCadena.AppendLine("P" & CStr(loRenglon.Item("can_art1")))
				
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

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 29/09/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' RJG: 03/10/08: Ajuste para trabajar con impresora de etiquetas.
'-------------------------------------------------------------------------------------------'