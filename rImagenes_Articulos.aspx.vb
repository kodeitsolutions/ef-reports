Imports System.Data

Partial Class rImagenes_Articulos 
    Inherits vis2Formularios.frmReporte
    
	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" ")	
			loComandoSeleccionar.AppendLine("SELECT			Articulos.Cod_Art	AS	Cod_Art,")
			loComandoSeleccionar.AppendLine("				Articulos.Nom_Art	AS	Nom_Art,")
			loComandoSeleccionar.AppendLine("				Articulos.Foto		AS	Archivo")
			loComandoSeleccionar.AppendLine(" FROM Articulos")
			loComandoSeleccionar.AppendLine(" WHERE	Articulos.Cod_Art between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 		AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)
			
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
		   
		    Dim lcXml As String = "<foto></foto>"
            Dim lcFoto As String = ""
            Dim loFotos As New System.Xml.XmlDocument()
		  
			'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Archivo")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loFotos.LoadXml(lcXml)

                'En cada renglón lee el contenido del campo foto
                
                For Each loFoto As System.Xml.XmlNode In loFotos.SelectNodes("fotos/foto")
                
                   
                        lcFoto = lcFoto & (loFoto.SelectSingleNode("nombre").InnerText).Trim & vbCrLf 
                        
				Next loFoto
                
               laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Archivo")= lcFoto
               
               lcFoto = "" 
                
            Next lnNumeroFila
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rImagenes_Articulos", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrImagenes_Articulos.ReportSource = loObjetoReporte
            

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
	
End Class

'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' MAT:  08/04/11: Codigo inicial
'-------------------------------------------------------------------------------------------'