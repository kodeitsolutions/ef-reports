'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLista_DDetalles"
'-------------------------------------------------------------------------------------------'
Partial Class rLista_DDetalles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			
            loComandoSeleccionar.AppendLine(" SELECT ")	  
            loComandoSeleccionar.AppendLine(" 			Cod_Com,")
            loComandoSeleccionar.AppendLine(" 			Nom_Com,")
            loComandoSeleccionar.AppendLine(" 			Clave,")
            loComandoSeleccionar.AppendLine(" 			CASE ")
            loComandoSeleccionar.AppendLine(" 					WHEN Status = 'I' THEN 'Inactivo' ")
            loComandoSeleccionar.AppendLine(" 					WHEN Status = 'A' THEN 'Activo' ")
            loComandoSeleccionar.AppendLine(" 					WHEN Status = 'S' THEN 'Suspendido' ")
            loComandoSeleccionar.AppendLine(" 			END AS Status, ")
            loComandoSeleccionar.AppendLine(" 			Contenido,")
            loComandoSeleccionar.AppendLine(" 			'' AS numero,")
            loComandoSeleccionar.AppendLine(" 			'' AS texto,")
            loComandoSeleccionar.AppendLine(" 			'' AS valor,")
            loComandoSeleccionar.AppendLine(" 			'' AS clave_renglon")
            loComandoSeleccionar.AppendLine(" FROM	Combos ")

            loComandoSeleccionar.AppendLine(" WHERE     Cod_Com	Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos
            
            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


	'******************************************************************************************
	' Inicio Se Procesa manualmetne los datos
	'******************************************************************************************


		Dim loTabla As New DataTable("curReportes")
		Dim loColumna As DataColumn 
		
		loColumna = New DataColumn("cod_com", getType(String))
		loColumna.MaxLength = 10
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("nom_com", getType(String))
		loColumna.MaxLength = 100
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("clave", getType(String))
		loColumna.MaxLength = 100
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("status", getType(String))
		loColumna.MaxLength = 30
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("numero", getType(String))
		loColumna.MaxLength = 10
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("texto", getType(String))
		loColumna.MaxLength = 1000
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("valor", getType(String))
		loColumna.MaxLength = 1000
		loTabla.Columns.Add(loColumna)
		
		loColumna = New DataColumn("clave_renglon", getType(String))
		loColumna.MaxLength = 1000
		loTabla.Columns.Add(loColumna)


		Dim loNuevaFila As DataRow
		Dim lnTotalFilas AS Integer = laDatosReporte.Tables(0).Rows.Count
		
		For lnNumeroFila As Integer = 0 To lnTotalFilas - 1  
			
			'Extrayendo los valores de la lista
			Dim lcXml As String = "<elementos></elementos>"
            Dim lcItem As String
            Dim lnItemNumero As Integer = 0
            Dim loItems As New System.Xml.XmlDocument()

            lcItem = ""
     
              lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("contenido")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loItems.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loItem As System.Xml.XmlNode In loItems.SelectNodes("elementos/elemento")

                    If lnNumeroFila <= laDatosReporte.Tables(0).Rows.Count - 1 Then
                        
						Dim loFila As DataRow = laDatosReporte.Tables(0).Rows(lnNumeroFila)
						loNuevaFila = loTabla.NewRow()
						loTabla.Rows.Add(loNuevaFila)

						lnItemNumero = lnItemNumero + 1
						
                        loNuevaFila.Item("cod_com")			= Trim(loFila("cod_com"))
						loNuevaFila.Item("nom_com")			= Trim(loFila("nom_com"))
						loNuevaFila.Item("clave")			= Trim(loFila("clave"))
						loNuevaFila.Item("status")			= Trim(loFila("status"))
						loNuevaFila.Item("numero")			= Cstr(lnItemNumero)
						loNuevaFila.Item("texto")			= Trim(loItem.InnerText).ToString
						loNuevaFila.Item("valor")			= Trim(loItem.Attributes("valor").Value).ToString
						loNuevaFila.Item("clave_renglon")	= Trim(loItem.Attributes("clave").Value).ToString
						
						loTabla.AcceptChanges()	                        
                    End If
                Next loItem
                
                lnItemNumero = 0

		Next lnNumeroFila
		
		
		Dim loDatosReporteFinal As New DataSet("curReportes") 
		loDatosReporteFinal.Tables.Add(loTabla)

	'******************************************************************************************
	' Fin Se Procesa manualmetne los datos
	'******************************************************************************************



            ''-------------------------------------------------------------------------------------------------------
            '' Verificando si el select (tabla nº0) trae registros
            ''-------------------------------------------------------------------------------------------------------

            'If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
            '    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
            '                              "No se Encontraron Registros para los Parámetros Especificados. ", _
            '                               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
            '                               "350px", _
            '                               "200px")
            'End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLista_DDetalles", loDatosReporteFinal)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrLista_DDetalles.ReportSource = loObjetoReporte

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
' CMS: 10/04/10: Codigo inicial.
'-------------------------------------------------------------------------------------------'