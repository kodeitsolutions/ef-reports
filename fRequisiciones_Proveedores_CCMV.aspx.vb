Imports System.Data
Partial Class fRequisiciones_Proveedores_CCMV

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Requisiciones.Status    				AS Status, 	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Cod_Pro    				AS Cod_Cli, 	")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro      				AS Nom_Cli, 	")
            loComandoSeleccionar.AppendLine("			Proveedores.Rif          				AS Rif,			")
            loComandoSeleccionar.AppendLine("			Proveedores.Nit          				AS Nit,			")
            loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis      				AS Dir_Fis,		")
            loComandoSeleccionar.AppendLine("			Proveedores.Telefonos    				AS Telefonos,	")
            loComandoSeleccionar.AppendLine("			Proveedores.Fax          				AS Fax,			")
            loComandoSeleccionar.AppendLine("			Proveedores.Contacto          			AS Contacto,	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Nom_Pro    				AS Nom_Gen, 	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Rif        				AS Rif_Gen, 	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Nit        				AS Nit_Gen, 	")
            loComandoSeleccionar.AppendLine("			SPACE(1)                 				AS Dir_Gen, 	")
            loComandoSeleccionar.AppendLine("			SPACE(1)                 				AS Tel_Gen, 	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Documento 				AS Documento, 	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Fec_Ini   				AS Fec_Ini,   	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Fec_Fin   				AS Fec_Fin,   	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Mon_Bru   				AS Mon_Bru,   	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Mon_Imp1  				AS Mon_Imp1,  	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Por_Des1  				AS Por_Des1,  	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Por_Rec1  				AS Por_Rec1,  	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Mon_Des1  				AS Mon_Des1,  	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Mon_Rec1  				AS Mon_Rec1,  	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Dis_Imp   				AS Dis_Imp,   	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Mon_Net   				AS Mon_Net,   	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Cod_For   				AS Cod_For,   	")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Formas_Pagos.Nom_For,1,24)	AS Nom_For, 	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Cod_Tra	  				AS Cod_Tra,   	")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Transportes.Nom_Tra,1,24)		AS Nom_Tra, 	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Cod_Ven               	AS Cod_Ven, 	")
            loComandoSeleccionar.AppendLine("			Requisiciones.Comentario            	AS Comentario,	")
            loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven                  	AS Nom_Ven, 	")

            loComandoSeleccionar.AppendLine("			Auditorias.Cod_Usu                  	AS Elaborado_Por, 	")

            loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Cod_Art     	AS Cod_Art, 	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Articulos.Generico = 'True' THEN Renglones_Requisiciones.Notas ELSE Articulos.Nom_Art END AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Renglon 		AS Renglon,		")
            loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Can_Art1		AS Can_Art1,	")
            loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Cod_Uni 		AS Cod_Uni, 	")
            loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Precio1 		AS Precio1, 	")
            loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Mon_Net 		AS Neto,		")
            loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Por_Imp1		AS Por_Imp, 	")
            loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Cod_Imp 		AS Cod_Imp, 	")
            loComandoSeleccionar.AppendLine("			Renglones_Requisiciones.Mon_Imp1		AS Impuesto, 	")
            loComandoSeleccionar.AppendLine("			0										AS Numero_Revision,")
            loComandoSeleccionar.AppendLine("			GETDATE()								AS Fecha_Revision")
            loComandoSeleccionar.AppendLine(" INTO      #Tmp ")
            loComandoSeleccionar.AppendLine(" FROM      Requisiciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Requisiciones, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Auditorias, ")
            loComandoSeleccionar.AppendLine("           Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE     Requisiciones.Documento =   Renglones_Requisiciones.Documento AND ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Cod_Pro   =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Cod_Ven   =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Cod_For   =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Requisiciones.Cod_Tra	=	Transportes.Cod_Tra   AND ")

            loComandoSeleccionar.AppendLine("           Auditorias.Tabla  =  'Requisiciones'   AND ")
            loComandoSeleccionar.AppendLine("           Auditorias.Accion  =  'Agregar'   AND ")
            loComandoSeleccionar.AppendLine("           Auditorias.Documento  =  Requisiciones.documento   AND ")

            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art       =   Renglones_Requisiciones.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loComandoSeleccionar.AppendLine("SELECT		*, 	")
            loComandoSeleccionar.AppendLine("           Auditorias.Cod_Usu as Aprobado_Por	")
            loComandoSeleccionar.AppendLine("FROM       #Tmp	")
            loComandoSeleccionar.AppendLine("LEFT JOIN  Auditorias  ON (Auditorias.Tabla = 'Requisiciones' And Auditorias.Accion='Confirmar' and Auditorias.Documento=#Tmp.Documento) ")


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            laDatosReporte.Tables(0).Columns.Add("FotoEmpleado1", GetType(Byte()))
            laDatosReporte.Tables(0).Columns.Add("FotoEmpleado2", GetType(Byte()))
            laDatosReporte.Tables(0).Columns.Add("FotoEmpleado3", GetType(Byte()))
			laDatosReporte.Tables(0).AcceptChanges()
			
			Dim lcFirmaActual As String = goUsuario.pcCodigo & ".jpg"
			Dim lcFirmaElaborado As String = CStr(laDatosReporte.Tables(0).Rows(0).Item("Elaborado_Por")).Trim() & ".jpg"
			Dim lcFirmaAprobado As String = CStr(laDatosReporte.Tables(0).Rows(0).Item("Aprobado_Por")).Trim() & ".jpg"

            Me.mCargarFoto(laDatosReporte.Tables(0), "FotoEmpleado1", lcFirmaActual)
            Me.mCargarFoto(laDatosReporte.Tables(0), "FotoEmpleado2", lcFirmaElaborado)
            Me.mCargarFoto(laDatosReporte.Tables(0), "FotoEmpleado3", lcFirmaAprobado)
   
            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fRequisiciones_Proveedores_CCMV", laDatosReporte)

            'CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpuesto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfRequisiciones_Proveedores_CCMV.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                      "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                       vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                       "auto", _
                       "auto")

        End Try

    End Sub

	Protected Sub mCargarFoto(ByRef loTabla As DataTable, ByVal lcNombreCampo As String, ByVal lcNombreImagen As String)


        ' Si la tabla no tiene registros

        If loTabla.Rows.Count <= 0 Then Return


        'En goEmpresa.pcCarpetaComplementos está la ruta de la carpeta de complementos de la empresa actual
		Dim lcFoto As String = Me.MapPath(goEmpresa.pcCarpetaComplementos & "/" & lcNombreImagen)

        ' Busca la imagen
		If Not Io.File.Exists(lcFoto) Then
			lcFoto = Me.MapPath(goEmpresa.pcCarpetaComplementos & "/C0010.jpg" )
		End If
		
        ' Se lee el archivo de la imagen
        Dim loArchivoImagen As IO.FileStream = New IO.FileStream(lcFoto, IO.FileMode.Open, IO.FileAccess.Read)
 
        Dim loImageBytes As Byte()
        ReDim loImageBytes(loArchivoImagen.Length)
        
		loArchivoImagen.Read(loImageBytes, 0, CInt(loArchivoImagen.Length))
 		loArchivoImagen.Close()

        ' va en el encabezado entonces solo lo haríamos en la primera fila de la tabla

        For j As Integer = 0 To (loTabla.Rows.Count - 1)

            ' se escribe en la tabla de registro

            loTabla.Rows(j).Item(lcNombreCampo) = loImageBytes


        Next



    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

        End Try

    End Sub

    Protected Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Load

    End Sub

    Protected Sub form1_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Unload

    End Sub
End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' JJD: 24/03/10: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' JJD: 25/08/10: Ajustes para el formato de fRequisiciones_Proveedores_CCMV.aspx.vb		'
'-------------------------------------------------------------------------------------------'
' RJG: 05/10/11: Adecuacion del formato según especificaciónes.								'
'-------------------------------------------------------------------------------------------'
' RJG: 10/10/11: Agregada etiqueta de "Documento NO Confirmado" segçun el estatus.			'
'-------------------------------------------------------------------------------------------'
