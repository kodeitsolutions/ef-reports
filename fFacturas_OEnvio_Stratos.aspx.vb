'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFacturas_OEnvio_Stratos"
'-------------------------------------------------------------------------------------------'
Partial Class fFacturas_OEnvio_Stratos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Rif = '') THEN Clientes.Rif ELSE Facturas.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Facturas.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Facturas.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Telefonos = '') THEN Clientes.Telefonos ELSE Facturas.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Clientes.Generico, ")
            loComandoSeleccionar.AppendLine("           Facturas.Documento, ")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Facturas.For_Env, ")
            loComandoSeleccionar.AppendLine("           Facturas.Dir_Ent, ")
            loComandoSeleccionar.AppendLine("           Facturas.Notas, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Transportes.Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Notas END AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Renglon, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Can_Art1")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Can_Art2 END) AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Cod_Uni")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Cod_Uni2 END) AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           RTRIM(LTRIM(Articulos.Garantia)) As Garantia,")            
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN (Renglones_Facturas.Can_Art1 * Articulos.Peso) ")
            loComandoSeleccionar.AppendLine("			    ELSE (Renglones_Facturas.Can_Art2 * Articulos.Peso) END) AS Peso, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN (Renglones_Facturas.Can_Art1 * Articulos.Volumen) ")
            loComandoSeleccionar.AppendLine("			    ELSE (Renglones_Facturas.Can_Art2 * Articulos.Volumen) END) AS Volumen, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Ubi, ")            
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Doc_Ori,")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Tip_Ori,")
            loComandoSeleccionar.AppendLine("           CASE")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Cotizaciones') THEN (Cotizaciones.Comentario)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Pedidos') THEN (Pedidos.Comentario)")
            loComandoSeleccionar.AppendLine("				WHEN (Renglones_Facturas.Tip_Ori = 'Entregas') THEN (Entregas.Comentario)")
            loComandoSeleccionar.AppendLine("           ELSE ' '")
            loComandoSeleccionar.AppendLine("           END AS Comentario_Origen")
            loComandoSeleccionar.AppendLine(" FROM      Facturas ")
            loComandoSeleccionar.AppendLine(" JOIN   Renglones_Facturas ON (Facturas.Documento	=   Renglones_Facturas.Documento)")
			loComandoSeleccionar.AppendLine(" LEFT JOIN  Clientes ON (Facturas.Cod_Cli		=   Clientes.Cod_Cli)")
			loComandoSeleccionar.AppendLine(" LEFT JOIN  Formas_Pagos ON (Facturas.Cod_For	=   Formas_Pagos.Cod_For)")
			loComandoSeleccionar.AppendLine(" LEFT JOIN  Vendedores ON (Facturas.Cod_Ven	=   Vendedores.Cod_Ven)")
            loComandoSeleccionar.AppendLine(" LEFT JOIN  Articulos ON (Articulos.Cod_Art	=   Renglones_Facturas.Cod_Art)")
            loComandoSeleccionar.AppendLine(" LEFT JOIN  Transportes ON (Facturas.Cod_Tra	=   Transportes.Cod_Tra)")
            loComandoSeleccionar.AppendLine(" LEFT JOIN  Cotizaciones ON (Cotizaciones.Documento = Renglones_Facturas.Doc_Ori)")
            loComandoSeleccionar.AppendLine(" LEFT JOIN  Pedidos On (Pedidos.Documento = Renglones_Facturas.Doc_Ori)")
            loComandoSeleccionar.AppendLine(" LEFT JOIN  Entregas ON (Entregas.Documento = Renglones_Facturas.Doc_Ori)")
            loComandoSeleccionar.AppendLine(" WHERE   " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY Renglones_Facturas.Renglon ASC")
   

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            Dim lcCadenaComentario As String = ""
            Dim lcComentario As String

            lcCadenaComentario = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcComentario = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("Comentario_Origen")

                If lcComentario = "" Then
                    Continue For
                End If
				lcCadenaComentario = lcCadenaComentario.Trim & lcComentario & ",  "
				
            Next lnNumeroFila

            lcCadenaComentario = lcCadenaComentario & ")"
            lcCadenaComentario = lcCadenaComentario.Replace("(,", "(")
            lcCadenaComentario = lcCadenaComentario.Replace(".", ",")
            lcCadenaComentario = lcCadenaComentario.Replace(",)", ")")
            lcCadenaComentario = lcCadenaComentario.Replace("(,", "(")
            lcCadenaComentario = lcCadenaComentario.Replace(", ", "")
            lcCadenaComentario = lcCadenaComentario.Replace(",", ", ")

   			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

			
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_OEnvio_Stratos", laDatosReporte)

			CType(loObjetoReporte.ReportDefinition.ReportObjects("Text24"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcCadenaComentario.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFacturas_OEnvio_Stratos.ReportSource = loObjetoReporte

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
'-----------------------------------------------------------------------------------------------'
' Fin del codigo
'-----------------------------------------------------------------------------------------------'
' CMS: 12/02/10: Codigo inicial
'-----------------------------------------------------------------------------------------------'
' CMS: 18/02/10: Se ajusto el formato para mostrar el logo de Stratos
'-----------------------------------------------------------------------------------------------'
' MAT: 14/03/11: Inclusión de las Observaciones en el Shipping de todos los documentos Origenes
'-----------------------------------------------------------------------------------------------'
' MAT: 05/09/11: Nuevos Ajustes al formato según requerimientos
'-----------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Ajuste de la vista de Diseño
'-------------------------------------------------------------------------------------------'	
' MAT: 15/09/11: Eliminación del Pie de Página de eFactory según Requerimientos
'-------------------------------------------------------------------------------------------'
' RJG: 04/06/12: Se agregaron los datos genericos del cliente (tomados de la factura actual). 
'-------------------------------------------------------------------------------------------'
