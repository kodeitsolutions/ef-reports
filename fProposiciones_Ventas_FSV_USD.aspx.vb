Imports System.Data
Partial Class fProposiciones_Ventas_FSV_USD
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Proposiciones.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Proposiciones.Nom_Pro>'' THEN Proposiciones.Nom_Pro ELSE Prospectos.Nom_pro END) AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Proposiciones.Rif>'' THEN Proposiciones.Rif ELSE Prospectos.Rif END) AS Rif,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Proposiciones.Nit>'' THEN Proposiciones.Nit ELSE Prospectos.Nit END) AS Nit,")
            loComandoSeleccionar.AppendLine("           REPLACE(REPLACE((CASE WHEN Proposiciones.Dir_Fis>'' THEN Proposiciones.Dir_Fis ELSE Prospectos.Dir_Fis END), CHAR(13), ' '), CHAR(10), ' ') AS Dir_Fis,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Proposiciones.Telefonos>'' THEN Proposiciones.Telefonos ELSE Prospectos.Telefonos END) AS Telefonos,")
            loComandoSeleccionar.AppendLine("           Prospectos.Fax, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Nom_Pro       As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Tasa, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art + Renglones_Proposiciones.Comentario    As  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones.Mon_Imp1 As  Impuesto, ")
            loComandoSeleccionar.AppendLine("           CAST('' AS VARCHAR(MAX)) As  Monto_Letras ")
            loComandoSeleccionar.AppendLine(" FROM      Proposiciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Proposiciones, ")
            loComandoSeleccionar.AppendLine("           Prospectos, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Proposiciones.Documento  =   Renglones_Proposiciones.Documento AND ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_Pro    =   Prospectos.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Proposiciones.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art        =   Renglones_Proposiciones.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
            'Genera la columna con el monto en letras
            If  (laDatosReporte.Tables.Count > 0)           AndAlso _
                (laDatosReporte.Tables(0).Rows.Count > 0)   Then 

                Dim lcMontoLetras As String 
                lcMontoLetras = goServicios.mConvertirMontoLetras(CDec(laDatosReporte.Tables(0).Rows(0).Item("Mon_Net")))

                For Each loFila As DataRow In laDatosReporte.Tables(0).Rows

                    loFila("Monto_Letras") = lcMontoLetras
                    
                Next

            End If

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fProposiciones_Ventas_FSV_USD", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfProposiciones_Ventas_FSV_USD.ReportSource = loObjetoReporte

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
' EAG: 14/08/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RJG: 21/08/15: Ajuste de interfaz (tamaño y posiciones de etiquetas). Se agrego el monto  '
'                en letras, y marca de agua.                                                '
'-------------------------------------------------------------------------------------------'
