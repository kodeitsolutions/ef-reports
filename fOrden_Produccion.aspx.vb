'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fOrden_Produccion"
'-------------------------------------------------------------------------------------------'
Partial Class fOrden_Produccion
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Ordenes_Produccion.Documento    AS  Doc_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Fec_Ini      AS  Fec_Ini_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Fec_Fin      AS  Fec_Fin_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Cod_Art      AS  Cod_Art_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Nom_Art      AS  Nom_Art_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Status       AS  Status_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Modelo       AS  Modelo_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Cod_Uni      AS  Cod_Uni_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Can_Art1     AS  Can_Art1_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Comentario   AS  Comentario, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Cos_Pro1     AS  Cos_Pro1_Pro, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_oProduccion.Documento AS  Doc_oPro, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_oProduccion.Renglon   AS  Renglon, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_oProduccion.Cod_Art   AS  Cod_Art_oPro, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_oProduccion.Nom_Art   AS  Nom_Art_oPro, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_oProduccion.Modelo    AS	Modelo_oPro, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_oProduccion.Can_Art1  AS  Can_Art1_oPro, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_oProduccion.Alm_Ori   AS  Alm_Ori, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_oProduccion.Cod_Uni   AS  Cod_Uni_oPro, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_oProduccion.Cos_Pro1  AS  Cos_Pro1, ")
            loComandoSeleccionar.AppendLine(" 			Ordenes_Produccion.Cod_Cli      AS  Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli                AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif                    AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit                    AS  Nit, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis                AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos              AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax                    AS  Fax ")
            loComandoSeleccionar.AppendLine(" FROM		Ordenes_Produccion ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Renglones_oProduccion ON Renglones_oProduccion.Documento = Ordenes_Produccion.Documento ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Clientes              ON Clientes.Cod_Cli = Ordenes_Produccion.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrden_Produccion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrden_Produccion.ReportSource = loObjetoReporte

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
' CMS: 07/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 26/02/10: Verificacion de la extraccion de los datos correctos de 
'-------------------------------------------------------------------------------------------'
