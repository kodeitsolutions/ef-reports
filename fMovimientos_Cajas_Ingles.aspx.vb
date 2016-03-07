'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fMovimientos_Cajas_Ingles"
'-------------------------------------------------------------------------------------------'
Partial Class fMovimientos_Cajas_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cajas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tasa, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cajas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Cajas.Nom_Caj, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban ")
            loComandoSeleccionar.AppendLine(" FROM      Movimientos_Cajas ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Cajas     ON Movimientos_Cajas.Cod_Caj = Cajas.Cod_Caj ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Bancos    ON Movimientos_Cajas.Cod_Ban = Bancos.Cod_Ban ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Conceptos ON Movimientos_Cajas.Cod_Con = Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            'loComandoSeleccionar.AppendLine(" INTO      #tmpMovimientosCajas01 ")

            'loComandoSeleccionar.AppendLine(" SELECT    #tmpMovimientosCajas01.*, ")
            'loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban          AS  Banco ")
            'loComandoSeleccionar.AppendLine(" INTO      #tmpMovimientosCajas02 ")
            'loComandoSeleccionar.AppendLine(" FROM      #tmpMovimientosCajas01, ")
            'loComandoSeleccionar.AppendLine("           Bancos ")
            'loComandoSeleccionar.AppendLine(" WHERE     #tmpMovimientosCajas01.Cod_Ban   =   Bancos.Cod_Ban ")

            'loComandoSeleccionar.AppendLine(" SELECT    #tmpMovimientosCajas02.*, ")
            'loComandoSeleccionar.AppendLine("           Tarjetas.Nom_Tar        AS  Tarjeta ")
            'loComandoSeleccionar.AppendLine(" FROM      #tmpMovimientosCajas02, ")
            'loComandoSeleccionar.AppendLine("           Tarjetas ")
            'loComandoSeleccionar.AppendLine(" WHERE     #tmpMovimientosCajas02.Cod_Tar   =   Tarjetas.Cod_Tar ")

            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fMovimientos_Cajas_Ingles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfMovimientos_Cajas_Ingles.ReportSource = loObjetoReporte

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
' Douglas Cortez: 03/05/2010: Programación inicial
'-------------------------------------------------------------------------------------------'
' MAT :  24/08/11 : Ajuste del Select, Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
' MAT :  15/09/11: Eliminación del Pie de Página del eFactory según requerimiento
'-------------------------------------------------------------------------------------------'
