'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fTraslados_Almacenes"
'-------------------------------------------------------------------------------------------'
Partial Class fTraslados_Almacenes
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSelect As New StringBuilder()

            loComandoSelect.Append("WITH curTemporal As ( ")
            loComandoSelect.Append("SELECT  Traslados.Documento, ")
            loComandoSelect.Append("        Traslados.Status, ")
            loComandoSelect.Append("        Traslados.Fec_Ini, ")
            loComandoSelect.Append("        Traslados.Fec_Fin, ")
            loComandoSelect.Append("        Traslados.Alm_Ori, ")
            loComandoSelect.Append("        Traslados.Alm_Des, ")
            loComandoSelect.Append("        Traslados.Cod_Tra, ")
            loComandoSelect.Append("        Traslados.Comentario, ")
            loComandoSelect.Append("        Renglones_Traslados.Renglon, ")
            loComandoSelect.Append("        Renglones_Traslados.Cod_Art, ")
            loComandoSelect.Append("        Renglones_Traslados.Cod_Uni, ")
            loComandoSelect.Append("        Renglones_Traslados.Can_Art1, ")
            loComandoSelect.Append("        Renglones_Traslados.Notas, ")
            loComandoSelect.Append("        Transportes.Nom_Tra, ")
            loComandoSelect.Append("        Almacenes.Nom_Alm As Nom_Ori ")
            loComandoSelect.Append("FROM 	Traslados, ")
            loComandoSelect.Append("        Renglones_Traslados, ")
            loComandoSelect.Append("        Transportes, ")
            loComandoSelect.Append("        Almacenes ")
            loComandoSelect.Append("WHERE	Traslados.Documento =   Renglones_Traslados.Documento AND ")
            loComandoSelect.Append("        Traslados.Cod_Tra   =   Transportes.Cod_Tra AND ")
            loComandoSelect.Append("        Traslados.Alm_Ori   =   Almacenes.Cod_Alm AND ")
            loComandoSelect.Append("        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSelect.Append(")")

            loComandoSelect.Append("SELECT  curTemporal.Documento, ")
            loComandoSelect.Append("        curTemporal.Status, ")
            loComandoSelect.Append("        curTemporal.Fec_Ini, ")
            loComandoSelect.Append("        curTemporal.Fec_Fin, ")
            loComandoSelect.Append("        curTemporal.Alm_Ori, ")
            loComandoSelect.Append("        curTemporal.Alm_Des, ")
            loComandoSelect.Append("        curTemporal.Cod_Tra, ")
            loComandoSelect.Append("        curTemporal.Comentario, ")
            loComandoSelect.Append("        curTemporal.Renglon, ")
            loComandoSelect.Append("        curTemporal.Cod_Art, ")
            loComandoSelect.Append("        curTemporal.Cod_Uni, ")
            loComandoSelect.Append("        curTemporal.Can_Art1, ")
            loComandoSelect.Append("        curTemporal.Notas, ")
            loComandoSelect.Append("        curTemporal.Nom_Tra, ")
            loComandoSelect.Append("        curTemporal.Nom_Ori, ")
            loComandoSelect.Append("        Almacenes.Nom_Alm As Nom_Des ")
            loComandoSelect.Append("FROM 	curTemporal, ")
            loComandoSelect.Append("        Almacenes ")
            loComandoSelect.Append("WHERE	curTemporal.Alm_Des   =   Almacenes.Cod_Alm")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSelect.ToString, "curReportes")

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
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fTraslados_Almacenes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfTraslados_Almacenes.ReportSource = loObjetoReporte

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
' GMO: 10/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 08/11/08: Ajustes al Select
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'