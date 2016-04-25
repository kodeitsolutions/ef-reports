'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_fActaServicio_RI"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fActaServicio_RI

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Rif,")
            loComandoSeleccionar.AppendLine("		Requisiciones.Documento,")
            loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Cod_Art,")
            loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Comentario,")
            loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Notas,")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art")
            loComandoSeleccionar.AppendLine("FROM	Requisiciones")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Requisiciones ")
            loComandoSeleccionar.AppendLine("		ON Requisiciones.Documento = Renglones_Requisiciones.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ")
            loComandoSeleccionar.AppendLine("		ON Renglones_Requisiciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores")
            loComandoSeleccionar.AppendLine("		ON Requisiciones.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE Renglones_Requisiciones.Cod_Uni = 'SRV'")
            loComandoSeleccionar.AppendLine("   AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------'
            ' Carga la imagen del logo en cusReportes                                                   '
            '-------------------------------------------------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------'
            ' Verificando si el select (tabla nº0) trae registros                                       '
            '-------------------------------------------------------------------------------------------'

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                          vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fActaServicio_RI", laDatosReporte)

            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_fActaServicio_RI.ReportSource = loObjetoReporte

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
' JJD: 27/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/09: Ajuste al formato del impuesto IVA
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 11/06/10: Proveedor Genarico
'-------------------------------------------------------------------------------------------'
' JJD: 11/03/5: Ajustes al formato para el cliente CEGASA
'-------------------------------------------------------------------------------------------'