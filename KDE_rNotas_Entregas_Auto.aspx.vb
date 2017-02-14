'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "KDE_rNotas_Entregas_Auto"
'-------------------------------------------------------------------------------------------'
Partial Class KDE_rNotas_Entregas_Auto

    Inherits vis2formularios.frmReporteAutomatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try

        Dim loComandoSeleccionar As New StringBuilder()


        loComandoSeleccionar.AppendLine("SELECT Entregas.Documento 			    AS Documento,")
        loComandoSeleccionar.AppendLine("		Entregas.Fec_Ini				AS Fecha,")
        loComandoSeleccionar.AppendLine("		Entregas.Comentario			    AS Comentario,")
        loComandoSeleccionar.AppendLine("		Renglones_Entregas.Cod_Art		AS Cod_Art,")
        loComandoSeleccionar.AppendLine("		Renglones_Entregas.Cod_Uni		AS Cod_Uni,")
        loComandoSeleccionar.AppendLine("		Renglones_Entregas.Can_Art1		AS Cantidad,")
        loComandoSeleccionar.AppendLine("		Renglones_Entregas.Precio1		AS Precio,")
        loComandoSeleccionar.AppendLine("		Renglones_Entregas.Mon_Net		AS Neto,")
        loComandoSeleccionar.AppendLine("		Renglones_Entregas.Notas		AS Notas,")
        loComandoSeleccionar.AppendLine("		Renglones_Entregas.Comentario   AS Com_Renglon,")
        loComandoSeleccionar.AppendLine("		Articulos.Nom_Art				AS Nom_Art,")
        loComandoSeleccionar.AppendLine("		Clientes.Cod_Cli				AS Cod_Cli,")
        loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli				AS Nom_Cli")
        loComandoSeleccionar.AppendLine("FROM Entregas")
        loComandoSeleccionar.AppendLine("	JOIN Renglones_Entregas ON Entregas.Documento = Renglones_Entregas.Documento")
        loComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Entregas.Cod_Art = Articulos.Cod_Art")
        loComandoSeleccionar.AppendLine("	JOIN Clientes ON Entregas.Cod_Cli = Clientes.Cod_Cli")
        loComandoSeleccionar.AppendLine("WHERE Entregas.Status = 'Confirmado'")


        'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

        Dim loServicios As New cusDatos.goDatos
        Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

        loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("KDE_rNotas_Entregas_Auto", laDatosReporte)

        Me.mTraducirReporte(loObjetoReporte)
        Me.mFormatearCamposReporte(loObjetoReporte)
        Me.crvKDE_rNotas_Entregas_Auto.ReportSource = loObjetoReporte

        'Catch loExcepcion As Exception

        '    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
        '              "No se pudo Completar el Proceso: " & loExcepcion.Message, _
        '               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
        '               "auto", _
        '               "auto")

        'End Try

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