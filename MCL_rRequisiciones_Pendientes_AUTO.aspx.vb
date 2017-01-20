'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MCL_rRequisiciones_Pendientes_AUTO"
'-------------------------------------------------------------------------------------------'
Partial Class MCL_rRequisiciones_Pendientes_AUTO

    Inherits vis2formularios.frmReporteAutomatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try

        Dim loComandoSeleccionar As New StringBuilder()

        loComandoSeleccionar.AppendLine("DECLARE @ldAnterior AS DATE = DATEADD(dd,DATEDIFF(dd,0,GETDATE()),-1)")
        loComandoSeleccionar.AppendLine("DECLARE @ldFecha AS DATE = CASE DATEPART(dw, @ldAnterior)  WHEN 7 THEN DATEADD(d,-2,@ldAnterior)")
        loComandoSeleccionar.AppendLine("                               WHEN 7 THEN DATEADD(d,-2,@ldAnterior)")
        loComandoSeleccionar.AppendLine("                               WHEN 6 THEN DATEADD(d,-1, @ldAnterior) ")
        loComandoSeleccionar.AppendLine("                           ELSE @ldAnterior END")
        loComandoSeleccionar.AppendLine("")
        loComandoSeleccionar.AppendLine("SELECT Requisiciones.Documento 			AS Documento,")
        loComandoSeleccionar.AppendLine("		Requisiciones.Fec_Ini				AS Fecha,")
        loComandoSeleccionar.AppendLine("		Requisiciones.Comentario			AS Comentario,")
        loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Cod_Art		AS Cod_Art,")
        loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Cod_Alm		AS Cod_Alm,")
        loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Cod_Uni		AS Cod_Uni,")
        loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Can_Art1	AS Cantidad,")
        loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Notas		AS Notas,")
        loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Comentario	AS Com_Renglon,")
        loComandoSeleccionar.AppendLine("		Articulos.Nom_Art					AS Nom_Art,")
        loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro					AS Nom_Pro")
        loComandoSeleccionar.AppendLine("FROM Requisiciones")
        loComandoSeleccionar.AppendLine("	JOIN Renglones_Requisiciones ON Requisiciones.Documento = Renglones_Requisiciones.Documento")
        loComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Requisiciones.Cod_Art = Articulos.Cod_Art")
        loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Requisiciones.Cod_Pro = Proveedores.Cod_Pro")
        loComandoSeleccionar.AppendLine("WHERE Requisiciones.Fec_Ini = @ldFecha")

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

        loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("MCL_rRequisiciones_Pendientes_AUTO", laDatosReporte)

        Me.mTraducirReporte(loObjetoReporte)
        Me.mFormatearCamposReporte(loObjetoReporte)
        Me.crvMCL_rRequisiciones_Pendientes_AUTO.ReportSource = loObjetoReporte

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