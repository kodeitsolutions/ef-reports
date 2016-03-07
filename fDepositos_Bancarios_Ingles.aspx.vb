'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDepositos_Bancarios_Ingles"
'-------------------------------------------------------------------------------------------'
Partial Class fDepositos_Bancarios_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Depositos.Documento, ")
            loComandoSeleccionar.AppendLine("           Depositos.Num_Dep, ")
            loComandoSeleccionar.AppendLine("           Depositos.Status, ")
            loComandoSeleccionar.AppendLine("           Depositos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Depositos.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Depositos.Cod_Mon AS Moneda, ")
            loComandoSeleccionar.AppendLine("           Depositos.Tasa, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Nom_Cue, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Efe, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Che, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Tar, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Otr, ")
            loComandoSeleccionar.AppendLine("           Depositos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Depositos.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Depositos.Comentario AS Notas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Cod_Tar, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Por_Com, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Mon_Com, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Por_Ret, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Mon_Ret, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           (Renglones_Depositos.Mon_Com + Renglones_Depositos.Mon_Imp) AS ComImp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Tipo, ")
            loComandoSeleccionar.AppendLine("           ISNULL((SELECT Nom_Ban FROM Bancos WHERE Cod_Ban = Renglones_Depositos.Cod_Ban), '')AS Banco_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Referencia, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Mon_Net AS  Mon_Ren, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING((CASE WHEN Renglones_Depositos.Tipo = 'Efectivo' Then ' Cash ' Else 'Source : ' + CAST(UPPER(RTRIM(Renglones_Depositos.Tip_Ori)) AS VARCHAR(20)) + '   Number : ' + CAST(RTRIM(Renglones_Depositos.Doc_Ori) AS VARCHAR(20)) End),1,50) AS Origen, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos.Comentario AS Comentario_Renglones_Depositos ")
            loComandoSeleccionar.AppendLine(" FROM      Depositos, ")
            loComandoSeleccionar.AppendLine("           Renglones_Depositos, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias, ")
            loComandoSeleccionar.AppendLine("           Conceptos, ")
            loComandoSeleccionar.AppendLine("           Bancos ")
            loComandoSeleccionar.AppendLine(" WHERE     Depositos.Documento     =   Renglones_Depositos.Documento ")
            loComandoSeleccionar.AppendLine("           And Depositos.Cod_Cue   =   Cuentas_Bancarias.Cod_Cue ")
            loComandoSeleccionar.AppendLine("           And Bancos.Cod_Ban      =   Cuentas_Bancarias.Cod_Ban ")
            loComandoSeleccionar.AppendLine("           And Depositos.Cod_Con   =   Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine("           And " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDepositos_Bancarios_Ingles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfDepositos_Bancarios_Ingles.ReportSource = loObjetoReporte

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
' MAT: 23/03/11: Mejora de la vista de Diseño
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Ajuste de la vista de Diseño
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Eliminación del Pie de Página de eFactory según Requerimientos
'-------------------------------------------------------------------------------------------'
