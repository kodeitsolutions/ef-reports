'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fMGanancia_Departamentos"
'-------------------------------------------------------------------------------------------'
Partial Class fMGanancia_Departamentos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine(" 		    Departamentos.Nom_Dep, ")
            loComandoSeleccionar.AppendLine(" 		    COUNT(Articulos.Cod_Dep)    AS  Can_Art, ")
            loComandoSeleccionar.AppendLine(" 		    AVG(Articulos.Por_Gan1)     AS  Por_Pre1, ")
            loComandoSeleccionar.AppendLine(" 		    AVG(Articulos.Por_Gan2)     AS  Por_Pre2, ")
            loComandoSeleccionar.AppendLine(" 		    AVG(Articulos.Por_Gan3)     AS  Por_Pre3, ")
            loComandoSeleccionar.AppendLine(" 		    AVG(Articulos.Por_Gan4)     AS  Por_Pre4, ")
            loComandoSeleccionar.AppendLine(" 		    AVG(Articulos.Por_Gan5)     AS  Por_Pre5 ")
            loComandoSeleccionar.AppendLine(" FROM      Articulos, ")
            loComandoSeleccionar.AppendLine("           Departamentos ")
            loComandoSeleccionar.AppendLine(" WHERE     Articulos.Cod_Dep   =   Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Articulos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine(" 		    Departamentos.Nom_Dep ")

            'me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            'Me.mCargarLogoEmpresa(loTablaLogo, "LogoEmpresa")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fMGanancia_Departamentos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfMGanancia_Departamentos.ReportSource = loObjetoReporte

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
' JJD: 19/04/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 28/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
