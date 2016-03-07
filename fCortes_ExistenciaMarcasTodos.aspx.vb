'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCortes_ExistenciaMarcasTodos"
'-------------------------------------------------------------------------------------------'
Partial Class fCortes_ExistenciaMarcasTodos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine(" 		    Marcas.Nom_Mar, ")
            loComandoSeleccionar.AppendLine(" 		    SUM(Renglones_Almacenes.Exi_Act1)   AS  Exi_Act1, ")
            loComandoSeleccionar.AppendLine(" 		    SUM(Renglones_Almacenes.Exi_Ped1)   AS  Exi_Ped1, ")
            loComandoSeleccionar.AppendLine(" 		    SUM(Renglones_Almacenes.Exi_Act1-Renglones_Almacenes.Exi_Ped1)   AS  Exi_Dis1, ")
            loComandoSeleccionar.AppendLine(" 		    SUM(Renglones_Almacenes.Exi_Por1)   AS  Exi_Por1, ")
            loComandoSeleccionar.AppendLine(" 		    SUM(Renglones_Almacenes.Exi_Cot1)   AS  Exi_Cot1, ")
            loComandoSeleccionar.AppendLine(" 		    SUM(Renglones_Almacenes.Exi_Des1)   AS  Exi_Des1 ")
            loComandoSeleccionar.AppendLine(" FROM      Articulos, ")
            loComandoSeleccionar.AppendLine("           Marcas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Almacenes ")
            loComandoSeleccionar.AppendLine(" WHERE     Articulos.Cod_Art       =   Renglones_Almacenes.Cod_Art ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar   =   Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine(" 		    Marcas.Nom_Mar ")

            'me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCortes_ExistenciaMarcasTodos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCortes_ExistenciaMarcasTodos.ReportSource = loObjetoReporte

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