'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFicha_DatosActivos"
'-------------------------------------------------------------------------------------------'
Partial Class fFicha_DatosActivos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            Dim lcTipo As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo & "Activos_Fijos")

            loComandoSeleccionar.AppendLine("SELECT	activos_fijos.Cod_Act,")
            loComandoSeleccionar.AppendLine("		activos_fijos.nom_act,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Marca,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Modelo,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Seriales,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Origen,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Situacion,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Cod_Gru,")
            loComandoSeleccionar.AppendLine("		Grupos_Activos.nom_gru,")
            loComandoSeleccionar.AppendLine("		activos_fijos.cod_ubi,")
            loComandoSeleccionar.AppendLine("		Ubicaciones_Activos.nom_ubi,")
            loComandoSeleccionar.AppendLine("		activos_fijos.cod_tip,")
            loComandoSeleccionar.AppendLine("		Tipos_Activos.nom_tip,")
            loComandoSeleccionar.AppendLine("		activos_fijos.mon_adq,")
            loComandoSeleccionar.AppendLine("		activos_fijos.fec_adq,")
            loComandoSeleccionar.AppendLine("		activos_fijos.fec_inc,")
            loComandoSeleccionar.AppendLine("		activos_fijos.cod_cen,")
            loComandoSeleccionar.AppendLine("		Centros_Costos.nom_cen,")
            loComandoSeleccionar.AppendLine("		activos_fijos.cod_mon,")
            loComandoSeleccionar.AppendLine("		Monedas.nom_mon,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Metodo,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Mon_Sal,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Mon_Ini,")
            loComandoSeleccionar.AppendLine("		activos_fijos.mon_acu,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Mon_ult,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Fin_dep,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Vid_año,")
            loComandoSeleccionar.AppendLine("		activos_fijos.vid_mes,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Cue_Act, ")
            loComandoSeleccionar.AppendLine("		Activo.Nom_cue AS Activo,")
            loComandoSeleccionar.AppendLine("		activos_fijos.cue_gas,")
            loComandoSeleccionar.AppendLine("		Depreciacion.Nom_Cue As Depreciacion,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Cue_Acu,")
            loComandoSeleccionar.AppendLine("		Acumulado.Nom_Cue AS Acumulado,")
            loComandoSeleccionar.AppendLine("		activos_fijos.Cue_Adi,")
            loComandoSeleccionar.AppendLine("		Adicional.Nom_Cue AS Adicional,")
            loComandoSeleccionar.AppendLine("		activos_fijos.comentario")
            loComandoSeleccionar.AppendLine("FROM activos_fijos")
            loComandoSeleccionar.AppendLine("	JOIN Grupos_Activos ON Grupos_Activos.cod_gru = activos_fijos.Cod_Gru")
            loComandoSeleccionar.AppendLine("	JOIN Ubicaciones_Activos ON Ubicaciones_Activos.cod_ubi = activos_fijos.cod_ubi")
            loComandoSeleccionar.AppendLine("	JOIN Tipos_Activos ON Tipos_Activos.cod_tip = activos_fijos.cod_tip ")
            loComandoSeleccionar.AppendLine("	JOIN Centros_Costos ON Centros_Costos.cod_cen = activos_fijos.cod_cen")
            loComandoSeleccionar.AppendLine("	JOIN Monedas ON Monedas.cod_mon = activos_fijos.cod_mon	")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Contables AS Activo ON Activo.Cod_cue = activos_fijos.Cue_Act")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Contables	AS Depreciacion ON Depreciacion.Cod_Cue = activos_fijos.cue_gas")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Contables	AS Acumulado On Acumulado.Cod_Cue = activos_fijos.Cue_Acu")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Contables AS Adicional ON Adicional.Cod_Cue = activos_fijos.Cue_Adi")
            loComandoSeleccionar.AppendLine("WHERE  " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFicha_DatosActivos", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFicha_DatosActivos.ReportSource = loObjetoReporte

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
' EAG: 19/08/2015 : Codigo Inicial
'-------------------------------------------------------------------------------------------'