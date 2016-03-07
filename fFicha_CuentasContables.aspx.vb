'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFicha_CuentasContables"
'-------------------------------------------------------------------------------------------'
Partial Class fFicha_CuentasContables
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            Dim lcTipo As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo & "Cuentas_Contables")

            loComandoSeleccionar.AppendLine("SELECT	Cuentas_Contables.Cod_Cue,")
            loComandoSeleccionar.AppendLine("		(CASE Cuentas_Contables.status")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO'")
            loComandoSeleccionar.AppendLine("		END) AS status,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Nom_Cue,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.nom_adi,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Tip_cue,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.movimiento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Categoria,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.comentario,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Centros,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Cod_Cen,")
            loComandoSeleccionar.AppendLine("		Centros_Costos.Nom_Cen,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Gasto,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Cod_gas,")
            loComandoSeleccionar.AppendLine("		COALESCE(Cuentas_Gastos.Nom_Gas,' ') AS Nom_Gas,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Auxiliar,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.cod_aux,")
            loComandoSeleccionar.AppendLine("		COALESCE(Auxiliares.Nom_Aux,'') AS Nom_Aux,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Tipo,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Clase,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Usa_Pre,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Usa_Doc,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.efectivo,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Usa_Cla,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Mon_Adi,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.fec_doc,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Cod_alt,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Nom_Alt,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Cla_Alt,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Gru_Alt,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Seg_Alt,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Fam_Alt,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Mon_Alt1,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Mon_Alt2,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Mon_Alt3,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Mon_Alt4,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Mon_Alt5,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Atributo_A,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Atributo_B,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Atributo_C,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Atributo_D,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Atributo_E,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Atributo_F")
            loComandoSeleccionar.AppendLine("From Cuentas_Contables")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Centros_Costos ON Centros_Costos.cod_cen = Cuentas_Contables.cod_cen")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Gastos ON Cuentas_Gastos.Cod_Gas = Cuentas_Contables.Cod_Gas")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Auxiliares ON Auxiliares.cod_aux = Cuentas_Contables.Cod_aux")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFicha_CuentasContables", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFicha_CuentasContables.ReportSource = loObjetoReporte

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
' MAT: 22/06/2011 : Codigo Inicial
'-------------------------------------------------------------------------------------------'