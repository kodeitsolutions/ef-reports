'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fBeneficiarios"
'-------------------------------------------------------------------------------------------'
Partial Class fBeneficiarios
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Proveedores.cod_pro, ")
            loComandoSeleccionar.AppendLine("		(CASE Proveedores.status")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO'")
            loComandoSeleccionar.AppendLine("		END) AS status,")
            loComandoSeleccionar.AppendLine("		Proveedores.nom_pro,")
            loComandoSeleccionar.AppendLine("		(CASE Proveedores.sexo")
            loComandoSeleccionar.AppendLine("			WHEN 'M' THEN 'MASCULINO'")
            loComandoSeleccionar.AppendLine("			WHEN 'F' THEN 'FEMENINO'")
            loComandoSeleccionar.AppendLine("			WHEN 'X' THEN 'NO APLICA'")
            loComandoSeleccionar.AppendLine("			ELSE Proveedores.sexo")
            loComandoSeleccionar.AppendLine("		END) AS sexo,")
            loComandoSeleccionar.AppendLine("		proveedores.cod_tip,")
            loComandoSeleccionar.AppendLine("		Tipos_Proveedores.nom_tip,")
            loComandoSeleccionar.AppendLine("		Proveedores.cod_zon,")
            loComandoSeleccionar.AppendLine("		Zonas.nom_zon,")
            loComandoSeleccionar.AppendLine("		Proveedores.rif,")
            loComandoSeleccionar.AppendLine("		Proveedores.nit,")
            loComandoSeleccionar.AppendLine("		Proveedores.Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Proveedores.fec_fin,")
            loComandoSeleccionar.AppendLine("		Proveedores.Contacto,")
            loComandoSeleccionar.AppendLine("		Proveedores.correo,")
            loComandoSeleccionar.AppendLine("		Proveedores.web,")
            loComandoSeleccionar.AppendLine("		Proveedores.telefonos,")
            loComandoSeleccionar.AppendLine("		Proveedores.fax,")
            loComandoSeleccionar.AppendLine("		Proveedores.movil,")
            loComandoSeleccionar.AppendLine("		Proveedores.cod_pos")
            loComandoSeleccionar.AppendLine("FROM Proveedores")
            loComandoSeleccionar.AppendLine("	JOIN Tipos_Proveedores ON Tipos_Proveedores.cod_tip = proveedores.cod_tip")
            loComandoSeleccionar.AppendLine("	JOIN Zonas ON Zonas.Cod_zon = Proveedores.cod_zon")
            loComandoSeleccionar.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fBeneficiarios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfBeneficiarios.ReportSource = loObjetoReporte

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
' EAG: 02/09/2015 : Codigo inicial
'-------------------------------------------------------------------------------------------'
