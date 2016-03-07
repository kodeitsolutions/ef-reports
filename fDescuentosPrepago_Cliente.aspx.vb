'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDescuentosPrepago_Cliente"
'-------------------------------------------------------------------------------------------'
Partial Class fDescuentosPrepago_Cliente
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            
            Dim loComandoSeleccionar As New StringBuilder()
            
			loComandoSeleccionar.AppendLine("SELECT		Clientes.Origen				AS Origen,")
			loComandoSeleccionar.AppendLine("			DA.Cod_Cli					AS Cod_Cli,")
			loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli			AS Nom_Cli,")
			loComandoSeleccionar.AppendLine("			DA.Cod_Cla					AS Cod_Cla,")
			loComandoSeleccionar.AppendLine("			CA.Nom_Cla					AS Nom_Cla,")
			loComandoSeleccionar.AppendLine("			DA.Por_Des1					AS Por_Des1,")
			loComandoSeleccionar.AppendLine("			CASE DA.Status")
			loComandoSeleccionar.AppendLine("				WHEN 'A'	THEN 'Activo'")
			loComandoSeleccionar.AppendLine("				WHEN 'I'	THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine("				WHEN 'S'	THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine("				ELSE			 'Desconocido'")
			loComandoSeleccionar.AppendLine("			END							AS Status")
			loComandoSeleccionar.AppendLine("FROM		Clientes ")
			loComandoSeleccionar.AppendLine("	JOIN	Descuentos_Articulos AS DA")
			loComandoSeleccionar.AppendLine("		ON	DA.Cod_Cli		= Clientes.Cod_Cli")
			loComandoSeleccionar.AppendLine("		AND	DA.Origen		= 'Clientes'")
			loComandoSeleccionar.AppendLine("		AND	DA.Adicional	= 'Prepago'")
			loComandoSeleccionar.AppendLine("		AND	DA.Clase		= 'Clase'")
			loComandoSeleccionar.AppendLine("	JOIN	Clases_Articulos AS CA ")
			loComandoSeleccionar.AppendLine("		ON	 CA.Cod_Cla = DA.Cod_Cla")
			loComandoSeleccionar.AppendLine("WHERE	  " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("ORDER BY	DA.Cod_Cla")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
                
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDescuentosPrepago_Cliente", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfDescuentosPrepago_Cliente.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 14/05/12: Código Inicial.															'
'-------------------------------------------------------------------------------------------'
