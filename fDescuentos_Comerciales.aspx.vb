'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDescuentos_Comerciales"
'-------------------------------------------------------------------------------------------'
Partial Class fDescuentos_Comerciales
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            
            Dim loComandoSeleccionar As New StringBuilder()
            
			loComandoSeleccionar.AppendLine("SELECT		Clientes.Origen,")
			loComandoSeleccionar.AppendLine("			DA.Cod_Cli,")
			loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli, ")
			loComandoSeleccionar.AppendLine("			DA.Cod_Can, ")
			loComandoSeleccionar.AppendLine("			DA.Cod_Ope, ")
			loComandoSeleccionar.AppendLine("			DA.Cod_Cla,		DA.Cod_Dep,		DA.Cod_Mar,		DA.Cod_Seg,")
			loComandoSeleccionar.AppendLine("			DA.Por_Des1,	DA.Por_Des2,	DA.Por_Des3, 	DA.Por_Des4,")
			loComandoSeleccionar.AppendLine("			DA.Por_Des5,	DA.Por_Des6,	DA.Por_Des7, 	DA.Por_Des8,")
			loComandoSeleccionar.AppendLine("			DA.Por_Des9,	DA.Por_Des10,	DA.Por_Des11, 	DA.Por_Des12,")
			loComandoSeleccionar.AppendLine("			DA.Por_Des13,	DA.Por_Des14,	DA.Por_Des15, 	DA.Por_Des16,")
			loComandoSeleccionar.AppendLine("			DA.Por_Des17,	DA.Por_Des18,	DA.Por_Des19, 	DA.Por_Des20,")
			loComandoSeleccionar.AppendLine("			CASE DA.Status")
			loComandoSeleccionar.AppendLine("				WHEN 'A'	THEN 'Activo'")
			loComandoSeleccionar.AppendLine("				WHEN 'I'	THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine("				WHEN 'S'	THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine("				ELSE			 'Desconocido'")
			loComandoSeleccionar.AppendLine("			END AS Status")
			loComandoSeleccionar.AppendLine("FROM		Clientes ")
			loComandoSeleccionar.AppendLine("	JOIN	Descuentos_Articulos AS DA")
			loComandoSeleccionar.AppendLine("		ON	DA.Cod_Cli		= Clientes.Cod_Cli")
			loComandoSeleccionar.AppendLine("		AND	DA.Origen		= 'Clientes'")
			loComandoSeleccionar.AppendLine("		AND	DA.Adicional	= 'Comercial'")
			loComandoSeleccionar.AppendLine("		AND	DA.Clase		= 'Comercial'")
			loComandoSeleccionar.AppendLine("WHERE	  " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loComandoSeleccionar.AppendLine("ORDER BY	DA.Cod_Cla, DA.Cod_Dep, DA.Cod_Mar, DA.Cod_Seg")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
                
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDescuentos_Comerciales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfDescuentos_Comerciales.ReportSource = loObjetoReporte

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
' RJG: 02/04/12: Código Inicial.															'
'-------------------------------------------------------------------------------------------'
' RJG: 05/05/12: Se agregaron los campos Adicional y Clase al filtro de Descuentos. Se		'
'				 agregó el encabezado con Logo.												'
'-------------------------------------------------------------------------------------------'
