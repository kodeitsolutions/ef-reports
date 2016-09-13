'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rUsuarios_Empresas"
'-------------------------------------------------------------------------------------------'
Partial Class rUsuarios_Empresas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Empresas_Usuarios.Cod_Usu								AS Cod_Usu,")
            loComandoSeleccionar.AppendLine("			Empresas_Usuarios.Cod_Emp								AS Cod_Emp,")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Empresas.Status = 'A'")
            loComandoSeleccionar.AppendLine("					THEN 'Activo'")
            loComandoSeleccionar.AppendLine("					ELSE 'Inactivo'")
            loComandoSeleccionar.AppendLine("			END) AS Status_Empresa,")
            loComandoSeleccionar.AppendLine("           Empresas.Nom_Emp										AS Nom_Emp,")
            loComandoSeleccionar.AppendLine("           Usuarios.Nom_Usu										AS Nom_Usu,")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Usuarios.Status = 'A' ")
            loComandoSeleccionar.AppendLine("					THEN 'Activo' ")
            loComandoSeleccionar.AppendLine("					ELSE 'Inactivo' ")
            loComandoSeleccionar.AppendLine("			END)													AS Status_Usuario,")
            loComandoSeleccionar.AppendLine("			(CASE	Empresas.Sistema ")
            loComandoSeleccionar.AppendLine("					WHEN 'Factory_Administrativo'	THEN 'Administrativo' ")
            loComandoSeleccionar.AppendLine("					WHEN 'Factory_Contabilidad'		THEN 'Contabilidad' ")
            loComandoSeleccionar.AppendLine("					WHEN 'Factory_Nomina'			THEN 'Nómina' ")
            loComandoSeleccionar.AppendLine("					ELSE Empresas.Sistema ")
            loComandoSeleccionar.AppendLine("			END)													AS Sistema")
            loComandoSeleccionar.AppendLine("FROM		Empresas_Usuarios")
            loComandoSeleccionar.AppendLine("	JOIN	Empresas ")
            loComandoSeleccionar.AppendLine("		ON	Empresas.Cod_Emp = Empresas_Usuarios.Cod_Emp")
            loComandoSeleccionar.AppendLine("		AND	Empresas.Sistema = Empresas_Usuarios.Sistema")
            loComandoSeleccionar.AppendLine("	JOIN	Usuarios")
            loComandoSeleccionar.AppendLine("		ON	Usuarios.Cod_Usu = Empresas_Usuarios.Cod_Usu")
            loComandoSeleccionar.AppendLine("WHERE		Usuarios.Cod_Usu BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Usuarios.Status   IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Usuarios.Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))
            loComandoSeleccionar.AppendLine("ORDER BY	Empresas_Usuarios.Cod_Usu, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos
            
            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rUsuarios_Empresas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrUsuarios_Empresas.ReportSource = loObjetoReporte

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
' GCR: 01/04/09: Codigo Inicial
'-------------------------------------------------------------------------------------------'
' CMS: 14/07/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'