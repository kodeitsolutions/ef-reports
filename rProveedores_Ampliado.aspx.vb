Imports System.Data
Partial Class rProveedores_Ampliado
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Proveedores.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Por_Isl, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Por_Ret, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Tip_Pro, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Zon, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Con, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Cod_Per, ")
            loComandoSeleccionar.AppendLine("		Personas.Nom_Per, ")
            loComandoSeleccionar.AppendLine("		Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("		Clases_Proveedores.Nom_Cla, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("FROM	proveedores ")
            loComandoSeleccionar.AppendLine("JOIN	Tipos_Proveedores ON Proveedores.Cod_Tip = Tipos_Proveedores.Cod_Tip ")
            loComandoSeleccionar.AppendLine("JOIN	Zonas ON Proveedores.Cod_Zon = Zonas.Cod_Zon ")
            loComandoSeleccionar.AppendLine("JOIN 	Clases_Proveedores ON Proveedores.Cod_Cla = Clases_Proveedores.Cod_Cla ")
            loComandoSeleccionar.AppendLine("JOIN	Vendedores ON Proveedores.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("JOIN   Conceptos ON Conceptos.Cod_Con = Proveedores.Cod_Con ")
            loComandoSeleccionar.AppendLine("JOIN   Personas ON Personas.Cod_Per = Proveedores.Cod_Per ")
            'loComandoSeleccionar.AppendLine("JOIN   Clases_Proveedores ON Clases_Proveedores.Cod_Cla = Proveedores.Cod_Cla ")
            loComandoSeleccionar.AppendLine("WHERE	Proveedores.Cod_Pro between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" AND 	Proveedores.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" AND 	Tipos_Proveedores.Cod_Tip between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" AND 	Zonas.Cod_Zon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" AND 	Clases_Proveedores.Cod_Cla between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" AND 	Vendedores.Cod_Ven between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" AND 	" & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY proveedores." & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProveedores_Ampliado", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProveedores_Ampliado.ReportSource = loObjetoReporte


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
' MVP:  10/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' YJP:  21/04/09: Estandarizacion de codigos y correccion de campo estatus
'-------------------------------------------------------------------------------------------'
' PMV:  17/06/15: Creacion del Reporte Listado Ampliado de Proveedores
'-------------------------------------------------------------------------------------------'