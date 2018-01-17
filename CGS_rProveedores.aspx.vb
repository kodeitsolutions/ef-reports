Imports System.Data
Partial Class CGS_rProveedores
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta AS VARCHAR(10) = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcTipo_Desde AS VARCHAR(30) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcTipo_Hasta AS VARCHAR(30) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcClase_Desde AS VARCHAR(30) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcClase_Hasta AS VARCHAR(30) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcMostrar AS VARCHAR(10) = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Proveedores.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Por_Isl, ")
            loComandoSeleccionar.AppendLine("		Proveedores.Por_Ret, ")
            loComandoSeleccionar.AppendLine("       CASE WHEN RTRIM(Proveedores.Atributo_A) NOT LIKE '%0.0%'")
            loComandoSeleccionar.AppendLine("            THEN 0")
            loComandoSeleccionar.AppendLine("            ELSE CONVERT(NUMERIC(18,2), Proveedores.Atributo_A) * 100")
            loComandoSeleccionar.AppendLine("       END       AS Por_Pat,")
            loComandoSeleccionar.AppendLine("		Tipos_Proveedores.Nom_Tip, ")
            loComandoSeleccionar.AppendLine("		Clases_Proveedores.Nom_Cla, ")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodPro_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro FROM Proveedores  WHERE Cod_Pro = @lcCodPro_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Pro_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodPro_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro  FROM Proveedores  WHERE Cod_Pro = @lcCodPro_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Pro_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcClase_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Cla FROM Clases_Proveedores WHERE Cod_Cla = @lcClase_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE ''")
            loComandoSeleccionar.AppendLine("		END												        AS Clase_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcClase_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Cla FROM Clases_Proveedores WHERE Cod_Cla = @lcClase_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE ''")
            loComandoSeleccionar.AppendLine("		END												        AS Clase_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcTipo_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Tip FROM Tipos_Proveedores WHERE Cod_Tip = @lcTipo_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Tipo_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcTipo_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Tip FROM Tipos_Proveedores WHERE Cod_Tip = @lcTipo_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Tipo_Hasta,")
            loComandoSeleccionar.AppendLine("       CASE WHEN @lcMostrar = 'TODOS' THEN 'Todos' ELSE 'Solo con compras' END AS Mostrar")
            loComandoSeleccionar.AppendLine("FROM Proveedores ")
            loComandoSeleccionar.AppendLine("   JOIN Tipos_Proveedores ON Proveedores.Cod_Tip = Tipos_Proveedores.Cod_Tip ")
            loComandoSeleccionar.AppendLine("   JOIN Clases_Proveedores ON Proveedores.Cod_Cla = Clases_Proveedores.Cod_Cla ")
            loComandoSeleccionar.AppendLine("WHERE Proveedores.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine("   AND Proveedores.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("   AND	Tipos_Proveedores.Cod_Tip BETWEEN @lcTipo_Desde AND @lcTipo_Hasta")
            loComandoSeleccionar.AppendLine("   AND	Clases_Proveedores.Cod_Cla BETWEEN @lcClase_Desde AND @lcClase_Hasta")

            If lcParametro4Desde = "COMPRAS" Then
                loComandoSeleccionar.AppendLine(" AND Proveedores.Cod_Pro IN (SELECT Cod_Pro FROM Compras)")
            End If
            loComandoSeleccionar.AppendLine("ORDER BY proveedores." & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rProveedores.ReportSource = loObjetoReporte


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