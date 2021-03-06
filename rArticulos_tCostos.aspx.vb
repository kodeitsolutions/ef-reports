﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_tCostos"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_tCostos
    Inherits vis2formularios.frmReporte

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
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcComandoSeleccionar As New StringBuilder()


            lcComandoSeleccionar.AppendLine(" SELECT		Articulos.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Nom_Art, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Cos_Ult1, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Cos_Pro1, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Cos_Cli1, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Cos_Ant1, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Cod_Dep, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Cod_Sec, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Cod_Mar, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("				Articulos.Cod_Cla ")
            lcComandoSeleccionar.AppendLine(" FROM			Articulos, ")
            lcComandoSeleccionar.AppendLine("				Departamentos, ")
            lcComandoSeleccionar.AppendLine("				Secciones, ")
            lcComandoSeleccionar.AppendLine("				Marcas, ")
            lcComandoSeleccionar.AppendLine("				Tipos_Articulos, ")
            lcComandoSeleccionar.AppendLine("				Clases_Articulos ")
            lcComandoSeleccionar.AppendLine(" WHERE		Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Sec = Secciones.Cod_Sec ")
            lcComandoSeleccionar.AppendLine(" 				AND Departamentos.Cod_Dep = Secciones.Cod_Dep ")
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Mar = Marcas.Cod_Mar ")
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Tip = Tipos_Articulos.Cod_Tip ")
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Art between " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.status IN (" & lcParametro1Desde & ")")
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Dep between " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Sec between " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Mar between " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Tip between " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Cla between " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine("      			And Articulos.Cod_Ubi between " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine(" 				And " & lcParametro7Hasta)
            'lcComandoSeleccionar.AppendLine(" ORDER BY		Articulos.Cod_Art, Articulos.Nom_Art")
            lcComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_tCostos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_tCostos.ReportSource = loObjetoReporte


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
' MVP:  10/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  11/07/08 : Adición de loObjetoReporte para eliminar los archivos temp en Uranus
'-------------------------------------------------------------------------------------------'
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR:  24/03/09: Estandarizacion	de código y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS: 05/05/09: Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Verificacion de registros y se agregaro el filtro_ Ubicación
'-------------------------------------------------------------------------------------------'
' CMS:  10/05/10: Se agrego la restriccion Departamentos.Cod_Dep = Secciones.Cod_Dep
'-------------------------------------------------------------------------------------------'