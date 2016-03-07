Imports System.Data
Partial Class rArticulos_SinVentas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
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


            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Act1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Ult1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Pro1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Cli1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Ant1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Sec, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine("FROM		Articulos, ")
            loComandoSeleccionar.AppendLine(" 			Departamentos, ")
            loComandoSeleccionar.AppendLine(" 			Secciones, ")
            loComandoSeleccionar.AppendLine(" 			Marcas, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Articulos, ")
            loComandoSeleccionar.AppendLine(" 			Clases_Articulos ")
            loComandoSeleccionar.AppendLine("WHERE		Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("	AND		Articulos.Cod_Sec = Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("	AND		Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("	AND		Articulos.Cod_Tip = Tipos_Articulos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("	AND		Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine("	AND		Secciones.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("	AND		Articulos.Cod_Art between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("	AND		" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("	AND		Articulos.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("	AND		Departamentos.Cod_Dep BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("	AND		" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("	AND		Secciones.Cod_Sec BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("	AND		" & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("	AND		Marcas.Cod_Mar BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("	AND		" & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("	AND		Tipos_Articulos.Cod_Tip BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("	AND		" & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("	AND		Clases_Articulos.Cod_Cla BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("	AND		" & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("	AND		Articulos.Cod_Art Not IN (SELECT Cod_Art FROM Renglones_Facturas GROUP BY Cod_Art) ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY	 Articulos.Cod_Art, Articulos.Nom_Art")

           

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_SinVentas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_SinVentas.ReportSource = loObjetoReporte


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
' JFP:  10/12/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP:  25/04/09 : Agreagar combo estatus y estandarizacion de codigo
'-------------------------------------------------------------------------------------------'
' AAP:  02/07/09 : Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  05/08/09 : Se Agrego la siguiente union: Secciones.Cod_Dep = Departamentos.Cod_Dep
'-------------------------------------------------------------------------------------------'
