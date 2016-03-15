'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rLista_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rLista_Articulos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            
           
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Articulos.Cod_Art		AS Codigo_Articulo, ")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art		AS Descripcion, ")
            loComandoSeleccionar.AppendLine("		Departamentos.Nom_Dep	AS Departamento, ")
            loComandoSeleccionar.AppendLine("		Secciones.Nom_Sec		AS Seccion,	")
            loComandoSeleccionar.AppendLine("		Articulos.usa_lot		AS Lote, ")
            loComandoSeleccionar.AppendLine("		Articulos.cod_uni1		AS Unidad_Medida, ")
            loComandoSeleccionar.AppendLine("		Articulos.Atributo_A, ")
            loComandoSeleccionar.AppendLine("       CASE WHEN Articulos.Atributo_A = ' ' THEN NULL ")
            loComandoSeleccionar.AppendLine("       ELSE     ")
            loComandoSeleccionar.AppendLine("           convert(numeric(18,2), replace(Articulos.Atributo_A,',','.') ) * 100 ")
            loComandoSeleccionar.AppendLine("       END AS Porcentaje_Desperdicio, ")
            loComandoSeleccionar.AppendLine("		Articulos.Generico		AS Art_Generico, ")
            loComandoSeleccionar.AppendLine("		Articulos.Tipo			AS Tipo_Uso ")
            loComandoSeleccionar.AppendLine("FROM	Articulos ")
            loComandoSeleccionar.AppendLine("	JOIN Departamentos ")
            loComandoSeleccionar.AppendLine("	 ON Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("	JOIN Secciones ")
            loComandoSeleccionar.AppendLine("	 ON Articulos.Cod_Sec = Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("	 AND Secciones.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("WHERE	Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		And Departamentos.Cod_Dep BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		And Secciones.Cod_Sec BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY Departamentos.Nom_Dep, Secciones.Nom_Sec ")
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rLista_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rLista_Articulos.ReportSource = loObjetoReporte

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
' MJP: 16/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GS:  14/03/16: Cambio a Listado de Artículos.
'-------------------------------------------------------------------------------------------'

