'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rFormulas_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rFormulas_Renglones
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1),goServicios.enuTipoRedondeoFecha.KN_InicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuTipoRedondeoFecha.KN_FinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            loComandoSeleccionar.AppendLine(" WITH		curTemporal	AS	( ")
            loComandoSeleccionar.AppendLine(" SELECT	Formulas.Documento, ")
            loComandoSeleccionar.AppendLine("			Formulas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Formulas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("			CASE")
            loComandoSeleccionar.AppendLine("			        WHEN Formulas.Status = 'A' THEN 'Activo'")
            loComandoSeleccionar.AppendLine("			        WHEN Formulas.Status = 'I' THEN 'Inactivo'")
            loComandoSeleccionar.AppendLine("			        WHEN Formulas.Status = 'S' THEN 'Suspendido'")
            loComandoSeleccionar.AppendLine("			END AS Status, ")
            loComandoSeleccionar.AppendLine("			Formulas.Comentario, ")
            loComandoSeleccionar.AppendLine("			Renglones_Formulas.Renglon, ")
            loComandoSeleccionar.AppendLine("			Renglones_Formulas.Cod_Art	AS	Cod_For, ")
            loComandoSeleccionar.AppendLine("			Renglones_Formulas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("			Renglones_Formulas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("			Renglones_Formulas.Precio1, ")
            loComandoSeleccionar.AppendLine("			Renglones_Formulas.Cos_Pro1, ")
            loComandoSeleccionar.AppendLine("			(Renglones_Formulas.Can_Art1 * Renglones_Formulas.Precio1) as Total_Precio,")
            loComandoSeleccionar.AppendLine("			(Renglones_Formulas.Can_Art1 * Renglones_Formulas.Cos_Pro1) as Total_Costo")
            loComandoSeleccionar.AppendLine(" FROM		Formulas, ")
            loComandoSeleccionar.AppendLine("			Renglones_Formulas, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Formulas.Documento					=	Renglones_Formulas.Documento ")
            loComandoSeleccionar.AppendLine("			AND Renglones_Formulas.Cod_Art		=	Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Formulas.Documento				Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Formulas.Fec_Ini				Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Formulas.Status					IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Formulas.origen='')")


            loComandoSeleccionar.AppendLine(" SELECT	curTemporal.*, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art AS Nom_For ")
            loComandoSeleccionar.AppendLine(" FROM		curTemporal, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		curTemporal.Cod_For	=	Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento & ", curTemporal.Renglon ASC")

            
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFormulas_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrFormulas_Renglones.ReportSource = loObjetoReporte

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
' JJD: 04/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 23/04/09: Agregar combo estatus y estandarizacion
'-------------------------------------------------------------------------------------------'
' CMS:  12/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'