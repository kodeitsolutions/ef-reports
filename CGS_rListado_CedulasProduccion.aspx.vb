'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rListado_CedulasProduccion"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rListado_CedulasProduccion
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try
            lcComandoSeleccionar.AppendLine("DECLARE @lcDoc_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcDoc_Hasta AS VARCHAR(10) = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Formulas.Documento						    AS Cedula,")
            lcComandoSeleccionar.AppendLine("		Formulas.Cod_Art						    AS Art_Result,")
            lcComandoSeleccionar.AppendLine("		Formulas.Nom_Art						    AS Nom_Art_Result,")
            lcComandoSeleccionar.AppendLine("		Formulas.Cod_Uni						    AS Uni_Result,")
            lcComandoSeleccionar.AppendLine("		Formulas.Fec_Ini						    AS Fec_Cedula,")
            lcComandoSeleccionar.AppendLine("		Formulas.Cod_Eta						    AS Etapa,")
            lcComandoSeleccionar.AppendLine("		Etapas.Nom_Eta							    AS Nom_Etapa,")
            lcComandoSeleccionar.AppendLine("		Renglones_Formulas.Renglon				    AS Renglon,")
            lcComandoSeleccionar.AppendLine("		Renglones_Formulas.Cod_Art				    AS Art_Base,")
            lcComandoSeleccionar.AppendLine("		Articulos.Nom_Art						    AS Nom_Art_Base,")
            lcComandoSeleccionar.AppendLine("		Renglones_Formulas.Cod_Uni				    AS Uni_Base,")
            lcComandoSeleccionar.AppendLine("       @lcDoc_Desde								AS Doc_Desde,")
            lcComandoSeleccionar.AppendLine("       CASE WHEN @lcDoc_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("       	 THEN @lcDoc_Hasta ELSE '' END			AS Doc_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Art_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Art_Hasta")
            lcComandoSeleccionar.AppendLine("FROM Formulas	")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Formulas ON Formulas.Documento = Renglones_Formulas.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Formulas.Cod_Art = Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine("	JOIN Etapas ON Formulas.Cod_Eta = Etapas.Cod_Eta ")
            lcComandoSeleccionar.AppendLine("WHERE Formulas.Documento BETWEEN @lcDoc_Desde AND  @lcDoc_Hasta")
            lcComandoSeleccionar.AppendLine("       AND Formulas.Status IN ( " & lcParametro1Desde & ")")
            lcComandoSeleccionar.AppendLine("       AND Formulas.Documento IN (SELECT Documento FROM Renglones_Formulas WHERE Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta)")
            lcComandoSeleccionar.AppendLine("ORDER BY Formulas.Documento")


            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(lcComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rListado_CedulasProduccion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rListado_CedulasProduccion.ReportSource = loObjetoReporte

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
' JJD: 14/10/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 20/04/09: se agregaron las condiciones: Ordenes_Compras.Fec_Ini, Proveedores.nom_pro y Ordenes_Compras.status
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' CMS: 18/06/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS: 22/07/09: Filtro BackOrder, lo conllevo al anexo del campo Can_Pen1,
'                 verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS:  13/08/09: Se Agrego la restricción Renglones_Pedidos.Can_Pen1 <> 0 cuando el filtro 
'                   BackOrder = BackOrder
'-------------------------------------------------------------------------------------------'
' CMS: 19/03/10: se agrego el filtro cod_art
'-------------------------------------------------------------------------------------------'