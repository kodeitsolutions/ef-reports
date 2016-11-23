'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_fOProduccion_Desperdicio"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fOProduccion_Desperdicio
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Proyectos.Cod_Pro				AS OProduccion,")
            loComandoSeleccionar.AppendLine("		Proyectos.Nom_Pro				AS Nom_OProduccion,")
            loComandoSeleccionar.AppendLine("		Proyectos.Responsable			AS Responsable,")
            loComandoSeleccionar.AppendLine("		Proyectos.Pro_Pri				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro				AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Mediciones.Rechazo				AS Desp_Estandar,")
            loComandoSeleccionar.AppendLine("		Lotes_OTrabajo.Cod_Art			AS Art_Producido,")
            loComandoSeleccionar.AppendLine("		Art_Producido.Nom_Art			AS Nom_Art_Producido,")
            loComandoSeleccionar.AppendLine("		Lotes_OTrabajo.Cod_Lot			AS Lotes_Producido,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_OTrabajo.Can_Art)	AS Total_Producido,")
            loComandoSeleccionar.AppendLine("		Lotes_Consumos.Cod_Art			AS Art_Consumido,")
            loComandoSeleccionar.AppendLine("		Art_Consumido.Nom_Art			AS Nom_Art_Consumido,")
            loComandoSeleccionar.AppendLine("		Lotes_Consumos.Cod_Lot			AS Lotes_Consumido,")
            loComandoSeleccionar.AppendLine("		Renglones_Consumo.Can_Art		AS Total_Consumido,")
            loComandoSeleccionar.AppendLine("		(Renglones_Consumo.Can_Art - SUM(Renglones_OTrabajo.Can_Art)) / SUM(Renglones_OTrabajo.Can_Art) AS Desp_Real")
            loComandoSeleccionar.AppendLine("FROM Proyectos")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Proyectos.Pro_Pri")
            loComandoSeleccionar.AppendLine("	JOIN Encabezados AS Consumo_Produccion  ON Consumo_Produccion.Proyecto = Proyectos.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND Consumo_Produccion.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("	JOIN Renglones AS Renglones_Consumo ON Renglones_Consumo.Documento = Consumo_Produccion.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_Consumo.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes AS Lotes_Consumos ON Consumo_Produccion.Documento = Lotes_Consumos.Num_Doc")
            loComandoSeleccionar.AppendLine("		AND Lotes_Consumos.Tip_Doc = 'Encabezados'")
            loComandoSeleccionar.AppendLine("		AND Lotes_Consumos.Tip_Ope = 'Salida'")
            loComandoSeleccionar.AppendLine("		AND Lotes_Consumos.Ren_Ori = Renglones_Consumo.Renglon")
            loComandoSeleccionar.AppendLine("	JOIN Articulos AS Art_Consumido ON Lotes_Consumos.Cod_Art = Art_Consumido.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes AS Lotes_Traslados ON Lotes_Consumos.Cod_Lot = Lotes_Traslados.Cod_Lot")
            loComandoSeleccionar.AppendLine("		AND Lotes_Traslados.Tip_Doc = 'Traslados'")
            loComandoSeleccionar.AppendLine("		AND Lotes_Traslados.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("	JOIN Traslados ON Lotes_Traslados.Num_Doc = Traslados.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Mediciones ON Mediciones.Cod_Reg = Lotes_Traslados.Num_Doc")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Traslados'")
            loComandoSeleccionar.AppendLine("	JOIN Encabezados AS Ordenes_Trabajo  ON Ordenes_Trabajo.Proyecto = Proyectos.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND Ordenes_Trabajo.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Renglones AS Renglones_OTrabajo ON Renglones_OTrabajo.Documento = Ordenes_Trabajo.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_OTrabajo.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes AS Lotes_OTrabajo ON Lotes_OTrabajo.Num_Doc = Ordenes_Trabajo.Documento")
            loComandoSeleccionar.AppendLine("		AND Lotes_OTrabajo.Tip_Doc = 'Encabezados'")
            loComandoSeleccionar.AppendLine("		AND Lotes_OTrabajo.Ren_Ori = Renglones_OTrabajo.Renglon")
            loComandoSeleccionar.AppendLine("		AND Lotes_OTrabajo.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("	JOIN Articulos AS Art_Producido ON Lotes_OTrabajo.Cod_Art = Art_Producido.Cod_Art")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("GROUP BY Proyectos.Cod_Pro, Proyectos.Nom_Pro, Proyectos.Responsable, Proyectos.Pro_Pri, Proveedores.Nom_Pro,Lotes_OTrabajo.Cod_Lot,	")
            loComandoSeleccionar.AppendLine("		 Consumo_Produccion.Documento, Traslados.Documento, Mediciones.Rechazo, Renglones_Consumo.Can_Art, Lotes_Consumos.Cod_Lot,")
            loComandoSeleccionar.AppendLine("		 Lotes_OTrabajo.Cod_Art, Art_Producido.Nom_Art, Lotes_Consumos.Cod_Art, Art_Consumido.Nom_Art")


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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fOProduccion_Desperdicio", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fOProduccion_Desperdicio.ReportSource = loObjetoReporte

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
' CMS: 07/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 26/02/10: Verificacion de la extraccion de los datos correctos de 
'-------------------------------------------------------------------------------------------'
