'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_fOProduccion_OTrabajo"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fOProduccion_OTrabajo
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Proyectos.Cod_Pro				            AS OProduccion,")
            loComandoSeleccionar.AppendLine("		Proyectos.Nom_Pro				            AS Nom_OProduccion,")
            loComandoSeleccionar.AppendLine("		Proyectos.Responsable			            AS Responsable,")
            loComandoSeleccionar.AppendLine("		Proyectos.Pro_Pri				            AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro				            AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Ordenes_Trabajo.Documento		            AS OTrabajo,")
            loComandoSeleccionar.AppendLine("		Renglones_OTrabajo.Cod_Reg		            AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		Renglones_OTrabajo.Nom_Reg		            AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		Renglones_OTrabajo.Can_Art		            AS Cantidad_Producida,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot		            AS Lote_Salida,")
            loComandoSeleccionar.AppendLine("		COALESCE(Mediciones.Rechazo,0)	            AS Porc_Desperdicio,")
            loComandoSeleccionar.AppendLine("		COALESCE((Mediciones.Rechazo*Renglones_OTrabajo.Can_Art)/100,0)	AS Cant_Desperdicio,")
            loComandoSeleccionar.AppendLine("		COALESCE(Consumo_Produccion.Documento,'')						AS Consumo_Produccion ")
            loComandoSeleccionar.AppendLine("FROM Proyectos")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Proyectos.Pro_Pri")
            loComandoSeleccionar.AppendLine("	JOIN Encabezados AS Ordenes_Trabajo ON Ordenes_Trabajo.Proyecto = Proyectos.Cod_Pro	")
            loComandoSeleccionar.AppendLine("		AND Ordenes_Trabajo.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Renglones AS Renglones_OTrabajo ON Renglones_OTrabajo.Documento = Ordenes_Trabajo.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_OTrabajo.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Ordenes_Trabajo.Documento = Operaciones_Lotes.Num_Doc")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Adicional = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Ren_Ori = Renglones_OTrabajo.Renglon")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Ordenes_Trabajo.Documento ")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Encabezados'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Encabezados AS Consumo_Produccion ON Consumo_Produccion.Proyecto = Proyectos.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND Consumo_Produccion.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fOProduccion_OTrabajo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fOProduccion_OTrabajo.ReportSource = loObjetoReporte

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
