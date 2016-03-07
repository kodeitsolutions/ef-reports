'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fLibres_Importaciones"
'-------------------------------------------------------------------------------------------'
Partial Class fLibres_Importaciones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Libres_Importaciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Libres_Importaciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Libres_Importaciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Libres_Importaciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Libres_Importaciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Libres_Importaciones.Status, ")
            loComandoSeleccionar.AppendLine("           Libres_Importaciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Libres_Importaciones.Cod_Mon          As  Moneda, ")
            loComandoSeleccionar.AppendLine("           Libres_Importaciones.Tasa, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Mon_Net      As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Por_Imp1     As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Cod_Imp      As  Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones.Mon_Imp1     As  Impuesto ")
            loComandoSeleccionar.AppendLine("           Proveedores.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Libres_Importaciones.Control ")
            loComandoSeleccionar.AppendLine(" FROM      Libres_Importaciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_LImportaciones, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Libres_Importaciones.Documento    =   Renglones_LImportaciones.Documento ")
            loComandoSeleccionar.AppendLine("           And Proveedores.cod_pro = Libres_Importaciones.cod_pro ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art           =   Renglones_LImportaciones.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY  Libres_Importaciones.Documento, Renglones_LImportaciones.Renglon ")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fLibres_Importaciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfLibres_Importaciones.ReportSource = loObjetoReporte

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
' CMS: 17/01/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
