Imports System.Data
Partial Class fLibres_Inventarios_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Libres_Inventarios.Documento, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Mon_Imp          As  Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Status, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Comentario, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Cod_Mon          As  Moneda, ")
            loComandoSeleccionar.AppendLine("           Libres_Inventarios.Tasa, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_uni1 AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Mon_Net      As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Por_Imp1     As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Cod_Imp      As  Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios.Mon_Imp1     As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Libres_Inventarios, ")
            loComandoSeleccionar.AppendLine("           Renglones_LInventarios, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Libres_Inventarios.Documento    =   Renglones_LInventarios.Documento ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art           =   Renglones_LInventarios.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY  Libres_Inventarios.Documento, Renglones_LInventarios.Renglon ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fLibres_Inventarios_Ingles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfLibres_Inventarios_Ingles.ReportSource = loObjetoReporte

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
' Douglas Cortez: 07/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT :  24/08/11 : Ajuste del Select, Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
' MAT :  15/09/11: Eliminación del Pie de Página del eFactory según requerimiento
'-------------------------------------------------------------------------------------------'
