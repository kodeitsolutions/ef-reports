'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_fRequisiciones_Internas_Auto"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fRequisiciones_Internas_Auto

    Inherits vis2formularios.frmReporteAutomatico

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try

        Dim loComandoSeleccionar As New StringBuilder()


        loComandoSeleccionar.AppendLine("SELECT	Requisiciones.Cod_Pro, ")
        loComandoSeleccionar.AppendLine("        (CASE WHEN (Proveedores.Generico = 0 AND Cast(Requisiciones.Nom_Pro As Varchar)= '') THEN Proveedores.Nom_Pro ELSE ")
        loComandoSeleccionar.AppendLine("            (CASE WHEN (Cast(Requisiciones.Nom_Pro As Varchar) = '') THEN Proveedores.Nom_Pro ELSE Requisiciones.Nom_Pro END) END) AS  Nom_Pro, ")
        loComandoSeleccionar.AppendLine("        (CASE WHEN (Proveedores.Generico = 0 AND Cast(Requisiciones.Nom_Pro As Varchar) = '') THEN Proveedores.Rif ELSE ")
        loComandoSeleccionar.AppendLine("            (CASE WHEN (Requisiciones.Rif = '') THEN Proveedores.Rif ELSE Requisiciones.Rif END) END) AS  Rif, ")
        loComandoSeleccionar.AppendLine("        Proveedores.Nit, ")
        loComandoSeleccionar.AppendLine("        (CASE WHEN (Proveedores.Generico = 0 AND Cast(Requisiciones.Nom_Pro As Varchar) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
        loComandoSeleccionar.AppendLine("            (CASE WHEN (SUBSTRING(Requisiciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Requisiciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
        loComandoSeleccionar.AppendLine("        (CASE WHEN (Proveedores.Generico = 0 AND Cast(Requisiciones.Nom_Pro As Varchar) = '') THEN Proveedores.Telefonos ELSE ")
        loComandoSeleccionar.AppendLine("            (CASE WHEN (Requisiciones.Telefonos = '') THEN Proveedores.Telefonos ELSE Requisiciones.Telefonos END) END) AS  Telefonos, ")
        loComandoSeleccionar.AppendLine("        Requisiciones.Documento,")
        loComandoSeleccionar.AppendLine("        Requisiciones.Fec_Ini,      ")
        loComandoSeleccionar.AppendLine("        Requisiciones.Comentario, ")
        loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Cod_Art, ")
        loComandoSeleccionar.AppendLine("        Renglones_Requisiciones.Comentario		AS Com_Ren,")
        loComandoSeleccionar.AppendLine("		Renglones_Requisiciones.Notas			AS Notas, ")
        loComandoSeleccionar.AppendLine("        Articulos.Nom_Art						AS Nom_Art,  ")
        loComandoSeleccionar.AppendLine("        Renglones_Requisiciones.Can_Art1, ")
        loComandoSeleccionar.AppendLine("        Renglones_Requisiciones.Cod_Uni ")
        loComandoSeleccionar.AppendLine("FROM      Requisiciones ")
        loComandoSeleccionar.AppendLine("JOIN Renglones_Requisiciones	ON Requisiciones.Documento  =   Renglones_Requisiciones.Documento ")
        loComandoSeleccionar.AppendLine("JOIN Proveedores				ON Requisiciones.Cod_Pro    =   Proveedores.Cod_Pro ")
        loComandoSeleccionar.AppendLine("JOIN Formas_Pagos				ON Requisiciones.Cod_For    =   Formas_Pagos.Cod_For ")
        loComandoSeleccionar.AppendLine("JOIN Vendedores				    ON Requisiciones.Cod_Ven    =   Vendedores.Cod_Ven ")
        loComandoSeleccionar.AppendLine("JOIN Articulos					ON Articulos.Cod_Art		=   Renglones_Requisiciones.Cod_Art ")
        loComandoSeleccionar.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)


        'Me.mEscribirConsulta(loComandoSeleccionar.ToString())



        Dim loServicios As New cusDatos.goDatos
        Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

        '-------------------------------------------------------------------------------------------'
        ' Carga la imagen del logo en cusReportes                                                   '
        '-------------------------------------------------------------------------------------------'
        Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

        '-------------------------------------------------------------------------------------------'
        ' Verificando si el select (tabla nº0) trae registros                                       '
        '-------------------------------------------------------------------------------------------'

        If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                      "No se Encontraron Registros para los Parámetros Especificados. ", _
                                      vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                       "350px", _
                                       "200px")
        End If

        loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fRequisiciones_Internas_Auto", laDatosReporte)

        Me.mTraducirReporte(loObjetoReporte)
        Me.mFormatearCamposReporte(loObjetoReporte)
        Me.crvCGS_fRequisiciones_Internas_Auto.ReportSource = loObjetoReporte

        'Catch loExcepcion As Exception

        '    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
        '              "No se pudo Completar el Proceso: " & loExcepcion.Message, _
        '               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
        '               "auto", _
        '               "auto")

        'End Try

    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                                 "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                                  vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                                  "auto", _
                                  "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 27/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/09: Ajuste al formato del impuesto IVA
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 11/06/10: Proveedor Genarico
'-------------------------------------------------------------------------------------------'
' JJD: 11/03/5: Ajustes al formato para el cliente CEGASA
'-------------------------------------------------------------------------------------------'