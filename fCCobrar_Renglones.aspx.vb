'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCCobrar_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class fCCobrar_Renglones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Cobrar.Cod_Cli, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cuentas_Cobrar.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Rif = '') THEN Clientes.Rif ELSE Cuentas_Cobrar.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Cuentas_Cobrar.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Cuentas_Cobrar.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Cobrar.Telefonos = '') THEN Clientes.Telefonos ELSE Cuentas_Cobrar.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Por_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           (Cuentas_Cobrar.Mon_Otr1 + Cuentas_Cobrar.Mon_Otr2 + Cuentas_Cobrar.Mon_Otr3)   AS  Mon_Otr, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")            
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Renglon                AS  Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Art                AS  Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Precio1                AS  Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Can_Art                AS  Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Uni                AS  Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Mon_Net                AS  Mon_Net_Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Comentario             AS  Comentario_Renglon ")
            
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Renglones_Documentos ON Cuentas_Cobrar.Documento = Renglones_Documentos.Documento ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Clientes ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Vendedores ON Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Cuentas_Cobrar.Cod_For ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Articulos ON Articulos.Cod_Art = Renglones_Documentos.Cod_Art ")             
            loComandoSeleccionar.AppendLine(" WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
            
			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCCobrar_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCCobrar_Renglones.ReportSource = loObjetoReporte

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
' CMS: 04/03/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplico el metodo carga de imagen 
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se Ajusto para tomar el cliente generico
'-------------------------------------------------------------------------------------------'