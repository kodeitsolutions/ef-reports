Imports System.Data
Partial Class fCuentas_Cobrar_GPV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("SELECT	  cuentas_cobrar.Cod_Cli,   ")
            loComandoSeleccionar.AppendLine("          Tipos_Documentos.nom_tip   AS Titulo,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Cod_tip,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Factura,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Control,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Doc_Ori,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Tip_Ori,   ")
            loComandoSeleccionar.AppendLine("          Clientes.Nom_Cli,   ")
            loComandoSeleccionar.AppendLine("          Clientes.Rif,   ")
            loComandoSeleccionar.AppendLine("          Clientes.Nit,   ")
            loComandoSeleccionar.AppendLine("          Clientes.Dir_Fis,   ")
            loComandoSeleccionar.AppendLine("          Clientes.Telefonos,   ")
            loComandoSeleccionar.AppendLine("          Clientes.Fax,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Nom_Cli        As  Nom_Gen,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Rif            As  Rif_Gen,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Nit            As  Nit_Gen,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Dir_Fis        As  Dir_Gen,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Telefonos      As  Tel_Gen,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Documento,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Fec_Ini,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Fec_Fin,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Mon_Bru,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Mon_Imp1,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Mon_Net,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Mon_Net,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Por_Des	   AS Por_Des1,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Mon_Des       As  Mon_Des,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Por_Rec	   AS Por_Rec1,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Mon_Rec       AS  Mon_Rec,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Cod_For,   ")
            loComandoSeleccionar.AppendLine("          Formas_Pagos.Nom_For,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Cod_Ven,   ")
            loComandoSeleccionar.AppendLine("          cuentas_cobrar.Comentario,   ")
            loComandoSeleccionar.AppendLine("          Vendedores.Nom_Ven,   ")
            loComandoSeleccionar.AppendLine("          renglones_documentos.Cod_Art,   ")
            loComandoSeleccionar.AppendLine("          Articulos.Nom_Art + Substring(renglones_documentos.Comentario,1,450)    As  Nom_Art,   ")
            loComandoSeleccionar.AppendLine("          renglones_documentos.Renglon,   ")
            loComandoSeleccionar.AppendLine("          renglones_documentos.Can_Art AS Can_Art1,   ")
            loComandoSeleccionar.AppendLine("          renglones_documentos.Cod_Uni,   ")
            loComandoSeleccionar.AppendLine("          renglones_documentos.Precio1,   ")
            loComandoSeleccionar.AppendLine("          renglones_documentos.Mon_Net  As  Neto,   ")
            loComandoSeleccionar.AppendLine("          renglones_documentos.Por_Imp As  Por_Imp,   ")
            loComandoSeleccionar.AppendLine("          renglones_documentos.Cod_Imp,   ")
            loComandoSeleccionar.AppendLine("          renglones_documentos.Mon_Imp As  Impuesto   ")
            loComandoSeleccionar.AppendLine("FROM       cuentas_cobrar   ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN    renglones_documentos ON renglones_documentos.Documento = cuentas_cobrar.Documento   ")
            loComandoSeleccionar.AppendLine("			AND renglones_documentos.cod_tip =  cuentas_cobrar.cod_tip  ")
            loComandoSeleccionar.AppendLine("		AND renglones_documentos.origen = 'Ventas'  ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN    Articulos  ON Articulos.Cod_Art   =   renglones_documentos.Cod_Art   ")
            loComandoSeleccionar.AppendLine("		JOIN    Clientes ON Clientes.Cod_Cli = cuentas_cobrar.Cod_Cli   ")
            loComandoSeleccionar.AppendLine("		JOIN    Formas_Pagos ON Formas_Pagos.Cod_For = cuentas_cobrar.Cod_For   ")
            loComandoSeleccionar.AppendLine("		JOIN    Vendedores ON Vendedores.Cod_Ven =  cuentas_cobrar.Cod_Ven  ")
            loComandoSeleccionar.AppendLine("		JOIN    Tipos_Documentos ON Tipos_Documentos.Cod_tip =  cuentas_cobrar.Cod_tip  ")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCuentas_Cobrar_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCuentas_Cobrar_GPV.ReportSource = loObjetoReporte

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
' EAG: 15/08/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
