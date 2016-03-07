'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCorte_Cuenta_Pedido_CCSJ"
'-------------------------------------------------------------------------------------------'
Partial Class fCorte_Cuenta_Pedido_CCSJ

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Encabezado del Pedido.                                                                    *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("CREATE TABLE #tmpPedido(   Documento               CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Control                 CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Fec_Ini                 DATETIME,")
            loConsulta.AppendLine("                            Fec_Fin                 DATETIME,")
            loConsulta.AppendLine("                            Cod_Cli                 CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Rif                     CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nit                     CHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Cod_Ven                 CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Cod_For                 CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Cod_Tra                 CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nombre_Cliente          VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Comentario_Documento    VARCHAR(MAX) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Total_Bruto             DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Total_Impuesto          DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Total_Neto              DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Direccion               VARCHAR(MAX) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Diagnostico             CHAR(100) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Intervencion            CHAR(100) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Paciente_Codigo         CHAR(10) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Paciente_Nombre         CHAR(100) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Paciente_Rif            CHAR(30) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Representante_Codigo    CHAR(10) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Representante_Nombre    CHAR(100) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Representante_Rif       CHAR(30) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Seguro_Codigo           CHAR(10) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Seguro_Nombre           CHAR(100) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Seguro_Rif              CHAR(30) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Medico_Codigo           CHAR(10) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Medico_Nombre           CHAR(100) COLLATE DATABASE_DEFAULT DEFAULT(''),")
            loConsulta.AppendLine("                            Dias_Hospitalizacion    INTEGER,")
            loConsulta.AppendLine("                            Horas_Aproximadas       INTEGER,")
            loConsulta.AppendLine("                            Monto_Abonado           DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Cobertura_Inicial       DECIMAL(28, 10)")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpPedido(     Documento, Control, Fec_Ini, Fec_Fin, ")
            loConsulta.AppendLine("                            Cod_Cli, Rif, Nit, Cod_Ven, Cod_For, Cod_Tra,")
            loConsulta.AppendLine("                            Comentario_Documento, Horas_Aproximadas, Dias_Hospitalizacion, ")
            loConsulta.AppendLine("                            Total_Bruto, Total_Impuesto, Total_Neto)")
            loConsulta.AppendLine("SELECT  Documento, Control, Fec_Ini, Fec_Fin, ")
            loConsulta.AppendLine("        Cod_Cli, Rif, Nit, Cod_Ven, Cod_For, Cod_Tra,")
            loConsulta.AppendLine("        Comentario, Numerico1, Numerico2,")
            loConsulta.AppendLine("        Mon_Bru, Mon_Imp1, Mon_Net")
            loConsulta.AppendLine("FROM    Pedidos")
            loConsulta.AppendLine("WHERE    " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Intervención.                                                                            *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("UPDATE  #tmpPedido")
            loConsulta.AppendLine("SET     Intervencion = (SELECT TOP 1 Nom_Tra ")
            loConsulta.AppendLine("                        FROM Transportes ")
            loConsulta.AppendLine("                        WHERE Cod_Tra = #tmpPedido.Cod_Tra)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Datos adicionales de: paciente, médico, seguro, y representante.                         *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("DECLARE @lcDocumento CHAR(10);")
            loConsulta.AppendLine("DECLARE @lcCliente CHAR(10);")
            loConsulta.AppendLine("DECLARE @lcRif CHAR(20);")
            loConsulta.AppendLine("DECLARE @lcNit CHAR(20);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  TOP 1    ")
            loConsulta.AppendLine("        @lcDocumento = Documento, ")
            loConsulta.AppendLine("        @lcCliente   = Cod_Cli, ")
            loConsulta.AppendLine("        @lcRif       = Rif, ")
            loConsulta.AppendLine("        @lcNit       = Nit")
            loConsulta.AppendLine("FROM    #tmpPedido;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Si el cliente y el Rif son iguales (o si no hay Rif):                                    *")
            loConsulta.AppendLine("-- Entonces no hay seguro, y el cliente es el mismo paciente.                               *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("IF(@lcCliente = @lcRif OR @lcRif = '')")
            loConsulta.AppendLine("BEGIN")
            loConsulta.AppendLine("    UPDATE  #tmpPedido")
            loConsulta.AppendLine("    SET     Paciente_Codigo = Clientes.Cod_Cli,")
            loConsulta.AppendLine("            Paciente_Nombre = Clientes.Nom_Cli,")
            loConsulta.AppendLine("            Paciente_Rif    = Clientes.Rif,")
            loConsulta.AppendLine("            Diagnostico     = Clases_Clientes.Nom_Cla,")
            loConsulta.AppendLine("            Direccion       = Clientes.Dir_Fis")
            loConsulta.AppendLine("    FROM    Clientes ")
            loConsulta.AppendLine("        JOIN Clases_Clientes ON Clases_Clientes.Cod_Cla = Clientes.Cod_Cla ")
            loConsulta.AppendLine("    WHERE   Clientes.Cod_Cli = @lcCliente;")
            loConsulta.AppendLine("END")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Si el cliente y el Rif NO son iguales:                                                   *")
            loConsulta.AppendLine("-- Entonces el cliente es el seguro, y el Rif es el paciente.                               *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("ELSE")
            loConsulta.AppendLine("BEGIN")
            loConsulta.AppendLine("    UPDATE  #tmpPedido")
            loConsulta.AppendLine("    SET     Paciente_Codigo = Clientes.Cod_Cli,")
            loConsulta.AppendLine("            Paciente_Nombre = Clientes.Nom_Cli,")
            loConsulta.AppendLine("            Paciente_Rif    = Clientes.Rif,")
            loConsulta.AppendLine("            Diagnostico     = Clases_Clientes.Nom_Cla")
            loConsulta.AppendLine("    FROM    Clientes ")
            loConsulta.AppendLine("        JOIN Clases_Clientes ON Clases_Clientes.Cod_Cla = Clientes.Cod_Cla ")
            loConsulta.AppendLine("    WHERE   Clientes.Cod_Cli = @lcRif;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("    UPDATE  #tmpPedido")
            loConsulta.AppendLine("    SET     Seguro_Codigo = Clientes.Cod_Cli,")
            loConsulta.AppendLine("            Seguro_Nombre = Clientes.Nom_Cli,")
            loConsulta.AppendLine("            Seguro_Rif    = Clientes.Rif,")
            loConsulta.AppendLine("            Direccion     = Clientes.Dir_Fis")
            loConsulta.AppendLine("    FROM    Clientes ")
            loConsulta.AppendLine("    WHERE   Clientes.Cod_Cli = @lcCliente;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("END;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Nombre del cliente del documento.                                                        *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("UPDATE  #tmpPedido")
            loConsulta.AppendLine("SET     Nombre_Cliente = ")
            loConsulta.AppendLine("        (CASE WHEN RTRIM(Seguro_Nombre) > ''")
            loConsulta.AppendLine("             THEN Seguro_Nombre ")
            loConsulta.AppendLine("             ELSE Paciente_Nombre ")
            loConsulta.AppendLine("        END )")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Si el Nit está definido: Entonces el Nit es el representante del paciente.               *")
            loConsulta.AppendLine("-- Nota: en este caso la dirección del seguro priva sobre la del representante.             *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("IF ( EXISTS(SELECT * FROM Clientes WHERE cod_cli = @lcNit) )")
            loConsulta.AppendLine("BEGIN")
            loConsulta.AppendLine("    UPDATE  #tmpPedido")
            loConsulta.AppendLine("    SET     Representante_Codigo = Clientes.Cod_Cli,")
            loConsulta.AppendLine("            Representante_Nombre = Clientes.Nom_Cli,")
            loConsulta.AppendLine("            Representante_Rif    = Clientes.Rif,")
            loConsulta.AppendLine("            Direccion     = (CASE WHEN RTRIM(Seguro_Nombre) > ''")
            loConsulta.AppendLine("                                 THEN Direccion ")
            loConsulta.AppendLine("                                 ELSE Clientes.Dir_Fis ")
            loConsulta.AppendLine("                             END )")
            loConsulta.AppendLine("    FROM    Clientes ")
            loConsulta.AppendLine("    WHERE   Clientes.Cod_Cli = @lcNit;")
            loConsulta.AppendLine("END")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Si el Nit NO está definido: Entonces el paciente es su propio representante.             *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("ELSE")
            loConsulta.AppendLine("BEGIN")
            loConsulta.AppendLine("    UPDATE  #tmpPedido")
            loConsulta.AppendLine("    SET     Representante_Codigo = Paciente_Codigo,")
            loConsulta.AppendLine("            Representante_Nombre = Paciente_Nombre,")
            loConsulta.AppendLine("            Representante_Rif    = Paciente_Rif")
            loConsulta.AppendLine("END")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Datos del médico.                                                                        *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("UPDATE  #tmpPedido")
            loConsulta.AppendLine("SET     Medico_Codigo = Vendedores.Cod_Ven,")
            loConsulta.AppendLine("        Medico_Nombre = Vendedores.Nom_Ven")
            loConsulta.AppendLine("FROM    Vendedores ")
            loConsulta.AppendLine("WHERE   Vendedores.Cod_Ven = #tmpPedido.Cod_Ven;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Datos de Corte de Cuenta.                                                                *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("UPDATE  #tmpPedido")
            loConsulta.AppendLine("SET     Monto_Abonado       = COALESCE((SELECT TOP 1 Val_Num")
            loConsulta.AppendLine("                                        FROM campos_propiedades ")
            loConsulta.AppendLine("                                        WHERE Origen = '' AND clase = '' ")
            loConsulta.AppendLine("                                            AND tipo = '' AND cod_reg = '' ")
            loConsulta.AppendLine("                                            AND cod_pro = ''), 0),")
            loConsulta.AppendLine("        Cobertura_Inicial   = COALESCE((SELECT TOP 1 Val_Num")
            loConsulta.AppendLine("                                        FROM campos_propiedades ")
            loConsulta.AppendLine("                                        WHERE Origen = '' AND clase = '' ")
            loConsulta.AppendLine("                                            AND tipo = '' AND cod_reg = '' ")
            loConsulta.AppendLine("                                            AND cod_pro = ''), 0);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Renglones del documento.                                                                 *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("CREATE TABLE #tmpRenglones( Renglon                 INTEGER,")
            loConsulta.AppendLine("                            Cod_Art                 CHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nom_Art                 CHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Cod_Hon                 CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nom_Hon                 CHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Cantidad                DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Bruto                   DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Por_Imp                 DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Impuesto                DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Neto                    DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Cod_Dep                 CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nom_Dep                 CHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Orden                   DECIMAL(28, 10)")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpRenglones(Renglon, Cod_Art, Nom_Art, ")
            loConsulta.AppendLine("            Cantidad, Bruto, Por_Imp, Impuesto, Neto,")
            loConsulta.AppendLine("            Cod_Hon, Cod_Dep, Nom_Dep, Orden)")
            loConsulta.AppendLine("SELECT      ROW_NUMBER() OVER (")
            loConsulta.AppendLine("                ORDER BY COALESCE(P.Val_Num, 1000000000) ASC, ")
            loConsulta.AppendLine("                         D.Cod_Dep ASC, R.Renglon ASC")
            loConsulta.AppendLine("            )                                   AS Renglon, ")
            loConsulta.AppendLine("            A.Cod_Art                           AS Cod_Art, ")
            loConsulta.AppendLine("            A.Nom_Art                           AS Nom_Art, ")
            loConsulta.AppendLine("            R.Can_Art1                          AS Cantidad, ")
            loConsulta.AppendLine("            R.Precio1                           AS Bruto, ")
            loConsulta.AppendLine("            R.Por_Imp1                          AS Por_Imp, ")
            loConsulta.AppendLine("            R.Mon_Imp1                          AS Impuesto, ")
            loConsulta.AppendLine("            (R.Mon_Net+R.Mon_Imp1)              AS Neto,")
            loConsulta.AppendLine("            R.Cod_Ven                           AS Cod_Hon,")
            loConsulta.AppendLine("            D.Cod_Dep                           AS Cod_Dep,")
            loConsulta.AppendLine("            D.Nom_Dep                           AS Nom_Dep,")
            loConsulta.AppendLine("            COALESCE(P.Val_Num, 1000000000)      AS Orden")
            loConsulta.AppendLine("FROM        Renglones_Pedidos AS R")
            loConsulta.AppendLine("    JOIN    Articulos AS A On A.Cod_Art = R.Cod_Art")
            loConsulta.AppendLine("    JOIN    Departamentos AS D On D.Cod_Dep = A.Cod_Dep")
            loConsulta.AppendLine("    LEFT JOIN Campos_Propiedades AS P")
            loConsulta.AppendLine("        ON  P.Cod_Reg = D.Cod_Dep")
            loConsulta.AppendLine("        AND P.Cod_Pro = 'DEPORD'")
            loConsulta.AppendLine("        AND P.Origen = 'Departamentos'")
            loConsulta.AppendLine("WHERE       Documento = @lcDocumento")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Nombre de Médicos para Honorarios.                                                       *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("UPDATE #tmpRenglones")
            loConsulta.AppendLine("SET Nom_Hon = vendedores.nom_ven")
            loConsulta.AppendLine("FROM vendedores")
            loConsulta.AppendLine("WHERE #tmpRenglones.Cod_Dep = 'HON'")
            loConsulta.AppendLine("    AND #tmpRenglones.Cod_Hon = vendedores.Cod_Ven")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Datos de póliza y abonado.                                                               *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("UPDATE #tmpPedido")
            loConsulta.AppendLine("SET     Cobertura_Inicial = COALESCE((  SELECT TOP 1 Val_Num ")
            loConsulta.AppendLine("                                        FROM campos_propiedades ")
            loConsulta.AppendLine("                                        WHERE Origen = 'pedidos' AND clase = '' AND tipo = ''")
            loConsulta.AppendLine("                                            AND cod_reg = @lcDocumento AND cod_pro = '400-01-003'), 0),")
            loConsulta.AppendLine("        Monto_Abonado     = COALESCE((  SELECT SUM(mon_sal) ")
            loConsulta.AppendLine("                                        FROM Cuentas_Cobrar ")
            loConsulta.AppendLine("                                        WHERE cod_cli = #tmpPedido.Paciente_Codigo AND cod_tip = 'ADEL'")
            loConsulta.AppendLine("                                            AND status IN ('Pendiente','Afectado')), 0)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("-- Datos para el reporte.                                                                   *")
            loConsulta.AppendLine("-- ******************************************************************************************")
            loConsulta.AppendLine("SELECT  Documento, Control, Fec_Ini, Fec_Fin, Nombre_Cliente, ")
            loConsulta.AppendLine("        Dias_Hospitalizacion, Horas_Aproximadas, ")
            loConsulta.AppendLine("        Comentario_Documento, Diagnostico, Intervencion,")
            loConsulta.AppendLine("        Paciente_Codigo,  ")
            loConsulta.AppendLine("        (RTRIM(Paciente_Codigo) + ' - ' + Paciente_Nombre) AS Paciente_Nombre,")
            loConsulta.AppendLine("        Paciente_Rif,")
            loConsulta.AppendLine("        Representante_Codigo, ")
            loConsulta.AppendLine("        (RTRIM(Representante_Codigo) + ' - ' + Representante_Nombre) AS Representante_Nombre,")
            loConsulta.AppendLine("        Representante_Rif,")
            loConsulta.AppendLine("        Seguro_Codigo,")
            loConsulta.AppendLine("        (RTRIM(Seguro_Codigo) + ' - ' + Seguro_Nombre) AS Seguro_Nombre,")
            loConsulta.AppendLine("        Seguro_Rif,")
            loConsulta.AppendLine("        Medico_Codigo,")
            loConsulta.AppendLine("        (RTRIM(Medico_Codigo) + ' - ' + Medico_Nombre) AS Medico_Nombre,")
            loConsulta.AppendLine("        Direccion, Total_Bruto, Total_Impuesto, Total_Neto,")
            loConsulta.AppendLine("        Renglon, Cod_Art, Nom_Art, Cod_Hon, Nom_Hon, Cantidad, ")
            loConsulta.AppendLine("        Bruto, Por_Imp, Impuesto, Neto, Cod_Dep, Nom_Dep,")
            loConsulta.AppendLine("        (-#tmpPedido.Monto_Abonado) AS Monto_Abonado, ")
            loConsulta.AppendLine("        (-#tmpPedido.Cobertura_Inicial) AS Cobertura_Inicial ")
            loConsulta.AppendLine("FROM    #tmpPedido")
            loConsulta.AppendLine("    CROSS JOIN #tmpRenglones")
            loConsulta.AppendLine("ORDER BY Renglon")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString()) 
            Dim loDatos As New cusDatos.goDatos()
            Dim laDatosReporte As DataSet = loDatos.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCorte_Cuenta_Pedido_CCSJ", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCorte_Cuenta_Pedido_CCSJ.ReportSource = loObjetoReporte

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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 15/02/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 07/04/14: Se ajustó la dirección fiscal: mostrará la del Seguro, representante o     '
'                paciente, según el caso. Se ajustaron los campos días y horas (estaban     '
'                invertidos). Se agregó el RIF del seguro.                                  '
'-------------------------------------------------------------------------------------------'
