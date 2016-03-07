'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fAdmision_Pacientes"
'-------------------------------------------------------------------------------------------'
Partial Class fAdmision_Pacientes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	    Pedidos.Cod_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Nom_Cli ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Pedidos.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Pedidos.Nom_Cli END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Pedidos.Rif = '') THEN Clientes.Rif ELSE Pedidos.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("            Clientes.Nit, ")
            loConsulta.AppendLine("            (CASE WHEN (Clientes.Generico = 0) THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Pedidos.Dir_Fis,1, 200) = '') ")
            loConsulta.AppendLine("                    THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ")
            loConsulta.AppendLine("                    ELSE SUBSTRING(Pedidos.Dir_Fis,1, 200) ")
            loConsulta.AppendLine("                END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("            (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Pedidos.Telefonos = '') THEN Clientes.Telefonos ELSE Pedidos.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("            Clientes.Fax                            AS Fax, ")
            loConsulta.AppendLine("            Pedidos.Nom_Cli                         AS Nom_Gen, ")
            loConsulta.AppendLine("            Pedidos.Rif                             AS Rif_Gen, ")
            loConsulta.AppendLine("            Pedidos.Nit                             AS Nit_Gen, ")
            loConsulta.AppendLine("            Pedidos.Dir_Fis                         AS Dir_Gen, ")
            loConsulta.AppendLine("            Pedidos.Telefonos                       AS Tel_Gen, ")
            loConsulta.AppendLine("            Pedidos.Documento                       AS Documento, ")
            loConsulta.AppendLine("            Pedidos.Fec_Ini                         AS Fec_Ini, ")
            loConsulta.AppendLine("            Pedidos.Fec_Fin                         AS Fec_Fin, ")
            loConsulta.AppendLine("            Pedidos.Mon_Bru                         AS Mon_Bru, ")
            loConsulta.AppendLine("            Pedidos.Por_Des1                        AS Por_Des1, ")
            loConsulta.AppendLine("            Pedidos.Por_Rec1                        AS Por_Rec1, ")
            loConsulta.AppendLine("            Pedidos.Mon_Des1                        AS Mon_Des1, ")
            loConsulta.AppendLine("            Pedidos.Mon_Rec1                        AS Mon_Rec1, ")
            loConsulta.AppendLine("            Pedidos.Mon_Imp1                        AS Mon_Imp1, ")
            loConsulta.AppendLine("            Pedidos.Dis_Imp                         AS Dis_Imp, ")
            loConsulta.AppendLine("            Pedidos.Por_Imp1                        AS Por_Imp1,  ")
            loConsulta.AppendLine("            Pedidos.Mon_Net                         AS Mon_Net, ")
            loConsulta.AppendLine("            Pedidos.Cod_For                         AS Cod_For, ")
            loConsulta.AppendLine("            Formas_Pagos.Nom_For                    AS Nom_For, ")
            loConsulta.AppendLine("            Pedidos.Cod_Ven                         AS Cod_Ven, ")
            loConsulta.AppendLine("            Pedidos.Comentario                      AS Comentario, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven                      AS Nom_Ven, ")
            loConsulta.AppendLine("            Renglones_Pedidos.Cod_Art               AS Cod_Art, ")
            loConsulta.AppendLine("            (CASE WHEN Articulos.Generico = 0 ")
            loConsulta.AppendLine("                THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("			    ELSE Renglones_Pedidos.Notas ")
            loConsulta.AppendLine("		    END)                                       AS Nom_Art,  ")
            loConsulta.AppendLine("            Renglones_Pedidos.Renglon               AS Renglon, ")
            loConsulta.AppendLine("            Renglones_Pedidos.Can_Art1              AS Can_Art1, ")
            loConsulta.AppendLine("            Renglones_Pedidos.Cod_Uni               AS Cod_Uni, ")
            loConsulta.AppendLine("            Renglones_Pedidos.Precio1               AS Precio1, ")
            loConsulta.AppendLine("            Renglones_Pedidos.Mon_Net               As Neto, ")
            loConsulta.AppendLine("            Renglones_Pedidos.Por_Imp1              As Por_Imp, ")
            loConsulta.AppendLine("            Renglones_Pedidos.Cod_Imp               AS Cod_Imp, ")
            loConsulta.AppendLine("            Renglones_Pedidos.Mon_Imp1              As Impuesto, ")
            loConsulta.AppendLine("            Pedidos.Cod_Tra                         AS Cod_Tra, ")
            loConsulta.AppendLine("            Transportes.Nom_Tra                     AS Nom_Tra, ")
            loConsulta.AppendLine("            (CASE WHEN (Pedidos.Rif = '') ")
            loConsulta.AppendLine("                THEN SPACE(250) ")
            loConsulta.AppendLine("                ELSE Pedidos.Rif ")
            loConsulta.AppendLine("            END)                                    AS Cod_Pac, ")
            loConsulta.AppendLine("            SPACE(20)                               AS Rif_Pac, ")
            loConsulta.AppendLine("            SPACE(10)                               AS Sex_Pac, ")
            loConsulta.AppendLine("            SPACE(250)	                           AS Nom_Pac, ")
            loConsulta.AppendLine("            CAST('' AS VARCHAR(MAX))	               AS Dir_Pac, ")
            loConsulta.AppendLine("            SPACE(20)                               AS ADN_Pac, ")
            loConsulta.AppendLine("            SPACE(20)                               AS Mov_Pac, ")
            loConsulta.AppendLine("            SPACE(20)                               AS Tel_Pac, ")
            loConsulta.AppendLine("            CAST(0 AS INT)                          AS Eda_Pac, ")
            loConsulta.AppendLine("            Pedidos.Fec_Fin                         AS FDN_Pac, ")
            loConsulta.AppendLine("            SPACE(100)                              AS Nom_Tit, ")
            loConsulta.AppendLine("            SPACE(30)                               AS Rif_Tit, ")
            loConsulta.AppendLine("            SPACE(50)                               AS Tel_Rif ")
            loConsulta.AppendLine("INTO        #tmpPedidos001 ")
            loConsulta.AppendLine("FROM        Pedidos")
            loConsulta.AppendLine("    JOIN    Renglones_Pedidos   ON Pedidos.Documento    =   Renglones_Pedidos.Documento")
            loConsulta.AppendLine("    JOIN    Clientes            ON Pedidos.Cod_Cli      =   Clientes.Cod_Cli")
            loConsulta.AppendLine("    JOIN    Formas_Pagos        ON Pedidos.Cod_For      =   Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN    Vendedores          ON Pedidos.Cod_Ven      =   Vendedores.Cod_Ven")
            loConsulta.AppendLine("    JOIN    Transportes         ON Pedidos.Cod_Tra      =   Transportes.Cod_Tra")
            loConsulta.AppendLine("    JOIN    Articulos           ON Articulos.Cod_Art    =   Renglones_Pedidos.Cod_Art")
            loConsulta.AppendLine("WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Verifica si el Rif del documento tiene un código de paciente ")
            loConsulta.AppendLine("--válido (en caso de no ser válido: el cliente es el paciente)")
            loConsulta.AppendLine("IF ( NOT EXISTS(SELECT * FROM Clientes C JOIN #tmpPedidos001 P ON C.Cod_Cli = P.Cod_Pac ) ) ")
            loConsulta.AppendLine("BEGIN")
            loConsulta.AppendLine("    UPDATE  #tmpPedidos001")
            loConsulta.AppendLine("    SET Cod_Pac = Cod_Cli;")
            loConsulta.AppendLine("END;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE  #tmpPedidos001 ")
            loConsulta.AppendLine("SET     Nom_Pac = Pacientes.Nom_Cli,")
            loConsulta.AppendLine("        Rif_Pac = Pacientes.Rif,")
            loConsulta.AppendLine("        Dir_Pac = Pacientes.Dir_Fis,")
            loConsulta.AppendLine("        Sex_Pac = (CASE Pacientes.Sexo ")
            loConsulta.AppendLine("                    WHEN 'M' THEN 'Masculino' ")
            loConsulta.AppendLine("                    WHEN 'F' THEN 'Femenino' ")
            loConsulta.AppendLine("                    ELSE 'No Aplica' END),")
            loConsulta.AppendLine("        FDN_Pac = Pacientes.Fec_Nac,")
            loConsulta.AppendLine("        Tel_Pac = Pacientes.Telefonos,")
            loConsulta.AppendLine("        Mov_Pac = Pacientes.Movil,")
            loConsulta.AppendLine("        ADN_Pac = Pacientes.Are_Neg,")
            loConsulta.AppendLine("        Eda_Pac = FLOOR(DATEDIFF(DAY, Pacientes.Fec_Nac, CAST(GETDATE() AS FLOAT))/365.25)")
            loConsulta.AppendLine("FROM    Clientes AS Pacientes")
            loConsulta.AppendLine("WHERE   Cod_Pac = Pacientes.Cod_Cli")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Parámetros para campos adicionales ")
            loConsulta.AppendLine("DECLARE @lcPaciente CHAR(10);")
            loConsulta.AppendLine("DECLARE @lcDocumento CHAR(10);")
            loConsulta.AppendLine("SELECT  TOP 1")
            loConsulta.AppendLine("        @lcPaciente = cod_pac,")
            loConsulta.AppendLine("        @lcDocumento = Documento")
            loConsulta.AppendLine("FROM    #tmpPedidos001;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Datos del Titular/representante")
            loConsulta.AppendLine("DECLARE @lcNit CHAR(30);")
            loConsulta.AppendLine("SET @lcNit = (SELECT TOP 1 Nit_Gen FROM #tmpPedidos001);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("IF(EXISTS (SELECT * FROM Clientes WHERE Cod_Cli = @lcNit))")
            loConsulta.AppendLine("BEGIN")
            loConsulta.AppendLine("     UPDATE  #tmpPedidos001")
            loConsulta.AppendLine("     SET     Nom_Tit = Clientes.Nom_Cli,")
            loConsulta.AppendLine("             Rif_Tit = Clientes.Rif,")
            loConsulta.AppendLine("             Tel_Rif = Clientes.Telefonos")
            loConsulta.AppendLine("     FROM Clientes ")
            loConsulta.AppendLine("     WHERE Clientes.Cod_Cli = @lcNit ")
            loConsulta.AppendLine("END;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Campos adicionales")
            loConsulta.AppendLine("CREATE TABLE #tmpCamposPropiedades( ")
            loConsulta.AppendLine("    Estado_Civil_Paciente       VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Lugar_Nacimiento            VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Ocupacion_Paciente          VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Tipo_Admision               VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Diagnostico                 VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Clave_Poliza                VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Numero_Poliza               VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Covertura_Inicial_Poliza    DECIMAL(28, 10),")
            loConsulta.AppendLine("    Deposito_Abono_Poliza       DECIMAL(28, 10),")
            loConsulta.AppendLine("    Tipo_Poliza                 VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Persona_Carta_Aval          VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Numero_Carta_Aval           VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Notas_Paciente              VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("    Notas_Admision              VARCHAR(MAX) COLLATE DATABASE_DEFAULT);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpCamposPropiedades(  Estado_Civil_Paciente, Lugar_Nacimiento, Ocupacion_Paciente,")
            loConsulta.AppendLine("                                    Tipo_Admision, Diagnostico, Clave_Poliza, Numero_Poliza, ")
            loConsulta.AppendLine("                                    Covertura_Inicial_Poliza, Deposito_Abono_Poliza, Tipo_Poliza,")
            loConsulta.AppendLine("                                    Persona_Carta_Aval, Numero_Carta_Aval, Notas_Paciente, Notas_Admision)")
            loConsulta.AppendLine("SELECT")
            loConsulta.AppendLine("    COALESCE((  SELECT  TOP 1 (CASE WHEN Val_Car>'' THEN Val_Car ELSE '[N/A]' END)  ")
            loConsulta.AppendLine("                FROM    campos_propiedades ")
            loConsulta.AppendLine("                WHERE   Origen = 'clientes' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcPaciente AND cod_pro = '100-10-002'), '[N/A]')         AS Estado_Civil_Paciente,")
            loConsulta.AppendLine("    COALESCE((  SELECT  TOP 1 (CASE WHEN Val_Car>'' THEN Val_Car ELSE '[N/A]' END)  ")
            loConsulta.AppendLine("                FROM    campos_propiedades ")
            loConsulta.AppendLine("                WHERE   Origen = 'clientes' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcPaciente AND cod_pro = '100-02-001'), '[N/A]')         AS Lugar_Nacimiento,")
            loConsulta.AppendLine("    COALESCE((  SELECT  TOP 1 (CASE WHEN Val_Car>'' THEN Val_Car ELSE '[N/A]' END)  ")
            loConsulta.AppendLine("                FROM    campos_propiedades ")
            loConsulta.AppendLine("                WHERE   Origen = 'clientes' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcPaciente AND cod_pro = '100-10-001'), '[N/A]')         AS Ocupacion_Paciente,")
            loConsulta.AppendLine("    COALESCE((  SELECT  TOP 1 (CASE WHEN Val_Car>'' THEN Val_Car ELSE '[N/A]' END) ")
            loConsulta.AppendLine("                FROM    campos_propiedades ")
            loConsulta.AppendLine("                WHERE   Origen = 'clientes' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcPaciente AND cod_pro = '100-00-002'), '[N/A]')        AS Tipo_Admision,")
            loConsulta.AppendLine("    COALESCE((  SELECT  TOP 1 (CASE WHEN COALESCE(CC.Cod_Cla, CP.Val_Car)>'' ")
            loConsulta.AppendLine("                                THEN RTRIM(COALESCE(CC.Cod_Cla, Val_Car)) + ' ' + RTRIM(CC.Nom_Cla) ")
            loConsulta.AppendLine("                                ELSE '[N/A]' END) ")
            loConsulta.AppendLine("                FROM    campos_propiedades CP")
            loConsulta.AppendLine("                    LEFT JOIN Clases_Clientes CC ON CC.Cod_Cla = CP.Val_Car")
            loConsulta.AppendLine("                WHERE   Origen = 'pedidos' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcDocumento AND cod_pro = '100-50-001'), '[N/A]')        AS Diagnostico,")
            loConsulta.AppendLine("    COALESCE((  SELECT TOP 1 (CASE WHEN Val_Car>'' THEN Val_Car ELSE '[N/A]' END) ")
            loConsulta.AppendLine("                FROM campos_propiedades ")
            loConsulta.AppendLine("                WHERE Origen = 'pedidos' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcDocumento AND cod_pro = '400-01-002'), '[N/A]')        AS Clave_Poliza,")
            loConsulta.AppendLine("    COALESCE((  SELECT TOP 1 (CASE WHEN Val_Car>'' THEN Val_Car ELSE '[N/A]' END) ")
            loConsulta.AppendLine("                FROM campos_propiedades ")
            loConsulta.AppendLine("                WHERE Origen = 'pedidos' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcDocumento AND cod_pro = '400-01-001'), '[N/A]')        AS Numero_Poliza,")
            loConsulta.AppendLine("    COALESCE((  SELECT TOP 1 Val_Num ")
            loConsulta.AppendLine("                FROM campos_propiedades ")
            loConsulta.AppendLine("                WHERE Origen = 'pedidos' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcDocumento AND cod_pro = '400-01-003'), 0)              AS Covertura_Inicial_Poliza,")
            loConsulta.AppendLine("    COALESCE((  SELECT SUM(mon_sal) ")
            loConsulta.AppendLine("                FROM Cuentas_Cobrar ")
            loConsulta.AppendLine("                WHERE cod_cli = @lcPaciente AND cod_tip = 'ADEL'")
            loConsulta.AppendLine("                    AND status IN ('Pendiente','Afectado')), 0)                             AS Deposito_Abono_Poliza,")
            loConsulta.AppendLine("    COALESCE((  SELECT TOP 1 Val_Car ")
            loConsulta.AppendLine("                FROM campos_propiedades ")
            loConsulta.AppendLine("                WHERE Origen = 'pedidos' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcDocumento AND cod_pro = '100-50-002'), '[N/A]')        AS Tipo_Poliza,")
            loConsulta.AppendLine("    COALESCE((  SELECT TOP 1 (CASE WHEN Val_Car>'' THEN Val_Car ELSE '[N/A]' END) ")
            loConsulta.AppendLine("                FROM campos_propiedades ")
            loConsulta.AppendLine("                WHERE Origen = 'pedidos' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcDocumento AND cod_pro = '400-01-004'), '[N/A]')        AS Persona_Carta_Aval,")
            loConsulta.AppendLine("    COALESCE((  SELECT TOP 1 (CASE WHEN Val_Car>'' THEN Val_Car ELSE '[N/A]' END) ")
            loConsulta.AppendLine("                FROM campos_propiedades ")
            loConsulta.AppendLine("                WHERE Origen = 'pedidos' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcDocumento AND cod_pro = '400-01-007'), '[N/A]')        AS Numero_Carta_Aval,")
            loConsulta.AppendLine("    COALESCE((  SELECT TOP 1 Val_Mem ")
            loConsulta.AppendLine("                FROM campos_propiedades ")
            loConsulta.AppendLine("                WHERE Origen = 'pedidos' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcDocumento AND cod_pro = '900-01-001'), '')             AS Notas_Paciente,")
            loConsulta.AppendLine("    COALESCE((  SELECT TOP 1 Val_Mem ")
            loConsulta.AppendLine("                FROM campos_propiedades ")
            loConsulta.AppendLine("                WHERE Origen = 'pedidos' AND clase = '' AND tipo = '' ")
            loConsulta.AppendLine("                    AND cod_reg = @lcDocumento AND cod_pro = '900-01-002'), '')             AS Notas_Admision")
            loConsulta.AppendLine(";")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	#tmpPedidos001.*,")
            loConsulta.AppendLine("        #tmpCamposPropiedades.*")
            loConsulta.AppendLine("FROM    #tmpPedidos001")
            loConsulta.AppendLine("    CROSS JOIN #tmpCamposPropiedades;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- DROP TABLE #tmpCamposPropiedades;")
            loConsulta.AppendLine("-- DROP TABLE #tmpPedidos001;")
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fAdmision_Pacientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfAdmision_Pacientes.ReportSource = loObjetoReporte

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
'-------------------------------------------------------------------------------'
' Fin del codigo                                                                '
'-------------------------------------------------------------------------------'
' JJD: 29/01/14: Ajuste del formato para adecuarlo a la "Admision de Pacientes" '
'-------------------------------------------------------------------------------'
' JJD: 30/01/14: Ajuste del formato para adecuarlo a la "Admision de Pacientes" '
'-------------------------------------------------------------------------------'
' RJG: 15/02/14: Se agregaron campos faltantes (pendiente los campos que vienen '
'                de propiedades).                                               '
'-------------------------------------------------------------------------------'
' RJG: 17/02/14: Se enlazaron los campos que vienen de propiedades.             '
'-------------------------------------------------------------------------------'
' RJG: 02/04/14: Se agregaron campos del representante. Ajustes menores en      '
'                interface.                                                     '
'-------------------------------------------------------------------------------'
