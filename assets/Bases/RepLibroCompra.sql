USE [DM_ADM]
GO
/****** Object:  StoredProcedure [dbo].[RepLibroCompra]    Script Date: 25/4/2025 9:22:39 a. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		SOFTECH SISTEMAS
-- Create date: <03/23/2011>
-- Last Update: <2021-02-19>
-- Description:	<Reporte de Libro de Compras>
-- =============================================
ALTER PROCEDURE [dbo].[RepLibroCompra] 
-- Add the parameters for the stored procedure here
    @sCo_fecha_d SMALLDATETIME = NULL ,
    @sCo_fecha_h SMALLDATETIME = NULL ,
    @cCo_Sucursal CHAR(6) = NULL ,
    @bIncluirOrden varchar(2) = NULL,
	@bImprimirColumnImport VARCHAR(2) = NULL,
	@bImprimirColumnArt33 VARCHAR(2) = NULL,
	@bImprimirColumnArt34 VARCHAR(2) = NULL,
    @sCampOrderBy VARCHAR(16) = NULL ,
    @sDir VARCHAR(6) = NULL ,
    @bHeaderRep BIT = 0
AS 
    BEGIN
        SET NOCOUNT ON ;
	
        IF @sCo_fecha_h IS NOT NULL 
            SET @sCo_fecha_h = DATEADD(ss, -60, DATEADD(day, 1, @sCo_fecha_h))
			
		IF @bImprimirColumnImport IS NULL
			SET @bImprimirColumnImport = 'NO'

		IF @bImprimirColumnArt33 IS NULL
			SET @bImprimirColumnArt33 = 'SI'

		IF @bImprimirColumnArt34 IS NULL
			SET @bImprimirColumnArt34 = 'SI'

/*---------------------------------------------------------------------------------------------------------
Esta parte es para obtener los documentos y darle formato a cada registro del libro de compras
---------------------------------------------------------------------------------------------------------*/

        DECLARE @temp TABLE
            (
              [nro_doc] [char](20) ,
              [co_tipo_doc] [char](6) ,
              [fecha_emis] [smalldatetime] ,
              [fe_us_in] [smalldatetime] ,
              [total_neto] [decimal](18, 2) ,
              [co_prov] [char](16) ,
              [prov_des] [char](100) ,
              [r] [char](18) ,
              [tipo_prov] [char](4) ,
              [contrib] [bit] ,
              [nac] [char](1) ,
              [co_sucu_in] [char](6) ,
              [nro_orig] [char](20) ,
              [doc_orig] [char](6) ,
              [aut] [bit] ,
              [co_mone] [char](6) ,
              [nro_fact] [varchar](20) ,
              [n_control] [char](20) ,
              [anulado] [bit] ,
              [fec_reg] [smalldatetime] ,
              [ven_ter] [bit] ,
              [base_imp] [decimal](18, 2) ,
              [tipo_imp] [char](1) ,
              [tasa] [decimal](18, 5) ,
              [monto_imp] [decimal](18, 2) ,
              [doc_afec] [char](20) ,
              [compras_exentas] [decimal](18, 2) ,
              [base_imponible] [decimal](18, 2) ,
              [monto_ret_imp] [decimal](18, 2) ,
              [monto_ret_imp_tercero] [decimal](18, 2) ,
              [num_comprobante] [char](14) ,
              [fec_comprobante] [smalldatetime] ,
              /************************************/
			  [base_imponible_scf] [decimal] (18,2),
			  [monto_imp_scf] [decimal] (18,2),
			  /**Factura al Tesoro Nacional**/
			  [FTNac_fechaEmis] [smalldatetime],
			  [num_plan_impor] [char](40),
			  [num_exp_impor] [char](40),
			  [valorMercancia] [decimal](18, 2),
			  [der_impor] [decimal](18, 2),
			  [tasa_regimenAplic] [decimal](18, 2),
			  [otrosGravables] [decimal](18, 2),
			  [otrosExentos] [decimal](18, 2),
			  /**************Art.34**************/
			  [base_imponible_deducible] [decimal] (18,2),
			  [monto_imp_deducible] [decimal] (18,2),
			  [base_imponible_prorrateo] [decimal] (18,2),
			  [monto_imp_prorrateo] [decimal] (18,2),

			  [monto_igtf] [decimal] (18,2), --IGTF
			  
              [descrip1] [varchar](100) ,
              [base_imp1] [decimal](18, 2) ,
              [monto_imp1] [decimal](18, 2) ,
              [retenido1] [decimal](18, 2) ,
              [retenidoter1] [decimal](18, 2) ,
              [descrip2] [varchar](100) ,
              [base_imp2] [decimal](18, 2) ,
              [monto_imp2] [decimal](18, 2) ,
              [retenido2] [decimal](18, 2) ,
              [retenidoter2] [decimal](18, 2) ,
              [descrip3] [varchar](100) ,
              [base_imp3] [decimal](18, 2) ,
              [monto_imp3] [decimal](18, 2) ,
              [retenido3] [decimal](18, 2) ,
              [retenidoter3] [decimal](18, 2) ,
              [descrip4] [varchar](100) ,
              [base_imp4] [decimal](18, 2) ,
              [monto_imp4] [decimal](18, 2) ,
              [retenido4] [decimal](18, 2) ,
              [retenidoter4] [decimal](18, 2) ,
              [descrip5] [varchar](100) ,
              [base_imp5] [decimal](18, 2) ,
              [monto_imp5] [decimal](18, 2) ,
              [retenido5] [decimal](18, 2) ,
              [retenidoter5] [decimal](18, 2) ,
              [descrip6] [varchar](100) ,
              [base_imp6] [decimal](18, 2) ,
              [monto_imp6] [decimal](18, 2) ,
              [retenido6] [decimal](18, 2) ,
              [retenidoter6] [decimal](18, 2) ,
              [descrip7] [varchar](100) ,
              [base_imp7] [decimal](18, 2) ,
              [monto_imp7] [decimal](18, 2) ,
              [retenido7] [decimal](18, 2) ,
              [retenidoter7] [decimal](18, 2) ,
              [descrip8] [varchar](100) ,
              [base_imp8] [decimal](18, 2) ,
              [monto_imp8] [decimal](18, 2) ,
              [retenido8] [decimal](18, 2) ,
              [retenidoter8] [decimal](18, 2) ,
              [descrip9] [varchar](100) ,
              [base_imp9] [decimal](18, 2) ,
              [monto_imp9] [decimal](18, 2) ,
              [retenido9] [decimal](18, 2) ,
              [retenidoter9] [decimal](18, 2) ,
              [descrip10] [varchar](100) ,
              [base_imp10] [decimal](18, 2) ,
              [monto_imp10] [decimal](18, 2) ,
              [retenido10] [decimal](18, 2) ,
              [retenidoter10] [decimal](18, 2) ,
              [descrip11] [varchar](100) ,
              [base_imp11] [decimal](18, 2) ,
              [monto_imp11] [decimal](18, 2) ,
              [retenido11] [decimal](18, 2) ,
              [retenidoter11] [decimal](18, 2) ,
              [descrip12] [varchar](100) ,
              [base_imp12] [decimal](18, 2) ,
              [monto_imp12] [decimal](18, 2) ,
              [retenido12] [decimal](18, 2) ,
              [retenidoter12] [decimal](18, 2) ,
              [descrip13] [varchar](100) ,
              [base_imp13] [decimal](18, 2) ,
              [monto_imp13] [decimal](18, 2) ,
              [retenido13] [decimal](18, 2) ,
              [retenidoter13] [decimal](18, 2),
              /*------------------------------*/
			  [descrip14] [varchar](100) ,
              [base_imp14] [decimal](18, 2) ,
              [monto_imp14] [decimal](18, 2) ,
              [retenido14] [decimal](18, 2) ,
              [retenidoter14] [decimal](18, 2),
			  [descrip15] [varchar](100) ,
              [base_imp15] [decimal](18, 2) ,
              [monto_imp15] [decimal](18, 2) ,
              [retenido15] [decimal](18, 2) ,
              [retenidoter15] [decimal](18, 2),
			  /************Art.34************/
			  [descrip16] [varchar](100) ,
              [base_imp16] [decimal](18, 2) ,
              [monto_imp16] [decimal](18, 2) ,
              [retenido16] [decimal](18, 2) ,
              [retenidoter16] [decimal](18, 2),
			  [descrip17] [varchar](100) ,
              [base_imp17] [decimal](18, 2) ,
              [monto_imp17] [decimal](18, 2) ,
              [retenido17] [decimal](18, 2) ,
              [retenidoter17] [decimal](18, 2),
			  [descrip18] [varchar](100) ,
              [base_imp18] [decimal](18, 2) ,
              [monto_imp18] [decimal](18, 2) ,
              [retenido18] [decimal](18, 2) ,
              [retenidoter18] [decimal](18, 2),
			  [descrip19] [varchar](100) ,
              [base_imp19] [decimal](18, 2) ,
              [monto_imp19] [decimal](18, 2) ,
              [retenido19] [decimal](18, 2) ,
              [retenidoter19] [decimal](18, 2),
			  --Se necesitan 52 campos para poder abastecer un cuadro resumen que muestre todas las tasas que permite Profit:
			  [descrip20] [varchar](100) ,
              [base_imp20] [decimal](18, 2) ,
              [monto_imp20] [decimal](18, 2) ,
              [retenido20] [decimal](18, 2) ,
              [retenidoter20] [decimal](18, 2) ,
              [descrip21] [varchar](100) ,
              [base_imp21] [decimal](18, 2) ,
              [monto_imp21] [decimal](18, 2) ,
              [retenido21] [decimal](18, 2) ,
              [retenidoter21] [decimal](18, 2) ,
              [descrip22] [varchar](100) ,
              [base_imp22] [decimal](18, 2) ,
              [monto_imp22] [decimal](18, 2) ,
              [retenido22] [decimal](18, 2) ,
              [retenidoter22] [decimal](18, 2) ,
              [descrip23] [varchar](100) ,
              [base_imp23] [decimal](18, 2) ,
              [monto_imp23] [decimal](18, 2) ,
              [retenido23] [decimal](18, 2) ,
              [retenidoter23] [decimal](18, 2) ,
              [descrip24] [varchar](100) ,
              [base_imp24] [decimal](18, 2) ,
              [monto_imp24] [decimal](18, 2) ,
              [retenido24] [decimal](18, 2) ,
              [retenidoter24] [decimal](18, 2) ,
              [descrip25] [varchar](100) ,
              [base_imp25] [decimal](18, 2) ,
              [monto_imp25] [decimal](18, 2) ,
              [retenido25] [decimal](18, 2) ,
              [retenidoter25] [decimal](18, 2) ,
              [descrip26] [varchar](100) ,
              [base_imp26] [decimal](18, 2) ,
              [monto_imp26] [decimal](18, 2) ,
              [retenido26] [decimal](18, 2) ,
              [retenidoter26] [decimal](18, 2) ,
              [descrip27] [varchar](100) ,
              [base_imp27] [decimal](18, 2) ,
              [monto_imp27] [decimal](18, 2) ,
              [retenido27] [decimal](18, 2) ,
              [retenidoter27] [decimal](18, 2) ,
              [descrip28] [varchar](100) ,
              [base_imp28] [decimal](18, 2) ,
              [monto_imp28] [decimal](18, 2) ,
              [retenido28] [decimal](18, 2) ,
              [retenidoter28] [decimal](18, 2) ,
              [descrip29] [varchar](100) ,
              [base_imp29] [decimal](18, 2) ,
              [monto_imp29] [decimal](18, 2) ,
              [retenido29] [decimal](18, 2) ,
              [retenidoter29] [decimal](18, 2) ,
              [descrip30] [varchar](100) ,
              [base_imp30] [decimal](18, 2) ,
              [monto_imp30] [decimal](18, 2) ,
              [retenido30] [decimal](18, 2) ,
              [retenidoter30] [decimal](18, 2) ,
              [descrip31] [varchar](100) ,
              [base_imp31] [decimal](18, 2) ,
              [monto_imp31] [decimal](18, 2) ,
              [retenido31] [decimal](18, 2) ,
              [retenidoter31] [decimal](18, 2),
			  [descrip32] [varchar](100) ,
              [base_imp32] [decimal](18, 2) ,
              [monto_imp32] [decimal](18, 2) ,
              [retenido32] [decimal](18, 2) ,
              [retenidoter32] [decimal](18, 2) ,
              [descrip33] [varchar](100) ,
              [base_imp33] [decimal](18, 2) ,
              [monto_imp33] [decimal](18, 2) ,
              [retenido33] [decimal](18, 2) ,
              [retenidoter33] [decimal](18, 2) ,
              [descrip34] [varchar](100) ,
              [base_imp34] [decimal](18, 2) ,
              [monto_imp34] [decimal](18, 2) ,
              [retenido34] [decimal](18, 2) ,
              [retenidoter34] [decimal](18, 2) ,
              [descrip35] [varchar](100) ,
              [base_imp35] [decimal](18, 2) ,
              [monto_imp35] [decimal](18, 2) ,
              [retenido35] [decimal](18, 2) ,
              [retenidoter35] [decimal](18, 2) ,
              [descrip36] [varchar](100) ,
              [base_imp36] [decimal](18, 2) ,
              [monto_imp36] [decimal](18, 2) ,
              [retenido36] [decimal](18, 2) ,
              [retenidoter36] [decimal](18, 2) ,
              [descrip37] [varchar](100) ,
              [base_imp37] [decimal](18, 2) ,
              [monto_imp37] [decimal](18, 2) ,
              [retenido37] [decimal](18, 2) ,
              [retenidoter37] [decimal](18, 2),
			  [descrip38] [varchar](100) ,
              [base_imp38] [decimal](18, 2) ,
              [monto_imp38] [decimal](18, 2) ,
              [retenido38] [decimal](18, 2) ,
              [retenidoter38] [decimal](18, 2) ,
              [descrip39] [varchar](100) ,
              [base_imp39] [decimal](18, 2) ,
              [monto_imp39] [decimal](18, 2) ,
              [retenido39] [decimal](18, 2) ,
              [retenidoter39] [decimal](18, 2) ,
              [descrip40] [varchar](100) ,
              [base_imp40] [decimal](18, 2) ,
              [monto_imp40] [decimal](18, 2) ,
              [retenido40] [decimal](18, 2) ,
              [retenidoter40] [decimal](18, 2) ,
              [descrip41] [varchar](100) ,
              [base_imp41] [decimal](18, 2) ,
              [monto_imp41] [decimal](18, 2) ,
              [retenido41] [decimal](18, 2) ,
              [retenidoter41] [decimal](18, 2) ,
              [descrip42] [varchar](100) ,
              [base_imp42] [decimal](18, 2) ,
              [monto_imp42] [decimal](18, 2) ,
              [retenido42] [decimal](18, 2) ,
              [retenidoter42] [decimal](18, 2) ,
              [descrip43] [varchar](100) ,
              [base_imp43] [decimal](18, 2) ,
              [monto_imp43] [decimal](18, 2) ,
              [retenido43] [decimal](18, 2) ,
              [retenidoter43] [decimal](18, 2) ,
              [descrip44] [varchar](100) ,
              [base_imp44] [decimal](18, 2) ,
              [monto_imp44] [decimal](18, 2) ,
              [retenido44] [decimal](18, 2) ,
              [retenidoter44] [decimal](18, 2) ,
              [descrip45] [varchar](100) ,
              [base_imp45] [decimal](18, 2) ,
              [monto_imp45] [decimal](18, 2) ,
              [retenido45] [decimal](18, 2) ,
              [retenidoter45] [decimal](18, 2) ,
              [descrip46] [varchar](100) ,
              [base_imp46] [decimal](18, 2) ,
              [monto_imp46] [decimal](18, 2) ,
              [retenido46] [decimal](18, 2) ,
              [retenidoter46] [decimal](18, 2) ,
              [descrip47] [varchar](100) ,
              [base_imp47] [decimal](18, 2) ,
              [monto_imp47] [decimal](18, 2) ,
              [retenido47] [decimal](18, 2) ,
              [retenidoter47] [decimal](18, 2),
			  [descrip48] [varchar](100) ,
              [base_imp48] [decimal](18, 2) ,
              [monto_imp48] [decimal](18, 2) ,
              [retenido48] [decimal](18, 2) ,
              [retenidoter48] [decimal](18, 2) ,
              [descrip49] [varchar](100) ,
              [base_imp49] [decimal](18, 2) ,
              [monto_imp49] [decimal](18, 2) ,
              [retenido49] [decimal](18, 2) ,
              [retenidoter49] [decimal](18, 2) ,
              [descrip50] [varchar](100) ,
              [base_imp50] [decimal](18, 2) ,
              [monto_imp50] [decimal](18, 2) ,
              [retenido50] [decimal](18, 2) ,
              [retenidoter50] [decimal](18, 2) ,
              [descrip51] [varchar](100) ,
              [base_imp51] [decimal](18, 2) ,
              [monto_imp51] [decimal](18, 2) ,
              [retenido51] [decimal](18, 2) ,
              [retenidoter51] [decimal](18, 2) ,
              [descrip52] [varchar](100) ,
              [base_imp52] [decimal](18, 2) ,
              [monto_imp52] [decimal](18, 2) ,
              [retenido52] [decimal](18, 2) ,
              [retenidoter52] [decimal](18, 2)
			  )

        DECLARE
            @nro_doc CHAR(20) ,
            @co_tipo_doc CHAR(6) ,
            @fecha_emis SMALLDATETIME ,
            @fe_us_in SMALLDATETIME ,
            @total_neto DECIMAL(18, 2) ,
            @co_prov CHAR(16) ,
            @prov_des CHAR(100) ,
            @r CHAR(18) ,
            @tipo_prov CHAR(4) ,
            @contrib BIT ,
            @nac CHAR(1) ,
            @co_sucu_in CHAR(6) ,
            @nro_orig CHAR(20) ,
            @doc_orig CHAR(6) ,
            @aut BIT ,
            @co_mone CHAR(6) ,
            @nro_fact VARCHAR(20) ,
            @n_control CHAR(20) ,
            @anulado BIT ,
            @fec_reg SMALLDATETIME ,
            @fec_comprobante SMALLDATETIME ,
            @base_imp DECIMAL(18, 2) ,
            @tipo_imp CHAR(1) ,
            @tasa DECIMAL(18, 2) ,
            @monto_imp DECIMAL(18, 2) ,
            @doc_afec CHAR(20) ,
            @compras_exentas DECIMAL(18, 2) ,
            @exento_Actual DECIMAL(18, 2) ,
            @monto_ret_imp DECIMAL(18, 2) ,
            @monto_ret_imp_tercero DECIMAL(18, 2) ,
            @base_imponible DECIMAL(18, 2) ,
            @num_comprobante CHAR(14),
            /**********************************/
			@base_imponible_scf DECIMAL(18,2),
			@monto_imp_scf DECIMAL (18,2),
			/**Factura al Tesoro Nacional**/
			@FTNac_fechaEmis SMALLDATETIME,
			@num_plan_impor CHAR(40),
			@num_exp_impor CHAR(40),
			@valorMercancia DECIMAL(18, 2),
			@der_impor DECIMAL(18, 2),
			@tasa_regimenAplic DECIMAL(18, 2),
			@otrosGravables DECIMAL(18, 2),
			@otrosExentos DECIMAL (18, 2),
			@compras_exentas_FCI DECIMAL (18, 2),
			/**************Art.34**************/
			@base_imponible_deducible DECIMAL(18,2),
			@monto_imp_deducible DECIMAL (18,2),
			@base_imponible_prorrateo DECIMAL(18,2),
			@monto_imp_prorrateo DECIMAL (18,2) , 

			@monto_igtf DECIMAL (18,2) -- IGTF

	  
        DECLARE
            @old_nro_doc CHAR(20) ,
            @old_co_tipo_doc CHAR(6) ,
            @old_tasa DECIMAL(18, 5) ,
            @lupdate INT 


        DECLARE Tempdocs SCROLL CURSOR
        FOR        
        select A.* from (
			(
				SELECT *
				FROM documentoslibrocompras2(@sCo_fecha_d, @sCo_fecha_h)
				--Sit.# 9986 
				WHERE ( @cCo_Sucursal IS NULL
					  OR co_sucu_in = @cCo_Sucursal)
				--!Sit.# 9986 
			)
			union all    
			SELECT OP.ord_num as nro_doc, 'OPAG' as co_tipo_doc, OP.fecha as fecha_emis, OP.fe_us_in, sum(OPR.monto_d - OPR.monto_h) * OP.tasa
			as total_neto, OP.cod_ben as co_prov, BE.ben_des as prov_des,BE.rif as r, 1 as nac, 
			case when BE.rif ='' then 'SR' 
			-->>JN 20200212 Sit# 97826
			--case when BE.tipo_per='2' then 'NR' else 
			--case when BE.tipo_per='4' then 'ND' else 
			--space(2) end end 
			ELSE CASE WHEN BE.rif IS NULL THEN 'SR'
			ELSE CASE WHEN BE.tipo_per = '1' THEN 'PNR'
            ELSE CASE WHEN BE.tipo_per = '2' THEN 'PNNR'
			ELSE CASE WHEN BE.tipo_per = '3' THEN 'PJD'
            ELSE CASE WHEN BE.tipo_per = '4' THEN 'PJND'
			ELSE CASE WHEN BE.tipo_per = '6' THEN 'TN'
            ELSE SPACE(4)
			  END
				END
                 END
                  END
                   END
				    END                  
			--<<JN 20200212 Sit# 97826
			END AS tipo_prov,
			OP.co_sucu_in, SPACE(20) as nro_orig, SPACE(6) as doc_orig, 1 as aut, OP.co_mone, SPACE(6) as nro_fact,	SPACE(6) as n_control, 
			OP.anulado, OP.fecha as fec_reg, round(SUM(case when OPR.monto_d > 0 then 1 else -1 end * OPR.monto_obj),2) * OP.tasa as base_imp, 
			OPR.tipo_imp, [dbo].[TasaImpuestoSobreVentaAUnaFecha](OPR.tipo_imp,OP.fecha,0) as tasa, 
			round(SUM(case when OPR.monto_d > 0 then 1 else -1 end * OPR.monto_iva),2)as monto_imp, SPACE(20) as doc_afec, 0 as compras_exentas, 
			0 as base_imponible, 0 as monto_ret_imp, 0 as monto_ret_imp_tercero,SPACE(14) as num_comprobante, '' as fec_comprobante
			/***************************************/
			, 0.00 base_imponible_scf, 0.00 monto_imp_scf,
			/**Factura al Tesoro Nacional**/
			NULL AS FTNac_fechaEmis,
			NULL AS num_plan_impor, NULL AS num_exp_impor, NULL AS valorMercancia, NULL AS der_impor, NULL AS tasa_regimenAplic, NULL AS otrosGravables, NULL AS otrosExentos, NULL AS ComprasExentas_FCI,
			/**************Art.34**************/
			0.00 AS base_imponible_deducible, 0.00 AS monto_imp_deducible, 0.00 AS base_imponible_prorrateo, 0.00 AS monto_imp_prorrateo , 
			0.00 AS monto_igtf --IGTF 

		FROM saOrdenPago AS OP
            INNER JOIN saOrdenPagoReng AS OPR ON OPR.ord_num = OP.ord_num
            INNER JOIN saBeneficiario AS BE ON BE.cod_ben = OP.cod_ben
		where OP.fecha between @sCo_fecha_d and @sCo_fecha_h and @bIncluirOrden = 'SI'
		--Sit.# 9986 ZPEREZ
		AND  ( @cCo_Sucursal IS NULL
					  OR OP.co_sucu_in = @cCo_Sucursal)
		--!Sit.# 9986 ZPEREZ
		group by  OP.ord_num, OP.fecha, OP.fe_us_in, OP.cod_ben, BE.ben_des,BE.rif,OP.tasa ,
			case when BE.rif ='' then 'SR' 		
			-->>JN 20200212 Sit# 97826 
			--else case when BE.tipo_per='2' then 'NR' else 
			--case when BE.tipo_per='4' then 'ND' else space(2) end end end, 
			ELSE CASE WHEN  BE.rif IS NULL THEN 'SR'
			ELSE CASE WHEN BE.tipo_per = '1' THEN 'PNR'
            ELSE CASE WHEN BE.tipo_per = '2' THEN 'PNNR'
			ELSE CASE WHEN BE.tipo_per = '3' THEN 'PJD'
            ELSE CASE WHEN BE.tipo_per = '4' THEN 'PJND'
			ELSE CASE WHEN BE.tipo_per = '6' THEN 'TN'
            ELSE SPACE(4)
			   END
				END
                 END
                  END
                   END
				    END
					 END,
		--<<JN 20200212 Sit# 97826
			OP.co_sucu_in, OP.co_mone, OP.anulado, OP.fecha, OPR.tipo_imp) A
		ORDER BY fecha_emis, fe_us_in, fec_reg, co_tipo_doc, nro_doc, tasa	         

        OPEN Tempdocs
        FETCH Tempdocs INTO @nro_doc, @co_tipo_doc, @fecha_emis, @fe_us_in, @total_neto, @co_prov, @prov_des, @r, @nac, @tipo_prov,
            @co_sucu_in, @nro_orig, @doc_orig, @aut, @co_mone, @nro_fact, @n_control, @anulado, @fec_reg, @base_imp,
            @tipo_imp, @tasa, @monto_imp, @doc_afec, @compras_exentas, @base_imponible, @monto_ret_imp,
            @monto_ret_imp_tercero, @num_comprobante, @fec_comprobante
            /***************************************/
			, @base_imponible_scf, @monto_imp_scf,
			/**Factura al Tesoro Nacional**/
			@FTNac_fechaEmis, @num_plan_impor, @num_exp_impor, @valorMercancia, @der_impor, @tasa_regimenAplic, @otrosGravables, @otrosExentos, @compras_exentas_FCI,
			/**************Art.34**************/
			@base_imponible_deducible, @monto_imp_deducible, @base_imponible_prorrateo, @monto_imp_prorrateo , @monto_igtf -- IGTF 


	  
        SET @old_nro_doc = ''
        SET @old_co_tipo_doc = ''
        SET @old_tasa = 0
        SET @lupdate = 0


        WHILE @@fetch_status != -1 
            BEGIN

                SET @lupdate = 0
                SET @exento_Actual = 0
                IF @tasa = 0 
                    BEGIN
                        SET @compras_exentas = @compras_exentas + @base_imp
                        SET @exento_Actual = @base_imp
                        SET @base_imp = CASE WHEN @compras_exentas_FCI IS NULL THEN 0 ELSE @base_imp END
                        SET @monto_imp = 0
	    	--set @exento_Actual = @compras_exentas
                    END
	   
					 IF @old_nro_doc = @nro_doc and @old_co_tipo_doc = @co_tipo_doc -- Si es el mismo doc 
						SET @monto_igtf = 0 

	-- si cambió el documento o es el primero
                IF @old_nro_doc <> @nro_doc
                    OR @old_co_tipo_doc <> @co_tipo_doc 
                    BEGIN
                        SET @old_nro_doc = @nro_doc
                        SET @old_co_tipo_doc = @co_tipo_doc
                        SET @old_tasa = @tasa
                        SET @lupdate = 0
                        SET @exento_Actual = @compras_exentas


						/***Factura al Tesoro Nacional (Compras Exentas de la Factura de Importación)***/
						--Si no es nulo, es una factura de importación
						IF @compras_exentas_FCI IS NOT NULL
						BEGIN
							IF @tasa = 0
								SET @exento_Actual = @compras_exentas_FCI
							ELSE
								SET @exento_Actual = @exento_Actual + @compras_exentas_FCI
							SET @total_neto	=		@total_neto			+	@exento_Actual
						END
						/***Factura al Tesoro Nacional (Compras Exentas de la Factura de Importación)***/

                    END
                ELSE 
                    BEGIN
					
						--Si es nulo, NO es una factura de importación
						IF @compras_exentas_FCI IS NULL
							BEGIN
								--Se repetía el valor de 'total_neto' y me duplicaba el total del campo 'Total de Compras Incluye I.V.A.'
								--No puedo colocar este campo en 0 al tratarse de una Fact. al Tesoro Nacional porque se utiliza para otros cálculos
								SET @total_neto = 0
							END
						
                        IF @old_tasa = 0 
                            BEGIN
                                SET @lupdate = 1
                                SET @old_tasa = @tasa
                            END
                    END

                IF @lupdate = 1 
                    BEGIN
                        IF @tasa <> 0 
							BEGIN
								UPDATE
									@temp
								SET base_imp = base_imp + @base_imp, tipo_imp = @tipo_imp, tasa = @tasa, monto_imp = @monto_imp,
									base_imponible = @base_imponible, monto_ret_imp = @monto_ret_imp,
									num_comprobante = @num_comprobante, fec_comprobante = @fec_comprobante,
									monto_ret_imp_tercero = @monto_ret_imp_tercero
									/*------------------------------------------------*/
									, base_imponible_scf = @base_imponible_scf, monto_imp_scf = @monto_imp_scf
									/*Factura al Tesoro Nacional*/
									, valorMercancia = @valorMercancia,
									/**************Art.34**************/
									base_imponible_deducible = @base_imponible_deducible, monto_imp_deducible = @monto_imp_deducible,
									base_imponible_prorrateo = @base_imponible_prorrateo, monto_imp_prorrateo = @monto_imp_prorrateo --, monto_igtf = @monto_igtf --IGTF 
								WHERE
									nro_doc = @nro_doc
									AND co_tipo_doc = @co_tipo_doc

								--SI ES UNA FACTURA DE IMPORTACIÓN:
								--QUE ME ACUMULE EL TOTAL_NETO YA QUE ESTE CAMPO NO VIENE TOTALIZADO AL TRATARSE DE UNA FACTURA DE IMPORTACIÓN
								IF @compras_exentas_FCI IS NOT NULL
									BEGIN
										UPDATE
											@temp
										SET total_neto = total_neto + @total_neto
										WHERE
											nro_doc = @nro_doc
											AND co_tipo_doc = @co_tipo_doc
									END

							END

                        ELSE
							--Si no es una factura de importación, actualiza las compras exentas
							IF @compras_exentas_FCI IS NULL
							BEGIN
								UPDATE
									@temp
								SET compras_exentas = compras_exentas + @exento_Actual
								WHERE
									nro_doc = @nro_doc
									AND co_tipo_doc = @co_tipo_doc
							END
                    END
                ELSE 
                    BEGIN
                        INSERT  INTO @temp
                                ( nro_doc, co_tipo_doc, fecha_emis, fe_us_in, total_neto, co_prov, prov_des, r, tipo_prov, contrib,
                                  nac, co_sucu_in, nro_orig, doc_orig, aut, co_mone, nro_fact, n_control, anulado,
                                  fec_reg, base_imp, tipo_imp, tasa, monto_imp, doc_afec, compras_exentas,
                                  base_imponible, monto_ret_imp, monto_ret_imp_tercero, num_comprobante, fec_comprobante
                                   /***************************************/ 
								  , base_imponible_scf, monto_imp_scf,
								  /**Factura al Tesoro Nacional**/
								  FTNac_fechaEmis, num_plan_impor, num_exp_impor, valorMercancia, der_impor, tasa_regimenAplic, otrosGravables, otrosExentos,
								  /**************Art.34**************/
								  base_imponible_deducible, monto_imp_deducible, base_imponible_prorrateo, monto_imp_prorrateo , 
								  monto_igtf --IGTF
								  
								  )
                                SELECT
                                    @nro_doc, @co_tipo_doc, @fecha_emis, @fe_us_in, @total_neto, @co_prov, @prov_des, @r,
                                    @tipo_prov, @contrib, @nac, @co_sucu_in, @nro_orig, @doc_orig, @aut, @co_mone,
                                    @nro_fact, @n_control, @anulado, @fec_reg, @base_imp, @tipo_imp, @tasa, @monto_imp,
                                    @doc_afec, @exento_Actual, @base_imponible, @monto_ret_imp, @monto_ret_imp_tercero,
                                    @num_comprobante, @fec_comprobante
                                    /***************************************/
									, @base_imponible_scf, @monto_imp_scf,
									/**Factura al Tesoro Nacional**/
									@FTNac_fechaEmis, @num_plan_impor, @num_exp_impor, @valorMercancia, @der_impor, @tasa_regimenAplic, @otrosGravables, @otrosExentos,
									/**************Art.34**************/
									@base_imponible_deducible, @monto_imp_deducible, @base_imponible_prorrateo, @monto_imp_prorrateo ,
									@monto_igtf -- IGTF


                    END
	  
                FETCH Tempdocs INTO @nro_doc, @co_tipo_doc, @fecha_emis, @fe_us_in, @total_neto, @co_prov, @prov_des, @r, @nac,
                    @tipo_prov, @co_sucu_in, @nro_orig, @doc_orig, @aut, @co_mone, @nro_fact, @n_control, @anulado,
                    @fec_reg, @base_imp, @tipo_imp, @tasa, @monto_imp, @doc_afec, @compras_exentas, @base_imponible,
                    @monto_ret_imp, @monto_ret_imp_tercero, @num_comprobante, @fec_comprobante
                    /***************************************/
                    ,@base_imponible_scf, @monto_imp_scf,
					/**Factura al Tesoro Nacional**/
					@FTNac_fechaEmis, @num_plan_impor, @num_exp_impor, @valorMercancia, @der_impor, @tasa_regimenAplic, @otrosGravables, @otrosExentos, @compras_exentas_FCI,
					/**************Art.34**************/
					@base_imponible_deducible, @monto_imp_deducible, @base_imponible_prorrateo, @monto_imp_prorrateo , 
					@monto_igtf -- IGTF
            END
        DEALLOCATE Tempdocs	


/*********************************Validar si existen Facturas al Tesoro Nacional*********************************/
		
		IF (@bImprimirColumnImport = 'NO' AND EXISTS (SELECT TM.* FROM @Temp TM
														INNER JOIN saDatosDeImportacion DI ON DI.fact_num = TM.nro_doc
														--	Sit.# 16965
														--	2017-05-11
														--	Si el documento se encuentra asociado a una Factura al Tesoro Nacional anulada, no debo considerar sus datos.
														INNER JOIN saFacturaCompraReng FTN ON FTN.rowguid = DI.rowguid_factura_renglon
														INNER JOIN saDocumentoCompra DOC ON DOC.nro_doc = FTN.doc_num
																					WHERE(@cCo_Sucursal IS NULL
																							OR TM.co_sucu_in = @cCo_Sucursal)
																							AND DOC.anulado = 0
																						)) 
		BEGIN
			RAISERROR ('Existen facturas de importación y no se están mostrando las columnas correspondientes.',16,1)
			RETURN
		END

/*********************************Validar si existen Facturas al Tesoro Nacional*********************************/ 

/*---------------------------------------------------------------------------------------------------------
Esta parte es para obtener el cuadro resumen y darle formato a cada registro del libro de compras
---------------------------------------------------------------------------------------------------------*/

        DECLARE @tempfinal TABLE
            (
              [descrip] [char](100) ,
              [nac] [bit] ,
              [tasa] [decimal](18, 2) ,
              [base_imp] [decimal](18, 2) ,
              [monto_imp] [decimal](18, 2) ,
              [retenido] [decimal](18, 2) ,
              [retenidoter] [decimal](18, 2)
            )
     
        DECLARE
            @ldescrip VARCHAR(100) ,
            @retenido DECIMAL(18, 2) ,
            @retenidoter DECIMAL(18, 2) ,
            @cont INT ,
            @base_imp_tot DECIMAL(18, 2) ,
            @monto_imp_tot DECIMAL(18, 2) ,
            @retenido_tot DECIMAL(18, 2) ,
            @retenido_ter_tot DECIMAL(18, 2)

			/**************Factura al Tesoro Nacional**************/

			--Para tomar en cuenta las compras exentas de las facturas de importación en el cuadro resumen
			SELECT @compras_exentas_FCI = SUM (CE.ComprasExentas_FCI) FROM
			(
				SELECT
					nro_doc, co_tipo_doc, ISNULL(ComprasExentas_FCI, 0) AS ComprasExentas_FCI
				FROM
					DocumentosLibrocompras2(@sCo_fecha_d, @sCo_fecha_h)
				--Sit.# 9986 ZPEREZ
					WHERE ( @cCo_Sucursal IS NULL
						  OR co_sucu_in = @cCo_Sucursal)
				--!Sit.# 9986 ZPEREZ
				GROUP BY
					nro_doc, co_tipo_doc, ComprasExentas_FCI
			) CE

			/**************Factura al Tesoro Nacional**************/
     
        DECLARE Tempdocs CURSOR
        FOR
            ( SELECT
                nro_doc, co_tipo_doc, total_neto, nac, anulado, base_imp, tasa, monto_imp, compras_exentas,
                monto_ret_imp, monto_ret_imp_tercero
              FROM
                DocumentosLibrocompras2(@sCo_fecha_d, @sCo_fecha_h)
              WHERE
                anulado = 0
				--Sit.# 9986 ZPEREZ
				AND ( @cCo_Sucursal IS NULL
						OR co_sucu_in = @cCo_Sucursal)
				--!Sit.# 9986 ZPEREZ
            )
            ORDER BY
            fecha_emis, n_control, tasa


        SET @old_nro_doc = ''
        SET @old_co_tipo_doc = ''

-- siempre voy a tener los registros (nacional y extranjero) a tasa 0 %
        INSERT  INTO @tempfinal
                ( descrip, nac, tasa, base_imp, monto_imp, retenido, retenidoter )
        VALUES
                ( 'Total Compras de Importación', 0, 0, @compras_exentas_FCI, 0, 0, 0 )
        INSERT  INTO @tempfinal
                ( descrip, nac, tasa, base_imp, monto_imp, retenido, retenidoter )
        VALUES
                ( 'Total Compras Internas No Gravadas', 1, 0, 0, 0, 0, 0 )

        OPEN Tempdocs
        FETCH Tempdocs INTO @nro_doc, @co_tipo_doc, @total_neto, @nac, @anulado, @base_imp, @tasa, @monto_imp,
            @compras_exentas, @retenido, @retenidoter
	  
        WHILE @@fetch_status != -1 
            BEGIN
                IF NOT EXISTS ( SELECT
                                    tasa
                                FROM
                                    @tempfinal
                                WHERE
                                    tasa = @tasa
                                    AND nac = @nac ) 
                    BEGIN
                        SET @ldescrip = 'Total Compras ' + CASE WHEN @nac = 1 THEN 'Internas'
                                                                ELSE 'Importación'
                                                           END + ' afectadas sólo alícuota '
                        INSERT  INTO @tempfinal
                                ( descrip, nac, tasa, base_imp, monto_imp, retenido, retenidoter )
                        VALUES
                                ( @ldescrip + CAST(@tasa AS VARCHAR), @nac, @tasa, 0, 0, 0, 0 )
                    END
	
	-- si cambió el documento o es el primero (para saber si se suma los campos OTROS
                IF @old_nro_doc <> @nro_doc
                    OR @old_co_tipo_doc <> @co_tipo_doc 
                    IF @compras_exentas != 0 
                        BEGIN
                            SET @old_nro_doc = @nro_doc
                            SET @old_co_tipo_doc = @co_tipo_doc
                            UPDATE
                                @tempfinal
                            SET base_imp = base_imp + @compras_exentas
                            WHERE
                                tasa = 0
                                AND nac = @nac
                        END			
	
	
                UPDATE
                    @tempfinal
                SET base_imp = base_imp + @base_imp, monto_imp = monto_imp + @monto_imp, retenido = retenido + @retenido,
                    retenidoter = retenidoter + @retenidoter
                WHERE
                    tasa = @tasa
                    AND nac = @nac
							
                FETCH Tempdocs INTO @nro_doc, @co_tipo_doc, @total_neto, @nac, @anulado, @base_imp, @tasa, @monto_imp,
                    @compras_exentas, @retenido, @retenidoter
            END
        DEALLOCATE Tempdocs	

        SET @cont = 0
        SET @base_imp_tot = 0
        SET @monto_imp_tot = 0
        SET @retenido_tot = 0
        SET @retenido_ter_tot = 0

        DECLARE TempTasas CURSOR
        FOR
            ( SELECT
                descrip, nac, tasa, base_imp, monto_imp, retenido, retenidoter
              FROM
                @TempFinal
              WHERE
                base_imp != 0
                OR retenido != 0
            )
            ORDER BY
            nac, tasa
        OPEN TempTasas
        FETCH TempTasas INTO @ldescrip, @nac, @tasa, @base_imp, @monto_imp, @retenido, @retenidoter
	  
        WHILE @@fetch_status != -1 
            BEGIN
                SET @cont = @cont + 1
                SET @base_imp_tot = @base_imp_tot + @base_imp
                SET @monto_imp_tot = @monto_imp_tot + @monto_imp
                SET @retenido_tot = @retenido_tot + @retenido
                SET @retenido_ter_tot = @retenido_ter_tot + @retenidoter
	
                IF @cont = 1 
                    UPDATE
                        @temp
                    SET descrip1 = @ldescrip, base_imp1 = @base_imp, monto_imp1 = @monto_imp, retenido1 = @retenido,
                        retenidoter1 = @retenidoter, descrip2 = 'Totales: ', base_imp2 = @base_imp_tot,
                        monto_imp2 = @monto_imp_tot, retenido2 = @retenido_tot, retenidoter2 = @retenido_ter_tot
                IF @cont = 2 
                    UPDATE
                        @temp
                    SET descrip2 = @ldescrip, base_imp2 = @base_imp, monto_imp2 = @monto_imp, retenido2 = @retenido,
                        retenidoter2 = @retenidoter, descrip3 = 'Totales: ', base_imp3 = @base_imp_tot,
                        monto_imp3 = @monto_imp_tot, retenido3 = @retenido_tot, retenidoter3 = @retenido_ter_tot
                IF @cont = 3 
                    UPDATE
                        @temp
                    SET descrip3 = @ldescrip, base_imp3 = @base_imp, monto_imp3 = @monto_imp, retenido3 = @retenido,
                        retenidoter3 = @retenidoter, descrip4 = 'Totales: ', base_imp4 = @base_imp_tot,
                        monto_imp4 = @monto_imp_tot, retenido4 = @retenido_tot, retenidoter4 = @retenido_ter_tot
                IF @cont = 4 
                    UPDATE
                        @temp
                    SET descrip4 = @ldescrip, base_imp4 = @base_imp, monto_imp4 = @monto_imp, retenido4 = @retenido,
                        retenidoter4 = @retenidoter, descrip5 = 'Totales: ', base_imp5 = @base_imp_tot,
                        monto_imp5 = @monto_imp_tot, retenido5 = @retenido_tot, retenidoter5 = @retenido_ter_tot
                IF @cont = 5 
                    UPDATE
                        @temp
                    SET descrip5 = @ldescrip, base_imp5 = @base_imp, monto_imp5 = @monto_imp, retenido5 = @retenido,
                        retenidoter5 = @retenidoter, descrip6 = 'Totales: ', base_imp6 = @base_imp_tot,
                        monto_imp6 = @monto_imp_tot, retenido6 = @retenido_tot, retenidoter6 = @retenido_ter_tot
                IF @cont = 6 
                    UPDATE
                        @temp
                    SET descrip6 = @ldescrip, base_imp6 = @base_imp, monto_imp6 = @monto_imp, retenido6 = @retenido,
                        retenidoter6 = @retenidoter, descrip7 = 'Totales: ', base_imp7 = @base_imp_tot,
                        monto_imp7 = @monto_imp_tot, retenido7 = @retenido_tot, retenidoter7 = @retenido_ter_tot
                IF @cont = 7 
                    UPDATE
                        @temp
                    SET descrip7 = @ldescrip, base_imp7 = @base_imp, monto_imp7 = @monto_imp, retenido7 = @retenido,
                        retenidoter7 = @retenidoter, descrip8 = 'Totales: ', base_imp8 = @base_imp_tot,
                        monto_imp8 = @monto_imp_tot, retenido8 = @retenido_tot, retenidoter8 = @retenido_ter_tot
                IF @cont = 8 
                    UPDATE
                        @temp
                    SET descrip8 = @ldescrip, base_imp8 = @base_imp, monto_imp8 = @monto_imp, retenido8 = @retenido,
                        retenidoter8 = @retenidoter, descrip9 = 'Totales: ', base_imp9 = @base_imp_tot,
                        monto_imp9 = @monto_imp_tot, retenido9 = @retenido_tot, retenidoter9 = @retenido_ter_tot
                IF @cont = 9 
                    UPDATE
                        @temp
                    SET descrip9 = @ldescrip, base_imp9 = @base_imp, monto_imp9 = @monto_imp, retenido9 = @retenido,
                        retenidoter9 = @retenidoter, descrip10 = 'Totales: ', base_imp10 = @base_imp_tot,
                        monto_imp10 = @monto_imp_tot, retenido10 = @retenido_tot, retenidoter10 = @retenido_ter_tot
                IF @cont = 10 
                    UPDATE
                        @temp
                    SET descrip10 = @ldescrip, base_imp10 = @base_imp, monto_imp10 = @monto_imp, retenido10 = @retenido,
                        retenidoter10 = @retenidoter, descrip11 = 'Totales: ', base_imp11 = @base_imp_tot,
                        monto_imp11 = @monto_imp_tot, retenido11 = @retenido_tot, retenidoter11 = @retenido_ter_tot
                IF @cont = 11 
                    UPDATE
                        @temp
                    SET descrip11 = @ldescrip, base_imp11 = @base_imp, monto_imp11 = @monto_imp, retenido11 = @retenido,
                        retenidoter11 = @retenidoter, descrip12 = 'Totales: ', base_imp12 = @base_imp_tot,
                        monto_imp12 = @monto_imp_tot, retenido12 = @retenido_tot, retenidoter12 = @retenido_ter_tot
                IF @cont = 12 
                    UPDATE
                        @temp
                    SET descrip12 = @ldescrip, base_imp12 = @base_imp, monto_imp12 = @monto_imp, retenido12 = @retenido,
                        retenidoter12 = @retenidoter, descrip13 = 'Totales: ', base_imp13 = @base_imp_tot,
                        monto_imp13 = @monto_imp_tot, retenido13 = @retenido_tot, retenidoter13 = @retenido_ter_tot
				/*-------------------------------------------------------*/
				IF @cont = 13 
                    UPDATE
                        @temp
                    SET descrip13 = @ldescrip, base_imp13 = @base_imp, monto_imp13 = @monto_imp, retenido13 = @retenido,
                        retenidoter13 = @retenidoter, descrip14 = 'Totales: ', base_imp14 = @base_imp_tot,
                        monto_imp14 = @monto_imp_tot, retenido14 = @retenido_tot, retenidoter14 = @retenido_ter_tot	
				IF @cont = 14 
                    UPDATE
                        @temp
                    SET descrip14 = @ldescrip, base_imp14 = @base_imp, monto_imp14 = @monto_imp, retenido14 = @retenido,
                        retenidoter14 = @retenidoter, descrip15 = 'Totales: ', base_imp15 = @base_imp_tot,
                        monto_imp15 = @monto_imp_tot, retenido15 = @retenido_tot, retenidoter15 = @retenido_ter_tot
                        	 
                FETCH TempTasas INTO @ldescrip, @nac, @tasa, @base_imp, @monto_imp, @retenido, @retenidoter
            END
        DEALLOCATE TempTasas
        
        /*-----------------------------------------------------------------------------------------------*/
        /*Modificación para la informacion de compras con derecho a crédito fiscal no deducible Art(33) */
        /*----------------------------------------------------------------------------------------------*/	
        
        DECLARE @tempCCFND TABLE
            (
              [descrip] [char](100) ,
              [nac] [bit] ,
              [tasa] [decimal](18, 2) ,
              [base_imp] [decimal](18, 2) ,
              [monto_imp] [decimal](18, 2) ,
              [base_imponible_scf] [DECIMAL] (18,2),
			  [monto_imp_scf] [decimal](18,2)
            )
     
        DECLARE
            @ldescripCCFND VARCHAR(100) ,
            @contCCFND INT ,
            @base_imp_totCCFND DECIMAL(18, 2) ,
            @monto_imp_totCCFND DECIMAL(18, 2) ,
			@base_imp_tot_scfCCFND DECIMAL (18,2),
			@monto_imp_totscfCCFND DECIMAL(18,2)
			
        DECLARE Tempdocs1 CURSOR
        FOR
            ( SELECT
                nro_doc, co_tipo_doc, total_neto, nac, anulado, base_imp, tasa, monto_imp, compras_exentas, base_imponible_scf, monto_imp_scf
              FROM
                DocumentosLibrocompras2(@sCo_fecha_d, @sCo_fecha_h)
              WHERE
                anulado = 0 AND tasa <> 0 AND @bImprimirColumnArt33 = 'SI'
				--Sit.# 9986 ZPEREZ
					AND ( @cCo_Sucursal IS NULL
							OR co_sucu_in = @cCo_Sucursal)
				--!Sit.# 9986 ZPEREZ
            )
            ORDER BY
            fecha_emis, n_control, nac, tasa
 
			SET @old_nro_doc = ''
			SET @old_co_tipo_doc = ''

			INSERT  INTO @tempCCFND
					( descrip, nac, tasa, base_imp, monto_imp, base_imponible_scf, monto_imp_scf )
			VALUES
					( 'Compras con Crédito Fiscal No Deducible (Art.33)', 0, 0, NULL, NULL, NULL, NULL )
			
	        OPEN Tempdocs1
		    FETCH Tempdocs1 INTO @nro_doc, @co_tipo_doc, @total_neto, @nac, @anulado, @base_imp, @tasa, @monto_imp,
			    @compras_exentas, @base_imponible_scf, @monto_imp_scf 
	  
			WHILE @@fetch_status != -1 
				BEGIN
					IF NOT EXISTS ( SELECT
                                    tasa
                                FROM
                                    @tempCCFND
                                WHERE
                                    tasa = @tasa
									AND nac = @nac                                   
								) 
					BEGIN

						SET @ldescrip = 'Total Compras ' + CASE WHEN @nac = 1 THEN 'Internas'
                                                                ELSE 'Importación'
                                                           END + ' afectadas sólo alícuota '


                        INSERT  INTO @tempCCFND
                                ( descrip, nac, tasa, base_imp, monto_imp, base_imponible_scf, monto_imp_scf )
                        VALUES
                                ( @ldescrip + CAST(@tasa AS VARCHAR), @nac, @tasa, 0, 0, 0, 0 )
					END
					
					UPDATE
						@tempCCFND
					SET base_imp = base_imp + @base_imp, 
						monto_imp = monto_imp + @monto_imp , 
						base_imponible_scf = base_imponible_scf + @base_imponible_scf, monto_imp_scf = monto_imp_scf + @monto_imp_scf
					WHERE
						tasa = @tasa
						AND nac = @nac
					
					FETCH Tempdocs1 INTO @nro_doc, @co_tipo_doc, @total_neto, @nac, @anulado, @base_imp, @tasa, @monto_imp,
						@compras_exentas, @base_imponible_scf, @monto_imp_scf 
				END
				
				DECLARE @tasaTasasTotal DECIMAL(18, 2),  
						@monto_imp_scfTasasTotal DECIMAL(18,2)
				
				IF NOT EXISTS(SELECT SUM(monto_imp_scf) FROM @tempCCFND HAVING SUM(monto_imp_scf) > 0) 
				BEGIN
					DELETE FROM @tempCCFND
				END
				
        DEALLOCATE Tempdocs1
		
		DELETE FROM @tempCCFND WHERE monto_imp_scf = 0	

        SET @contCCFND = @cont + 2--8
        SET @base_imp_totCCFND = 0
        SET @monto_imp_totCCFND = 0
        SET @base_imp_tot_scfCCFND = 0
		SET @monto_imp_totscfCCFND = 0

        DECLARE TempTasas1 CURSOR
        FOR
            ( SELECT
                descrip, nac, tasa, base_imp, monto_imp, base_imponible_scf, monto_imp_scf
              FROM
                @tempCCFND
             )
            ORDER BY
            nac, tasa
			
        OPEN TempTasas1
        FETCH TempTasas1 INTO @ldescrip, @nac, @tasa, @base_imp, @monto_imp, @base_imponible_scf, @monto_imp_scf
	  
        WHILE @@fetch_status != -1 
            BEGIN

                SET @contCCFND =  @contCCFND + 1
                SET @base_imp_totCCFND = @base_imp_totCCFND + ISNULL(@base_imp,0)
                SET @monto_imp_totCCFND = @monto_imp_totCCFND + ISNULL(@monto_imp,0)
                SET @base_imp_tot_scfCCFND = @base_imp_tot_scfCCFND + ISNULL(@base_imponible_scf,0)
				SET @monto_imp_totscfCCFND = @monto_imp_totscfCCFND + ISNULL(@monto_imp_scf,0)
                
				IF @contCCFND = 4
                    UPDATE
                        @temp
                    SET descrip4 = @ldescrip, base_imp4 = @base_imponible_scf, monto_imp4 = @monto_imp_scf, 
						descrip5 = 'Totales: ', base_imp5 = @base_imp_tot_scfCCFND, monto_imp5 = @monto_imp_totscfCCFND 
				IF @contCCFND = 5
                    UPDATE
                        @temp
                    SET descrip5 = @ldescrip, base_imp5 = @base_imponible_scf, monto_imp5 = @monto_imp_scf, 
						descrip6 = 'Totales: ', base_imp6 = @base_imp_tot_scfCCFND, monto_imp6 = @monto_imp_totscfCCFND 
				IF @contCCFND = 6
                    UPDATE
                        @temp
                    SET descrip6 = @ldescrip, base_imp6 = @base_imponible_scf, monto_imp6 = @monto_imp_scf, 
						descrip7 = 'Totales: ', base_imp7 = @base_imp_tot_scfCCFND, monto_imp7 = @monto_imp_totscfCCFND 
				IF @contCCFND = 7
                    UPDATE
                        @temp
                    SET descrip7 = @ldescrip, base_imp7 = @base_imponible_scf, monto_imp7 = @monto_imp_scf, 
						descrip8 = 'Totales: ', base_imp8 = @base_imp_tot_scfCCFND, monto_imp8 = @monto_imp_totscfCCFND 
				IF @contCCFND = 8
                    UPDATE
                        @temp
                    SET descrip8 = @ldescrip, base_imp8 = @base_imponible_scf, monto_imp8 = @monto_imp_scf, 
						descrip9 = 'Totales: ', base_imp9 = @base_imp_tot_scfCCFND, monto_imp9 = @monto_imp_totscfCCFND 

                IF @contCCFND = 9
                    UPDATE
                        @temp
                    SET descrip9 = @ldescrip, base_imp9 = @base_imponible_scf, monto_imp9 = @monto_imp_scf, 
						descrip10 = 'Totales: ', base_imp10 = @base_imp_tot_scfCCFND, monto_imp10 = @monto_imp_totscfCCFND 
						--,descrip11 = 'Total Crédito Fiscal Deducible: ', monto_imp11 = (@monto_imp_totCCFND +  @monto_imp_totscfCCFND)
						--No es necesario realizar esta totalización con la nueva funcionalidad Art.34
                IF @contCCFND = 10 
                    UPDATE
                        @temp
                    SET descrip10 = @ldescrip, base_imp10 = @base_imponible_scf, monto_imp10 = @monto_imp_scf, 
						descrip11 = 'Totales: ', base_imp11 = @base_imp_tot_scfCCFND, monto_imp11 = @monto_imp_totscfCCFND
						--,descrip12 = 'Total Crédito Fiscal Deducible: ', monto_imp12 = (@monto_imp_totCCFND +  @monto_imp_totscfCCFND)
						--No es necesario realizar esta totalización con la nueva funcionalidad Art.34
                IF @contCCFND = 11 
                    UPDATE
                        @temp
                    SET descrip11 = @ldescrip, base_imp11 = @base_imponible_scf, monto_imp11 = @monto_imp_scf, 
						descrip12 = 'Totales: ', base_imp12 = @base_imp_tot_scfCCFND, monto_imp12 = @monto_imp_totscfCCFND
						--,descrip13 = 'Total Crédito Fiscal Deducible: ', monto_imp13 = (@monto_imp_totCCFND +  @monto_imp_totscfCCFND)
						--No es necesario realizar esta totalización con la nueva funcionalidad Art.34
						/*-------------------------------------------------------*/
				IF @contCCFND = 12 
                    UPDATE
                        @temp
                    SET descrip12 = @ldescrip, base_imp12 = @base_imponible_scf, monto_imp12 = @monto_imp_scf, 
						descrip13 = 'Totales: ', base_imp13 = @base_imp_tot_scfCCFND, monto_imp13 = @monto_imp_totscfCCFND 
						--descrip14 = 'Total Crédito Fiscal Deducible: ', monto_imp14 = (@monto_imp_totCCFND +  @monto_imp_totscfCCFND)
						--No es necesario realizar esta totalización con la nueva funcionalidad Art.34
				IF @contCCFND = 13 
                    UPDATE
                        @temp
                    SET descrip13 = @ldescrip, base_imp13 = @base_imponible_scf, monto_imp13 = @monto_imp_scf, 
						descrip14 = 'Totales: ', base_imp14 = @base_imp_tot_scfCCFND, monto_imp14 = @monto_imp_totscfCCFND
						--descrip15 = 'Total Crédito Fiscal Deducible: ', monto_imp15 = (@monto_imp_totCCFND +  @monto_imp_totscfCCFND)
						--No es necesario realizar esta totalización con la nueva funcionalidad Art.34
				--Se necesita llegar hasta el campo descrip21 para abastecer un cuadro resumen que muestre todas las tasas
				IF @contCCFND = 14
                    UPDATE
                        @temp
                    SET descrip14 = @ldescrip, base_imp14 = @base_imponible_scf, monto_imp14 = @monto_imp_scf, 
						descrip15 = 'Totales: ', base_imp15 = @base_imp_tot_scfCCFND, monto_imp15 = @monto_imp_totscfCCFND 
				IF @contCCFND = 15
                    UPDATE
                        @temp
                    SET descrip15 = @ldescrip, base_imp15 = @base_imponible_scf, monto_imp15 = @monto_imp_scf, 
						descrip16 = 'Totales: ', base_imp16 = @base_imp_tot_scfCCFND, monto_imp16 = @monto_imp_totscfCCFND 
				IF @contCCFND = 16
                    UPDATE
                        @temp
                    SET descrip16 = @ldescrip, base_imp16 = @base_imponible_scf, monto_imp16 = @monto_imp_scf, 
						descrip17 = 'Totales: ', base_imp17 = @base_imp_tot_scfCCFND, monto_imp17 = @monto_imp_totscfCCFND 
				IF @contCCFND = 17
                    UPDATE
                        @temp
                    SET descrip17 = @ldescrip, base_imp17 = @base_imponible_scf, monto_imp17 = @monto_imp_scf, 
						descrip18 = 'Totales: ', base_imp18 = @base_imp_tot_scfCCFND, monto_imp18 = @monto_imp_totscfCCFND 
				IF @contCCFND = 18
                    UPDATE
                        @temp
                    SET descrip18 = @ldescrip, base_imp18 = @base_imponible_scf, monto_imp18 = @monto_imp_scf, 
						descrip19 = 'Totales: ', base_imp19 = @base_imp_tot_scfCCFND, monto_imp19 = @monto_imp_totscfCCFND 
				IF @contCCFND = 19
                    UPDATE
                        @temp
                    SET descrip19 = @ldescrip, base_imp19 = @base_imponible_scf, monto_imp19 = @monto_imp_scf, 
						descrip20 = 'Totales: ', base_imp20 = @base_imp_tot_scfCCFND, monto_imp20 = @monto_imp_totscfCCFND 
				IF @contCCFND = 20
                    UPDATE
                        @temp
                    SET descrip20 = @ldescrip, base_imp20 = @base_imponible_scf, monto_imp20 = @monto_imp_scf, 
						descrip21 = 'Totales: ', base_imp21 = @base_imp_tot_scfCCFND, monto_imp21 = @monto_imp_totscfCCFND 
				IF @contCCFND = 21
                    UPDATE
                        @temp
                    SET descrip21 = @ldescrip, base_imp21 = @base_imponible_scf, monto_imp21 = @monto_imp_scf, 
						descrip22 = 'Totales: ', base_imp22 = @base_imp_tot_scfCCFND, monto_imp22 = @monto_imp_totscfCCFND 
				IF @contCCFND = 22
                    UPDATE
                        @temp
                    SET descrip22 = @ldescrip, base_imp22 = @base_imponible_scf, monto_imp22 = @monto_imp_scf, 
						descrip23 = 'Totales: ', base_imp23 = @base_imp_tot_scfCCFND, monto_imp23 = @monto_imp_totscfCCFND 
				IF @contCCFND = 23
                    UPDATE
                        @temp
                    SET descrip23 = @ldescrip, base_imp23 = @base_imponible_scf, monto_imp23 = @monto_imp_scf, 
						descrip24 = 'Totales: ', base_imp24 = @base_imp_tot_scfCCFND, monto_imp24 = @monto_imp_totscfCCFND 
				IF @contCCFND = 24
                    UPDATE
                        @temp
                    SET descrip24 = @ldescrip, base_imp24 = @base_imponible_scf, monto_imp24 = @monto_imp_scf, 
						descrip25 = 'Totales: ', base_imp25 = @base_imp_tot_scfCCFND, monto_imp25 = @monto_imp_totscfCCFND 
				IF @contCCFND = 25
                    UPDATE
                        @temp
                    SET descrip25 = @ldescrip, base_imp25 = @base_imponible_scf, monto_imp25 = @monto_imp_scf, 
						descrip26 = 'Totales: ', base_imp26 = @base_imp_tot_scfCCFND, monto_imp26 = @monto_imp_totscfCCFND 
				
                FETCH TempTasas1 INTO @ldescrip, @nac, @tasa, @base_imp, @monto_imp, @base_imponible_scf, @monto_imp_scf
            END
        DEALLOCATE TempTasas1
        
/*----------------------------------------------------------------------------------------------*/

/********************************************CUADRO RESUMEN ART.34********************************************/

/****************COMPRAS CON CRÉDITO FISCAL TOTALMENTE DEDUCIBLE****************/

 DECLARE @temp_CFDeducible TABLE
            (
              [descrip] [char](100) ,
              [nac] [bit] ,
              [tasa] [decimal](18, 2) ,
              [base_imp] [decimal](18, 2) ,
              [monto_imp] [decimal](18, 2) ,
              [base_imponible_CFDeducible] [DECIMAL] (18,2),
			  [monto_imp_CFDeducible] [decimal](18,2)
            )
     
        DECLARE
            @ldescrip_CFDeducible VARCHAR(100) ,
            @cont_CFDeducible INT ,
			@base_imp_CFDeducible DECIMAL (18,2),
			@monto_imp_CFDeducible DECIMAL(18,2)
			
        DECLARE Tempdocs_CFDeducible CURSOR
        FOR
            ( SELECT
                nro_doc, co_tipo_doc, total_neto, nac, anulado, base_imp, tasa, monto_imp, compras_exentas, base_imponible_deducible, monto_imp_deducible
              FROM
                DocumentosLibrocompras2(@sCo_fecha_d, @sCo_fecha_h)
              WHERE
                anulado = 0 AND tasa <> 0 AND @bImprimirColumnArt34 = 'SI' AND monto_imp_deducible <> 0
				--Sit.# 9986 ZPEREZ
					AND ( @cCo_Sucursal IS NULL
						  OR co_sucu_in = @cCo_Sucursal)
				--!Sit.# 9986 ZPEREZ
            )
            ORDER BY
            fecha_emis, n_control, nac, tasa
 
			SET @old_nro_doc = ''
			SET @old_co_tipo_doc = ''

			INSERT  INTO @temp_CFDeducible
					( descrip, nac, tasa, base_imp, monto_imp, base_imponible_CFDeducible, monto_imp_CFDeducible )
			VALUES
					( 'Compras con Crédito Fiscal Totalmente Deducible (Art.34)', 0, 0, NULL, NULL, NULL, NULL )
			
	        OPEN Tempdocs_CFDeducible
		    FETCH Tempdocs_CFDeducible INTO @nro_doc, @co_tipo_doc, @total_neto, @nac, @anulado, @base_imp, @tasa, @monto_imp,
			    @compras_exentas, @base_imponible_deducible, @monto_imp_deducible
	  
			WHILE @@fetch_status != -1 
				BEGIN
					IF NOT EXISTS ( SELECT
                                    tasa
                                FROM
                                    @temp_CFDeducible
                                WHERE
                                    tasa = @tasa
									AND nac = @nac                                   
								) 
					BEGIN
                        SET @ldescrip = 'Total Compras ' + CASE WHEN @nac = 1 THEN 'Internas'
                                                                ELSE 'Importación'
                                                           END + ' afectadas sólo alícuota '

                        INSERT  INTO @temp_CFDeducible
                                ( descrip, nac, tasa, base_imp, monto_imp, base_imponible_CFDeducible, monto_imp_CFDeducible )
                        VALUES
                                ( @ldescrip + CAST(@tasa AS VARCHAR), @nac, @tasa, 0, 0, 0, 0 )
					END
					
					UPDATE
						@temp_CFDeducible
					SET base_imp = base_imp + @base_imp, 
						monto_imp = monto_imp + @monto_imp , 
						base_imponible_CFDeducible = base_imponible_CFDeducible + @base_imponible_deducible, monto_imp_CFDeducible = monto_imp_CFDeducible + @monto_imp_deducible
					WHERE
						tasa = @tasa
						AND nac = @nac
					
					FETCH Tempdocs_CFDeducible INTO @nro_doc, @co_tipo_doc, @total_neto, @nac, @anulado, @base_imp, @tasa, @monto_imp,
						@compras_exentas, @base_imponible_deducible, @monto_imp_deducible
				END
				
				IF NOT EXISTS(SELECT SUM(monto_imp_CFDeducible) FROM @temp_CFDeducible HAVING SUM(monto_imp_CFDeducible) > 0) 
				BEGIN
					DELETE FROM @temp_CFDeducible
				END
				
				
        DEALLOCATE Tempdocs_CFDeducible	

		IF (@contCCFND = @cont + 2)
			--SI NO HAY UN CUADRO RESUMEN DE COMPRAS CON CRÉDITO FISCAL NO DEDUCIBLE (ART.33),
			--NO DEBO AGREGAR MAS LÍNEAS PARA IMPRIMIR EL CUADRO RESUMEN DE COMPRAS CON CRÉDITO FISCAL TOTALMENTE DEDUCIBLE (ART.34)
			SET @cont_CFDeducible = @contCCFND
		ELSE
			SET @cont_CFDeducible = @contCCFND + 2
        
        SET @base_imp_CFDeducible = 0
		SET @monto_imp_CFDeducible = 0

        DECLARE TempTasas_CFDeducible CURSOR
        FOR
            ( SELECT
                descrip, nac, tasa, base_imp, monto_imp, base_imponible_CFDeducible, monto_imp_CFDeducible
              FROM
                @temp_CFDeducible
             )
            ORDER BY
            nac, tasa
			
        OPEN TempTasas_CFDeducible
        FETCH TempTasas_CFDeducible INTO @ldescrip, @nac, @tasa, @base_imp, @monto_imp, @base_imponible_deducible, @monto_imp_deducible
	  
        WHILE @@fetch_status != -1 
            BEGIN

                SET @cont_CFDeducible =  @cont_CFDeducible + 1
                SET @base_imp_CFDeducible = @base_imp_CFDeducible + ISNULL(@base_imponible_deducible,0)
				SET @monto_imp_CFDeducible = @monto_imp_CFDeducible + ISNULL(@monto_imp_deducible,0)
                
				IF @cont_CFDeducible = 4
                    UPDATE
                        @temp
                    SET descrip4 = @ldescrip, base_imp4 = @base_imponible_deducible, monto_imp4 = @monto_imp_deducible, 
						descrip5 = 'Totales: ', base_imp5 = @base_imp_CFDeducible, monto_imp5 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 5
                    UPDATE
                        @temp
                    SET descrip5 = @ldescrip, base_imp5 = @base_imponible_deducible, monto_imp5 = @monto_imp_deducible, 
						descrip6 = 'Totales: ', base_imp6 = @base_imp_CFDeducible, monto_imp6 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 6
                    UPDATE
                        @temp
                    SET descrip6 = @ldescrip, base_imp6 = @base_imponible_deducible, monto_imp6 = @monto_imp_deducible, 
						descrip7 = 'Totales: ', base_imp7 = @base_imp_CFDeducible, monto_imp7 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 7
                    UPDATE
                        @temp
                    SET descrip7 = @ldescrip, base_imp7 = @base_imponible_deducible, monto_imp7 = @monto_imp_deducible, 
						descrip8 = 'Totales: ', base_imp8 = @base_imp_CFDeducible, monto_imp8 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 8
                    UPDATE
                        @temp
                    SET descrip8 = @ldescrip, base_imp8 = @base_imponible_deducible, monto_imp8 = @monto_imp_deducible, 
						descrip9 = 'Totales: ', base_imp9 = @base_imp_CFDeducible, monto_imp9 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 9
                    UPDATE
                        @temp
                    SET descrip9 = @ldescrip, base_imp9 = @base_imponible_deducible, monto_imp9 = @monto_imp_deducible, 
						descrip10 = 'Totales: ', base_imp10 = @base_imp_CFDeducible, monto_imp10 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 10
                    UPDATE
                        @temp
                    SET descrip10 = @ldescrip, base_imp10 = @base_imponible_deducible, monto_imp10 = @monto_imp_deducible, 
						descrip11 = 'Totales: ', base_imp11 = @base_imp_CFDeducible, monto_imp11 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 11
                    UPDATE
                        @temp
                    SET descrip11 = @ldescrip, base_imp11 = @base_imponible_deducible, monto_imp11 = @monto_imp_deducible, 
						descrip12 = 'Totales: ', base_imp12 = @base_imp_CFDeducible, monto_imp12 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 12
                    UPDATE
                        @temp
                    SET descrip12 = @ldescrip, base_imp12 = @base_imponible_deducible, monto_imp12 = @monto_imp_deducible, 
						descrip13 = 'Totales: ', base_imp13 = @base_imp_CFDeducible, monto_imp13 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 13
                    UPDATE
                        @temp
                    SET descrip13 = @ldescrip, base_imp13 = @base_imponible_deducible, monto_imp13 = @monto_imp_deducible, 
						descrip14 = 'Totales: ', base_imp14 = @base_imp_CFDeducible, monto_imp14 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 14
                    UPDATE
                        @temp
                    SET descrip14 = @ldescrip, base_imp14 = @base_imponible_deducible, monto_imp14 = @monto_imp_deducible, 
						descrip15 = 'Totales: ', base_imp15 = @base_imp_CFDeducible, monto_imp15 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 15
                    UPDATE
                        @temp
                    SET descrip15 = @ldescrip, base_imp15 = @base_imponible_deducible, monto_imp15 = @monto_imp_deducible, 
						descrip16 = 'Totales: ', base_imp16 = @base_imp_CFDeducible, monto_imp16 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 16
                    UPDATE
                        @temp
                    SET descrip16 = @ldescrip, base_imp16 = @base_imponible_deducible, monto_imp16 = @monto_imp_deducible, 
						descrip17 = 'Totales: ', base_imp17 = @base_imp_CFDeducible, monto_imp17 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 17
                    UPDATE
                        @temp
                    SET descrip17 = @ldescrip, base_imp17 = @base_imponible_deducible, monto_imp17 = @monto_imp_deducible, 
						descrip18 = 'Totales: ', base_imp18 = @base_imp_CFDeducible, monto_imp18 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 18
                    UPDATE
                        @temp
                    SET descrip18 = @ldescrip, base_imp18 = @base_imponible_deducible, monto_imp18 = @monto_imp_deducible, 
						descrip19 = 'Totales: ', base_imp19 = @base_imp_CFDeducible, monto_imp19 = @monto_imp_CFDeducible
				--Se necesita llegar hasta el campo descrip29 para abastecer un cuadro resumen que muestre todas las tasas 
				IF @cont_CFDeducible = 19
                    UPDATE
                        @temp
                    SET descrip19 = @ldescrip, base_imp19 = @base_imponible_deducible, monto_imp19 = @monto_imp_deducible, 
						descrip20 = 'Totales: ', base_imp20 = @base_imp_CFDeducible, monto_imp20 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 20
                    UPDATE
                        @temp
                    SET descrip20 = @ldescrip, base_imp20 = @base_imponible_deducible, monto_imp20 = @monto_imp_deducible, 
						descrip21 = 'Totales: ', base_imp21 = @base_imp_CFDeducible, monto_imp21 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 21
                    UPDATE
                        @temp
                    SET descrip21 = @ldescrip, base_imp21 = @base_imponible_deducible, monto_imp21 = @monto_imp_deducible, 
						descrip22 = 'Totales: ', base_imp22 = @base_imp_CFDeducible, monto_imp22 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 22
                    UPDATE
                        @temp
                    SET descrip22 = @ldescrip, base_imp22 = @base_imponible_deducible, monto_imp22 = @monto_imp_deducible, 
						descrip23 = 'Totales: ', base_imp23 = @base_imp_CFDeducible, monto_imp23 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 23
                    UPDATE
                        @temp
                    SET descrip23 = @ldescrip, base_imp23 = @base_imponible_deducible, monto_imp23 = @monto_imp_deducible, 
						descrip24 = 'Totales: ', base_imp24 = @base_imp_CFDeducible, monto_imp24 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 24
                    UPDATE
                        @temp
                    SET descrip24 = @ldescrip, base_imp24 = @base_imponible_deducible, monto_imp24 = @monto_imp_deducible, 
						descrip25 = 'Totales: ', base_imp25 = @base_imp_CFDeducible, monto_imp25 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 25
                    UPDATE
                        @temp
                    SET descrip25 = @ldescrip, base_imp25 = @base_imponible_deducible, monto_imp25 = @monto_imp_deducible, 
						descrip26 = 'Totales: ', base_imp26 = @base_imp_CFDeducible, monto_imp26 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 26
                    UPDATE
                        @temp
                    SET descrip26 = @ldescrip, base_imp26 = @base_imponible_deducible, monto_imp26 = @monto_imp_deducible, 
						descrip27 = 'Totales: ', base_imp27 = @base_imp_CFDeducible, monto_imp27 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 27
                    UPDATE
                        @temp
                    SET descrip27 = @ldescrip, base_imp27 = @base_imponible_deducible, monto_imp27 = @monto_imp_deducible, 
						descrip28 = 'Totales: ', base_imp28 = @base_imp_CFDeducible, monto_imp28 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 28
                    UPDATE
                        @temp
                    SET descrip28 = @ldescrip, base_imp28 = @base_imponible_deducible, monto_imp28 = @monto_imp_deducible, 
						descrip29 = 'Totales: ', base_imp29 = @base_imp_CFDeducible, monto_imp29 = @monto_imp_CFDeducible
				IF @cont_CFDeducible = 29
                    UPDATE
                        @temp
                    SET descrip29 = @ldescrip, base_imp29 = @base_imponible_deducible, monto_imp29 = @monto_imp_deducible, 
						descrip30 = 'Totales: ', base_imp30 = @base_imp_CFDeducible, monto_imp30 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 30
                    UPDATE
                        @temp
                    SET descrip30 = @ldescrip, base_imp30 = @base_imponible_deducible, monto_imp30 = @monto_imp_deducible, 
						descrip31 = 'Totales: ', base_imp31 = @base_imp_CFDeducible, monto_imp31 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 31
                    UPDATE
                        @temp
                    SET descrip31 = @ldescrip, base_imp31 = @base_imponible_deducible, monto_imp31 = @monto_imp_deducible, 
						descrip32 = 'Totales: ', base_imp32 = @base_imp_CFDeducible, monto_imp32 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 32
                    UPDATE
                        @temp
                    SET descrip32 = @ldescrip, base_imp32 = @base_imponible_deducible, monto_imp32 = @monto_imp_deducible, 
						descrip33 = 'Totales: ', base_imp33 = @base_imp_CFDeducible, monto_imp33 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 33
                    UPDATE
                        @temp
                    SET descrip33 = @ldescrip, base_imp33 = @base_imponible_deducible, monto_imp33 = @monto_imp_deducible, 
						descrip34 = 'Totales: ', base_imp34 = @base_imp_CFDeducible, monto_imp34 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 34
                    UPDATE
                        @temp
                    SET descrip34 = @ldescrip, base_imp34 = @base_imponible_deducible, monto_imp34 = @monto_imp_deducible, 
						descrip35 = 'Totales: ', base_imp35 = @base_imp_CFDeducible, monto_imp35 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 35
                    UPDATE
                        @temp
                    SET descrip35 = @ldescrip, base_imp35 = @base_imponible_deducible, monto_imp35 = @monto_imp_deducible, 
						descrip36 = 'Totales: ', base_imp36 = @base_imp_CFDeducible, monto_imp36 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 36
                    UPDATE
                        @temp
                    SET descrip36 = @ldescrip, base_imp36 = @base_imponible_deducible, monto_imp36 = @monto_imp_deducible, 
						descrip37 = 'Totales: ', base_imp37 = @base_imp_CFDeducible, monto_imp37 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 37
                    UPDATE
                        @temp
                    SET descrip37 = @ldescrip, base_imp37 = @base_imponible_deducible, monto_imp37 = @monto_imp_deducible, 
						descrip38 = 'Totales: ', base_imp38 = @base_imp_CFDeducible, monto_imp38 = @monto_imp_CFDeducible 
				IF @cont_CFDeducible = 38
                    UPDATE
                        @temp
                    SET descrip38 = @ldescrip, base_imp38 = @base_imponible_deducible, monto_imp38 = @monto_imp_deducible, 
						descrip39 = 'Totales: ', base_imp39 = @base_imp_CFDeducible, monto_imp39 = @monto_imp_CFDeducible  

                FETCH TempTasas_CFDeducible INTO @ldescrip, @nac, @tasa, @base_imp, @monto_imp, @base_imponible_deducible, @monto_imp_deducible
            END
        DEALLOCATE TempTasas_CFDeducible

/****************!COMPRAS CON CRÉDITO FISCAL TOTALMENTE DEDUCIBLE****************/



/****************COMPRAS CON CRÉDITO FISCAL SUJETO A PRORRATEO*****************/


DECLARE @temp_CFProrrateo TABLE
            (
              [descrip] [char](100) ,
              [nac] [bit] ,
              [tasa] [decimal](18, 2) ,
              [base_imp] [decimal](18, 2) ,
              [monto_imp] [decimal](18, 2) ,
              [base_imponible_CFProrrateo] [DECIMAL] (18,2),
			  [monto_imp_CFProrrateo] [decimal](18,2)
            )
     
        DECLARE
            @ldescrip_CFProrrateo VARCHAR(100) ,
            @cont_CFProrrateo INT ,
			@base_imp_CFProrrateo DECIMAL (18,2),
			@monto_imp_CFProrrateo DECIMAL(18,2)
			
        DECLARE Tempdocs_CFProrrateo CURSOR
        FOR
            ( SELECT
                nro_doc, co_tipo_doc, total_neto, nac, anulado, base_imp, tasa, monto_imp, compras_exentas, base_imponible_prorrateo, monto_imp_prorrateo
              FROM
                DocumentosLibrocompras2(@sCo_fecha_d, @sCo_fecha_h)
              WHERE
                anulado = 0 AND tasa <> 0 AND @bImprimirColumnArt34 = 'SI' AND monto_imp_prorrateo <> 0
				--Sit.# 9986 ZPEREZ
					AND ( @cCo_Sucursal IS NULL
						  OR co_sucu_in = @cCo_Sucursal)
				--!Sit.# 9986 ZPEREZ
            )
            ORDER BY
            fecha_emis, n_control, nac, tasa
 
			SET @old_nro_doc = ''
			SET @old_co_tipo_doc = ''

			INSERT  INTO @temp_CFProrrateo
					( descrip, nac, tasa, base_imp, monto_imp, base_imponible_CFProrrateo, monto_imp_CFProrrateo )
			VALUES
					( 'Compras con Crédito Fiscal Sujeto a Prorrateo (Art.34)', 0, 0, NULL, NULL, NULL, NULL )
			
	        OPEN Tempdocs_CFProrrateo
		    FETCH Tempdocs_CFProrrateo INTO @nro_doc, @co_tipo_doc, @total_neto, @nac, @anulado, @base_imp, @tasa, @monto_imp,
			    @compras_exentas, @base_imponible_prorrateo, @monto_imp_prorrateo
	  
			WHILE @@fetch_status != -1 
				BEGIN
					IF NOT EXISTS ( SELECT
                                    tasa
                                FROM
                                    @temp_CFProrrateo
                                WHERE
                                    tasa = @tasa
									AND nac = @nac
								) 
					BEGIN

                        SET @ldescrip = 'Total Compras ' + CASE WHEN @nac = 1 THEN 'Internas'
                                                                ELSE 'Importación'
                                                           END + ' afectadas sólo alícuota '

                        INSERT  INTO @temp_CFProrrateo
                                ( descrip, nac, tasa, base_imp, monto_imp, base_imponible_CFProrrateo, monto_imp_CFProrrateo )
                        VALUES
                                ( @ldescrip + CAST(@tasa AS VARCHAR), @nac, @tasa, 0, 0, 0, 0 )
					END
					
					UPDATE
						@temp_CFProrrateo
					SET base_imp = base_imp + @base_imp, 
						monto_imp = monto_imp + @monto_imp , 
						base_imponible_CFProrrateo = base_imponible_CFProrrateo + @base_imponible_prorrateo, monto_imp_CFProrrateo = monto_imp_CFProrrateo + @monto_imp_prorrateo
					WHERE
						tasa = @tasa
						AND nac = @nac
					
					FETCH Tempdocs_CFProrrateo INTO @nro_doc, @co_tipo_doc, @total_neto, @nac, @anulado, @base_imp, @tasa, @monto_imp,
						@compras_exentas, @base_imponible_prorrateo, @monto_imp_prorrateo
				END
				
				IF NOT EXISTS(SELECT SUM(monto_imp_CFProrrateo) FROM @temp_CFProrrateo HAVING SUM(monto_imp_CFProrrateo) > 0) 
				BEGIN
					DELETE FROM @temp_CFProrrateo
				END
				
				
        DEALLOCATE Tempdocs_CFProrrateo	

        IF (@cont_CFDeducible = @contCCFND + 2 AND @contCCFND != @cont + 2) OR (@cont_CFDeducible = @contCCFND)
			--SI NO EXISTE UN CUADRO RESUMEN DE COMPRAS CON CRÉDITO FISCAL TOTALMENTE DEDUCIBLE
			--O SI NO EXISTE EL CUADRO DE COMPRAS TOTALMENTE DEDUCIBLE NI EL CUADRO DE COMPRAS NO DEDUCIBLE
			--NO DEBO AGREGAR MAS LÍNEAS PARA IMPRIMIR EL CUADRO RESUMEN DE COMPRAS CON CRÉDITO FISCAL SUJETO A PRORRATEO
			SET @cont_CFProrrateo = @cont_CFDeducible
		ELSE
			SET @cont_CFProrrateo = @cont_CFDeducible + 2

        SET @base_imp_CFProrrateo = 0
		SET @monto_imp_CFProrrateo = 0

        DECLARE TempTasas_CFProrrateo CURSOR
        FOR
            ( SELECT
                descrip, nac, tasa, base_imp, monto_imp, base_imponible_CFProrrateo, monto_imp_CFProrrateo
              FROM
                @temp_CFProrrateo
             )
            ORDER BY
            nac, tasa

        OPEN TempTasas_CFProrrateo
        FETCH TempTasas_CFProrrateo INTO @ldescrip, @nac, @tasa, @base_imp, @monto_imp, @base_imponible_prorrateo, @monto_imp_prorrateo
	  
        WHILE @@fetch_status != -1 
            BEGIN

                SET @cont_CFProrrateo =  @cont_CFProrrateo + 1
                SET @base_imp_CFProrrateo = @base_imp_CFProrrateo + ISNULL(@base_imponible_prorrateo,0)
				SET @monto_imp_CFProrrateo = @monto_imp_CFProrrateo + ISNULL(@monto_imp_prorrateo,0)
				
				
				IF @cont_CFProrrateo = 4
                    UPDATE
                        @temp
                    SET descrip4 = @ldescrip, base_imp4 = @base_imponible_prorrateo, monto_imp4 = @monto_imp_prorrateo, 
						descrip5 = 'Totales: ', base_imp5 = @base_imp_CFProrrateo, monto_imp5 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 5
                    UPDATE
                        @temp
                    SET descrip5 = @ldescrip, base_imp5 = @base_imponible_prorrateo, monto_imp5 = @monto_imp_prorrateo, 
						descrip6 = 'Totales: ', base_imp6 = @base_imp_CFProrrateo, monto_imp6 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 6
                    UPDATE
                        @temp
                    SET descrip6 = @ldescrip, base_imp6 = @base_imponible_prorrateo, monto_imp6 = @monto_imp_prorrateo, 
						descrip7 = 'Totales: ', base_imp7 = @base_imp_CFProrrateo, monto_imp7 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 7
                    UPDATE
                        @temp
                    SET descrip7 = @ldescrip, base_imp7 = @base_imponible_prorrateo, monto_imp7 = @monto_imp_prorrateo, 
						descrip8 = 'Totales: ', base_imp8 = @base_imp_CFProrrateo, monto_imp8 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 8
                    UPDATE
                        @temp
                    SET descrip8 = @ldescrip, base_imp8 = @base_imponible_prorrateo, monto_imp8 = @monto_imp_prorrateo, 
						descrip9 = 'Totales: ', base_imp9 = @base_imp_CFProrrateo, monto_imp9 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 9
                    UPDATE
                        @temp
                    SET descrip9 = @ldescrip, base_imp9 = @base_imponible_prorrateo, monto_imp9 = @monto_imp_prorrateo, 
						descrip10 = 'Totales: ', base_imp10 = @base_imp_CFProrrateo, monto_imp10 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 10
                    UPDATE
                        @temp
                    SET descrip10 = @ldescrip, base_imp10 = @base_imponible_prorrateo, monto_imp10 = @monto_imp_prorrateo, 
						descrip11 = 'Totales: ', base_imp11 = @base_imp_CFProrrateo, monto_imp11 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 11
                    UPDATE
                        @temp
                    SET descrip11 = @ldescrip, base_imp11 = @base_imponible_prorrateo, monto_imp11 = @monto_imp_prorrateo, 
						descrip12 = 'Totales: ', base_imp12 = @base_imp_CFProrrateo, monto_imp12 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 12
                    UPDATE
                        @temp
                    SET descrip12 = @ldescrip, base_imp12 = @base_imponible_prorrateo, monto_imp12 = @monto_imp_prorrateo, 
						descrip13 = 'Totales: ', base_imp13 = @base_imp_CFProrrateo, monto_imp13 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 13
                    UPDATE
                        @temp
                    SET descrip13 = @ldescrip, base_imp13 = @base_imponible_prorrateo, monto_imp13 = @monto_imp_prorrateo, 
						descrip14 = 'Totales: ', base_imp14 = @base_imp_CFProrrateo, monto_imp14 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 14
                    UPDATE
                        @temp
                    SET descrip14 = @ldescrip, base_imp14 = @base_imponible_prorrateo, monto_imp14 = @monto_imp_prorrateo, 
						descrip15 = 'Totales: ', base_imp15 = @base_imp_CFProrrateo, monto_imp15 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 15
                    UPDATE
                        @temp
                    SET descrip15 = @ldescrip, base_imp15 = @base_imponible_prorrateo, monto_imp15 = @monto_imp_prorrateo, 
						descrip16 = 'Totales: ', base_imp16 = @base_imp_CFProrrateo, monto_imp16 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 16
                    UPDATE
                        @temp
                    SET descrip16 = @ldescrip, base_imp16 = @base_imponible_prorrateo, monto_imp16 = @monto_imp_prorrateo, 
						descrip17 = 'Totales: ', base_imp17 = @base_imp_CFProrrateo, monto_imp17 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 17
                    UPDATE
                        @temp
                    SET descrip17 = @ldescrip, base_imp17 = @base_imponible_prorrateo, monto_imp17 = @monto_imp_prorrateo, 
						descrip18 = 'Totales: ', base_imp18 = @base_imp_CFProrrateo, monto_imp18 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 18
                    UPDATE
                        @temp
                    SET descrip18 = @ldescrip, base_imp18 = @base_imponible_prorrateo, monto_imp18 = @monto_imp_prorrateo, 
						descrip19 = 'Totales: ', base_imp19 = @base_imp_CFProrrateo, monto_imp19 = @monto_imp_CFProrrateo
				--Se necesita llegar hasta el campo descrip37 para abastecer un cuadro resumen que muestre todas las tasas
				IF @cont_CFProrrateo = 19
                    UPDATE
                        @temp
                    SET descrip19 = @ldescrip, base_imp19 = @base_imponible_prorrateo, monto_imp19 = @monto_imp_prorrateo, 
						descrip20 = 'Totales: ', base_imp20 = @base_imp_CFProrrateo, monto_imp20 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 20
                    UPDATE
                        @temp
                    SET descrip20 = @ldescrip, base_imp20 = @base_imponible_prorrateo, monto_imp20 = @monto_imp_prorrateo, 
						descrip21 = 'Totales: ', base_imp21 = @base_imp_CFProrrateo, monto_imp21 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 21
                    UPDATE
                        @temp
                    SET descrip21 = @ldescrip, base_imp21 = @base_imponible_prorrateo, monto_imp21 = @monto_imp_prorrateo, 
						descrip22 = 'Totales: ', base_imp22 = @base_imp_CFProrrateo, monto_imp22 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 22
                    UPDATE
                        @temp
                    SET descrip22 = @ldescrip, base_imp22 = @base_imponible_prorrateo, monto_imp22 = @monto_imp_prorrateo, 
						descrip23 = 'Totales: ', base_imp23 = @base_imp_CFProrrateo, monto_imp23 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 23
                    UPDATE
                        @temp
                    SET descrip23 = @ldescrip, base_imp23 = @base_imponible_prorrateo, monto_imp23 = @monto_imp_prorrateo, 
						descrip24 = 'Totales: ', base_imp24 = @base_imp_CFProrrateo, monto_imp24 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 24
                    UPDATE
                        @temp
                    SET descrip24 = @ldescrip, base_imp24 = @base_imponible_prorrateo, monto_imp24 = @monto_imp_prorrateo, 
						descrip25 = 'Totales: ', base_imp25 = @base_imp_CFProrrateo, monto_imp25 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 25
                    UPDATE
                        @temp
                    SET descrip25 = @ldescrip, base_imp25 = @base_imponible_prorrateo, monto_imp25 = @monto_imp_prorrateo, 
						descrip26 = 'Totales: ', base_imp26 = @base_imp_CFProrrateo, monto_imp26 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 26
                    UPDATE
                        @temp
                    SET descrip26 = @ldescrip, base_imp26 = @base_imponible_prorrateo, monto_imp26 = @monto_imp_prorrateo, 
						descrip27 = 'Totales: ', base_imp27 = @base_imp_CFProrrateo, monto_imp27 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 27
                    UPDATE
                        @temp
                    SET descrip27 = @ldescrip, base_imp27 = @base_imponible_prorrateo, monto_imp27 = @monto_imp_prorrateo, 
						descrip28 = 'Totales: ', base_imp28 = @base_imp_CFProrrateo, monto_imp28 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 28
                    UPDATE
                        @temp
                    SET descrip28 = @ldescrip, base_imp28 = @base_imponible_prorrateo, monto_imp28 = @monto_imp_prorrateo, 
						descrip29 = 'Totales: ', base_imp29 = @base_imp_CFProrrateo, monto_imp29 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 29
                    UPDATE
                        @temp
                    SET descrip29 = @ldescrip, base_imp29 = @base_imponible_prorrateo, monto_imp29 = @monto_imp_prorrateo, 
						descrip30 = 'Totales: ', base_imp30 = @base_imp_CFProrrateo, monto_imp30 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 30
                    UPDATE
                        @temp
                    SET descrip30 = @ldescrip, base_imp30 = @base_imponible_prorrateo, monto_imp30 = @monto_imp_prorrateo, 
						descrip31 = 'Totales: ', base_imp31 = @base_imp_CFProrrateo, monto_imp31 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 31
                    UPDATE
                        @temp
                    SET descrip31 = @ldescrip, base_imp31 = @base_imponible_prorrateo, monto_imp31 = @monto_imp_prorrateo, 
						descrip32 = 'Totales: ', base_imp32 = @base_imp_CFProrrateo, monto_imp32 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 32
                    UPDATE
                        @temp
                    SET descrip32 = @ldescrip, base_imp32 = @base_imponible_prorrateo, monto_imp32 = @monto_imp_prorrateo, 
						descrip33 = 'Totales: ', base_imp33 = @base_imp_CFProrrateo, monto_imp33 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 33
                    UPDATE
                        @temp
                    SET descrip33 = @ldescrip, base_imp33 = @base_imponible_prorrateo, monto_imp33 = @monto_imp_prorrateo, 
						descrip34 = 'Totales: ', base_imp34 = @base_imp_CFProrrateo, monto_imp34 = @monto_imp_CFProrrateo  
				IF @cont_CFProrrateo = 34
                    UPDATE
                        @temp
                    SET descrip34 = @ldescrip, base_imp34 = @base_imponible_prorrateo, monto_imp34 = @monto_imp_prorrateo, 
						descrip35 = 'Totales: ', base_imp35 = @base_imp_CFProrrateo, monto_imp35 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 35
                    UPDATE
                        @temp
                    SET descrip35 = @ldescrip, base_imp35 = @base_imponible_prorrateo, monto_imp35 = @monto_imp_prorrateo, 
						descrip36 = 'Totales: ', base_imp36 = @base_imp_CFProrrateo, monto_imp36 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 36
                    UPDATE
                        @temp
                    SET descrip36 = @ldescrip, base_imp36 = @base_imponible_prorrateo, monto_imp36 = @monto_imp_prorrateo, 
						descrip37 = 'Totales: ', base_imp37 = @base_imp_CFProrrateo, monto_imp37 = @monto_imp_CFProrrateo  
-----------------------------------------------------------
				IF @cont_CFProrrateo = 37
                    UPDATE
                        @temp
                    SET descrip37 = @ldescrip, base_imp37 = @base_imponible_prorrateo, monto_imp37 = @monto_imp_prorrateo, 
						descrip38 = 'Totales: ', base_imp38 = @base_imp_CFProrrateo, monto_imp38 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 38
                    UPDATE
                        @temp
                    SET descrip38 = @ldescrip, base_imp38 = @base_imponible_prorrateo, monto_imp38 = @monto_imp_prorrateo, 
						descrip39 = 'Totales: ', base_imp39 = @base_imp_CFProrrateo, monto_imp39 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 39
                    UPDATE
                        @temp
                    SET descrip39 = @ldescrip, base_imp39 = @base_imponible_prorrateo, monto_imp39 = @monto_imp_prorrateo, 
						descrip40 = 'Totales: ', base_imp40 = @base_imp_CFProrrateo, monto_imp40 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 40
                    UPDATE
                        @temp
                    SET descrip40 = @ldescrip, base_imp40 = @base_imponible_prorrateo, monto_imp40 = @monto_imp_prorrateo, 
						descrip41 = 'Totales: ', base_imp41 = @base_imp_CFProrrateo, monto_imp41 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 41
                    UPDATE
                        @temp
                    SET descrip41 = @ldescrip, base_imp41 = @base_imponible_prorrateo, monto_imp41 = @monto_imp_prorrateo, 
						descrip42 = 'Totales: ', base_imp42 = @base_imp_CFProrrateo, monto_imp42 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 42
                    UPDATE
                        @temp
                    SET descrip42 = @ldescrip, base_imp42 = @base_imponible_prorrateo, monto_imp42 = @monto_imp_prorrateo, 
						descrip43 = 'Totales: ', base_imp43 = @base_imp_CFProrrateo, monto_imp43 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 43
                    UPDATE
                        @temp
                    SET descrip43 = @ldescrip, base_imp43 = @base_imponible_prorrateo, monto_imp43 = @monto_imp_prorrateo, 
						descrip44 = 'Totales: ', base_imp44 = @base_imp_CFProrrateo, monto_imp44 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 44
                    UPDATE
                        @temp
                    SET descrip44 = @ldescrip, base_imp44 = @base_imponible_prorrateo, monto_imp44 = @monto_imp_prorrateo, 
						descrip45 = 'Totales: ', base_imp45 = @base_imp_CFProrrateo, monto_imp45 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 45
                    UPDATE
                        @temp
                    SET descrip45 = @ldescrip, base_imp45 = @base_imponible_prorrateo, monto_imp45 = @monto_imp_prorrateo, 
						descrip46 = 'Totales: ', base_imp46 = @base_imp_CFProrrateo, monto_imp46 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 46
                    UPDATE
                        @temp
                    SET descrip46 = @ldescrip, base_imp46 = @base_imponible_prorrateo, monto_imp46 = @monto_imp_prorrateo, 
						descrip47 = 'Totales: ', base_imp47 = @base_imp_CFProrrateo, monto_imp47 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 47
                    UPDATE
                        @temp
                    SET descrip47 = @ldescrip, base_imp47 = @base_imponible_prorrateo, monto_imp47 = @monto_imp_prorrateo, 
						descrip48 = 'Totales: ', base_imp48 = @base_imp_CFProrrateo, monto_imp48 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 48
                    UPDATE
                        @temp
                    SET descrip48 = @ldescrip, base_imp48 = @base_imponible_prorrateo, monto_imp48 = @monto_imp_prorrateo, 
						descrip49 = 'Totales: ', base_imp49 = @base_imp_CFProrrateo, monto_imp49 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 49
                    UPDATE
                        @temp
                    SET descrip49 = @ldescrip, base_imp49 = @base_imponible_prorrateo, monto_imp49 = @monto_imp_prorrateo, 
						descrip50 = 'Totales: ', base_imp50 = @base_imp_CFProrrateo, monto_imp50 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 50
                    UPDATE
                        @temp
                    SET descrip50 = @ldescrip, base_imp50 = @base_imponible_prorrateo, monto_imp50 = @monto_imp_prorrateo, 
						descrip51 = 'Totales: ', base_imp51 = @base_imp_CFProrrateo, monto_imp51 = @monto_imp_CFProrrateo 
				IF @cont_CFProrrateo = 51
                    UPDATE
                        @temp
                    SET descrip51 = @ldescrip, base_imp51 = @base_imponible_prorrateo, monto_imp51 = @monto_imp_prorrateo, 
						descrip52 = 'Totales: ', base_imp52 = @base_imp_CFProrrateo, monto_imp52 = @monto_imp_CFProrrateo   

                FETCH TempTasas_CFProrrateo INTO @ldescrip, @nac, @tasa, @base_imp, @monto_imp, @base_imponible_prorrateo, @monto_imp_prorrateo
            END
        DEALLOCATE TempTasas_CFProrrateo


/****************!COMPRAS CON CRÉDITO FISCAL SUJETO A PRORRATEO*****************/

/********************************************!CUADRO RESUMEN ART.34********************************************/

        IF Not Exists(SELECT * FROM	@Temp WHERE ( @cCo_Sucursal IS NULL OR co_sucu_in = @cCo_Sucursal)/*ORDER BY CAST(fecha_emis AS DATE), fec_reg, n_control, tasa*/) 
        BEGIN   
			INSERT INTO @temp (
              [prov_des], [base_imp] ,[compras_exentas] ,[monto_imp] ,[monto_ret_imp], [monto_ret_imp_tercero]  )
              VALUES ('No hubo movimientos en el mes', 0,0,0,0,0)
        END
        
        SELECT
            *
        FROM
            @Temp
        --F.D. 13/12/2012 Se coloco el filtro sucursal    
        WHERE ( @cCo_Sucursal IS NULL
                  OR co_sucu_in = @cCo_Sucursal)
        ORDER BY
             fecha_emis, fe_us_in, fec_reg, n_control, tasa -- Se agrego fec_reg
    END



