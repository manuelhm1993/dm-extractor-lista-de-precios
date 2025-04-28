USE [DM_ADM]
GO
/****** Object:  StoredProcedure [dbo].[ARepArticuloConPrecio1]    Script Date: 25/4/2025 9:20:09 a. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



-- =============================================
-- Author:		<SOFTECH SISTEMAS>
-- Create date: <30-04-10>
-- Description:	<Reporte de Artículos con sus Precios>
-- =============================================
ALTER PROCEDURE  [dbo].[ARepArticuloConPrecio1]
	-- Add the parameters for the stored procedure here
	@sCo_Art_d CHAR(30) = NULL,
	@sCo_Art_h CHAR(30) = NULL,
	@sCo_Linea_d char(6) = NULL,
	@sCo_Linea_h char(6) = NULL,
	@sCo_SubLinea_d char(6) = NULL,
	@sCo_SubLinea_h char(6) = NULL,
	@sCo_Categoria_d char(6) = NULL,
	@sCo_Categoria_h char(6) = NULL,
	@sCo_Color_d char(6) = NULL,
	@sCo_Color_h char(6) = NULL,
	@sCo_Almacen char(6) = NULL,
	@sCo_Precio05 char(6) = NULL,
	@sCo_NivelStock CHAR(4) = NULL ,
	@sCo_FechaHasta datetime = NULL,
	@sCo_Precio01 char(6) = NULL,
	@sCo_Precio02 char(6) = NULL,
	@sCo_Precio03 char(6) = NULL,
	@sCo_Precio04 char(6) = NULL,
	@sCo_Clasificado char(4)= NULL,	----->Filtro Clasificado por
	@sCo_FechaMRLL_d datetime = NULL,
	@sCo_FechaMRLL_h datetime = NULL,
	@sCo_MostrarMRLL char(2)= NULL,
	@bIncluirImpuesto char(2) = NULL,
	@sCo_Sucursal char(6) = NULL,
	@sCampOrderBy varchar(16) = NULL,
	@sDir varchar(6) = NULL,
	@bHeaderRep bit = 0
AS
BEGIN
	SET NOCOUNT ON;

	 IF (@sCo_Almacen IS NULL or @sCo_Precio05 IS NULL)
	    BEGIN
		  RAISERROR('Debe seleccionar los 2 almacenes',16, 1);
		  RETURN -1
		END  
	
	IF (@sCo_Precio01 IS NULL AND @sCo_Precio02 IS NULL AND @sCo_Precio03 IS NULL AND @sCo_Precio04 IS NULL)
	    BEGIN
		  RAISERROR('Debe seleccionar un Tipo de Precio',16, 1);
		  RETURN -1
		END  
	
	
	
	DECLARE @bIncluirImpuestoCalculo bit

--------------Valores por Defecto---------------- 
	IF @sCo_NivelStock IS NULL 
		SET @sCo_NivelStock = 'TODO' 
		
	IF @sCo_FechaHasta IS NULL
		SET @sCo_FechaHasta = getdate();

	SET @sCo_FechaHasta = dbo.FechaConMinutos(@sCo_FechaHasta)

	IF (@sCo_Clasificado IS NULL)
		SET @sCo_Clasificado = ''

	IF (@bIncluirImpuesto IS NULL or @bIncluirImpuesto = 'NO' )
		SET @bIncluirImpuestoCalculo = 0
	else
		SET @bIncluirImpuestoCalculo = 1
	
	IF @sCo_MostrarMRLL IS NULL 
	SET @sCo_MostrarMRLL = 'NO' 

--------------Fin Valores por Defecto---------------- 
IF  EXISTS (SELECT * FROM SYSOBJECTS WHERE TYPE = 'U' AND NAME = 'LC_STOCK_ALMA')
BEGIN
	DROP TABLE LC_STOCK_ALMA
END
IF  EXISTS (SELECT * FROM SYSOBJECTS WHERE TYPE = 'U' AND NAME = 'LC_ART_ORDEN')
BEGIN
	DROP TABLE LC_ART_ORDEN
END
IF  EXISTS (SELECT * FROM SYSOBJECTS WHERE TYPE = 'U' AND NAME = 'LC_STOCK_ALMA1')
BEGIN
	DROP TABLE LC_STOCK_ALMA1
END

SELECT CO_ART,ART_DES, RANK ( ) OVER ( ORDER BY CO_LIN,CO_SUBL,ART_DES ) ORDEN INTO LC_ART_ORDEN
FROM SAARTICULO
ORDER BY CO_LIN,CO_SUBL,ART_DES

SELECT @sCo_Almacen = @sCo_Almacen

Select A.co_art, A.art_des, UP.co_uni,A.modelo AS des_uni, 
	AL.co_alma, AL.des_alma, ISNULL(B.STOCK,0.00000) AS StockActual,
	A.co_lin, D.lin_des, 
	A.co_subl, E.subl_des, 
	T.co_cat, C.des_color AS  cat_des,
	@sCo_FechaHasta as FechaPrecio,
	PRE01.co_precio as co_precio01, PRE01.des_precio as des_precio01, 
	CASE when PRE01.co_precio is not null then
	(select top 1 PR.monto from saArtPrecio PR where PR.co_art = A.co_art and PR.co_precio = PRE01.co_precio order by desde desc)
	else null END as Precio01,

	PRE02.co_precio as co_precio02, PRE02.des_precio as des_precio02, 
	CASE when PRE02.co_precio is not null then
	(select top 1 PR.monto from saArtPrecio PR where PR.co_art = A.co_art and PR.co_precio = PRE02.co_precio order by desde desc)
	else null END as Precio02,

	PRE03.co_precio as co_precio03, PRE03.des_precio as des_precio03, 
	CASE when PRE03.co_precio is not null then
	(select top 1 PR.monto from saArtPrecio PR where PR.co_art = A.co_art and PR.co_precio = PRE03.co_precio order by desde desc)
	else null END as Precio03,

	PRE04.co_precio as co_precio04, PRE04.des_precio as des_precio04, 
	CASE when PRE04.co_precio is not null then
	(select top 1 PR.monto from saArtPrecio PR where PR.co_art = A.co_art and PR.co_precio = PRE04.co_precio order by desde desc)
	else null END as Precio04,

	null as Precio05,
	@sCo_Clasificado as Clasificado INTO LC_STOCK_ALMA
 from
	saArticulo AS A
		CROSS JOIN saAlmacen AS AL
		LEFT JOIN saStockAlmacen B ON A.co_art = B.co_art
                                          AND B.tipo = 'ACT'
                                          AND AL.co_alma = B.co_alma
		LEFT JOIN saArtUnidad AS AUP ON	  AUP.co_art = A.co_art
                                          AND AUP.uni_principal = 1
        LEFT JOIN saUnidad AS UP ON UP.co_uni = AUP.co_uni
		INNER JOIN saLineaArticulo AS D ON A.co_lin = D.co_lin
        INNER JOIN saSublinea AS E ON A.co_lin = E.co_lin AND A.co_subl = E.co_subl
		INNER JOIN saCatArticulo AS T ON A.co_cat = T.co_cat
		INNER JOIN saColor AS C ON A.co_color = C.co_color
		LEFT JOIN saTipoPrecio PRE01 ON PRE01.co_precio = @sCo_Precio01 and @sCo_Precio01 is not null
		LEFT JOIN saTipoPrecio PRE02 ON PRE02.co_precio = @sCo_Precio02 and @sCo_Precio02 is not null
		LEFT JOIN saTipoPrecio PRE03 ON PRE03.co_precio = @sCo_Precio03 and @sCo_Precio03 is not null
		LEFT JOIN saTipoPrecio PRE04 ON PRE04.co_precio = @sCo_Precio04 and @sCo_Precio04 is not null
	Where
		(@sCo_art_d IS NULL OR A.co_art >= @sCo_art_d) AND
		(@sCo_art_h IS NULL OR A.co_art <= @sCo_art_h) AND
		(@sCo_Linea_d IS NULL OR A.co_lin >= @sCo_Linea_d) AND 
		(@sCo_Linea_h IS NULL OR A.co_lin <= @sCo_Linea_h) AND
		(@sCo_SubLinea_d IS NULL OR A.co_subl >= @sCo_SubLinea_d) AND 
		(@sCo_SubLinea_h IS NULL OR A.co_subl <= @sCo_SubLinea_h) AND
		(@sCo_Categoria_d IS NULL OR A.co_cat >= @sCo_Categoria_d) AND 
		(@sCo_Categoria_h IS NULL OR A.co_cat <= @sCo_Categoria_h) AND
		(@sCo_Color_d IS NULL OR A.co_color >= @sCo_Color_d) AND 
		(@sCo_Color_h IS NULL OR A.co_color <= @sCo_Color_h) AND
		(@sCo_Almacen IS NULL OR @sCo_Almacen = AL.co_alma)
		AND ( @sCo_Sucursal IS NULL OR A.co_sucu_in = @sCo_Sucursal)
ORDER BY 
(CASE 
		WHEN @sCo_Clasificado = 'LIN' THEN A.co_lin 
		WHEN @sCo_Clasificado = 'LIN' THEN A.co_art
		WHEN @sCo_Clasificado = 'SBL' THEN A.co_lin 
		WHEN @sCo_Clasificado = 'SBL' THEN A.co_subl
		WHEN @sCo_Clasificado = 'SBL' THEN A.co_art
		WHEN @sCo_Clasificado = 'CAT' THEN A.co_cat
		WHEN @sCo_Clasificado = 'CAT' THEN A.co_art
		ELSE A.co_art
		END),	
CASE @sDir WHEN 'DESC' THEN  CASE @sCampOrderBy WHEN 'art_des' THEN A.art_des	ELSE A.co_art END END DESC,
CASE @sDir WHEN 'ASC' THEN	CASE @sCampOrderBy WHEN 'art_des' THEN A.art_des ELSE A.co_art  END END	ASC

SELECT @sCo_Almacen = @sCo_Precio05

Select A.co_art, A.art_des, UP.co_uni,A.modelo AS des_uni, 
	AL.co_alma, AL.des_alma, ISNULL(B.STOCK,0.00000) AS StockActual,
	A.co_lin, D.lin_des, 
	A.co_subl, E.subl_des, 
	T.co_cat, C.des_color AS  cat_des,
	@sCo_FechaHasta as FechaPrecio,
	PRE01.co_precio as co_precio01, PRE01.des_precio as des_precio01, 
	CASE when PRE01.co_precio is not null then
	(select top 1 PR.monto from saArtPrecio PR where PR.co_art = A.co_art and PR.co_precio = PRE01.co_precio order by desde desc)
	else null END as Precio01,

	PRE02.co_precio as co_precio02, PRE02.des_precio as des_precio02, 
	CASE when PRE02.co_precio is not null then
	(select top 1 PR.monto from saArtPrecio PR where PR.co_art = A.co_art and PR.co_precio = PRE02.co_precio order by desde desc)
	else null END as Precio02,

	PRE03.co_precio as co_precio03, PRE03.des_precio as des_precio03, 
	CASE when PRE03.co_precio is not null then
	(select top 1 PR.monto from saArtPrecio PR where PR.co_art = A.co_art and PR.co_precio = PRE03.co_precio order by desde desc)
	else null END as Precio03,

	PRE04.co_precio as co_precio04, PRE04.des_precio as des_precio04, 
	CASE when PRE04.co_precio is not null then
	(select top 1 PR.monto from saArtPrecio PR where PR.co_art = A.co_art and PR.co_precio = PRE04.co_precio order by desde desc)
	else null END as Precio04,

	null as Precio05,
	@sCo_Clasificado as Clasificado INTO LC_STOCK_ALMA1
 from
	saArticulo AS A
		CROSS JOIN saAlmacen AS AL
		LEFT JOIN saStockAlmacen B ON A.co_art = B.co_art
                                          AND B.tipo = 'ACT'
                                          AND AL.co_alma = B.co_alma
		LEFT JOIN saArtUnidad AS AUP ON	  AUP.co_art = A.co_art
                                          AND AUP.uni_principal = 1
        LEFT JOIN saUnidad AS UP ON UP.co_uni = AUP.co_uni
		INNER JOIN saLineaArticulo AS D ON A.co_lin = D.co_lin
        INNER JOIN saSublinea AS E ON A.co_lin = E.co_lin AND A.co_subl = E.co_subl
		INNER JOIN saCatArticulo AS T ON A.co_cat = T.co_cat
		INNER JOIN saColor AS C ON A.co_color = C.co_color
		LEFT JOIN saTipoPrecio PRE01 ON PRE01.co_precio = @sCo_Precio01 and @sCo_Precio01 is not null
		LEFT JOIN saTipoPrecio PRE02 ON PRE02.co_precio = @sCo_Precio02 and @sCo_Precio02 is not null
		LEFT JOIN saTipoPrecio PRE03 ON PRE03.co_precio = @sCo_Precio03 and @sCo_Precio03 is not null
		LEFT JOIN saTipoPrecio PRE04 ON PRE04.co_precio = @sCo_Precio04 and @sCo_Precio04 is not null
	Where
		(@sCo_art_d IS NULL OR A.co_art >= @sCo_art_d) AND
		(@sCo_art_h IS NULL OR A.co_art <= @sCo_art_h) AND
		(@sCo_Linea_d IS NULL OR A.co_lin >= @sCo_Linea_d) AND 
		(@sCo_Linea_h IS NULL OR A.co_lin <= @sCo_Linea_h) AND
		(@sCo_SubLinea_d IS NULL OR A.co_subl >= @sCo_SubLinea_d) AND 
		(@sCo_SubLinea_h IS NULL OR A.co_subl <= @sCo_SubLinea_h) AND
		(@sCo_Categoria_d IS NULL OR A.co_cat >= @sCo_Categoria_d) AND 
		(@sCo_Categoria_h IS NULL OR A.co_cat <= @sCo_Categoria_h) AND
		(@sCo_Color_d IS NULL OR A.co_color >= @sCo_Color_d) AND 
		(@sCo_Color_h IS NULL OR A.co_color <= @sCo_Color_h) AND
		(@sCo_Almacen IS NULL OR @sCo_Almacen = AL.co_alma)
		AND ( @sCo_Sucursal IS NULL OR A.co_sucu_in = @sCo_Sucursal)
ORDER BY 
(CASE 
		WHEN @sCo_Clasificado = 'LIN' THEN A.co_lin 
		WHEN @sCo_Clasificado = 'LIN' THEN A.co_art
		WHEN @sCo_Clasificado = 'SBL' THEN A.co_lin 
		WHEN @sCo_Clasificado = 'SBL' THEN A.co_subl
		WHEN @sCo_Clasificado = 'SBL' THEN A.co_art
		WHEN @sCo_Clasificado = 'CAT' THEN A.co_cat
		WHEN @sCo_Clasificado = 'CAT' THEN A.co_art
		ELSE A.co_art
		END),	
CASE @sDir 
    WHEN 'DESC' THEN  
		CASE @sCampOrderBy 
			WHEN 'art_des' THEN A.art_des
			ELSE A.co_art
		END 
END 
    DESC,
CASE @sDir 
    WHEN 'ASC' THEN	
		CASE @sCampOrderBy 
			WHEN 'art_des' THEN A.art_des
			ELSE A.co_art
        END 
END
	ASC
END

--GENERAR CURSOR TEMPORAL FINAL 
DECLARE pointer CURSOR FOR
SELECT T.co_art, T.art_des, T.co_uni, T.des_uni, SUM(T.StockActual) AS StockActual, T.co_lin, T.lin_des, T.co_subl, T.subl_des, T.co_cat,T.cat_des, T.FechaPrecio, T.co_precio01, T.des_precio01, 
T.Precio01, T.co_precio02, T.des_precio02, T.Precio02, T.co_precio03, T.des_precio03, T.Precio03, T.co_precio04, T.des_precio04, T.Precio04, T.Precio05 , T.Clasificado,
(select orden from LC_ART_ORDEN where co_art = t.co_art) as co_cat1
 FROM (SELECT * FROM LC_STOCK_ALMA UNION ALL SELECT * FROM LC_STOCK_ALMA1) AS T 
 GROUP BY T.co_art, T.art_des, T.co_uni, T.des_uni, T.co_lin, T.lin_des, T.co_subl, T.subl_des, T.co_cat,T.cat_des, T.FechaPrecio, T.co_precio01, T.des_precio01, 
T.Precio01, T.co_precio02, T.des_precio02, T.Precio02, T.co_precio03, T.des_precio03, T.Precio03, T.co_precio04, T.des_precio04, T.Precio04, T.Precio05 , T.Clasificado
 ORDER BY T.lin_des,T.subl_des
 
 
 DECLARE @co_art CHAR(30), @art_des VARCHAR(120), @co_uni CHAR(6), @des_uni VARCHAR(60), @StockActual DECIMAL(18,5), @co_lin CHAR(6), @lin_des VARCHAR(60), @co_subl CHAR(6), @subl_des VARCHAR(60), @co_cat CHAR(6) ,@cat_des VARCHAR(60), @FechaPrecio SMALLDATETIME, @co_precio01 CHAR(6) , @des_precio01 VARCHAR(60), 
@Precio01 DECIMAL(18,5), @co_precio02 CHAR(6), @des_precio02 VARCHAR(60), @Precio02 DECIMAL(18,5), @co_precio03 CHAR(6), @des_precio03 VARCHAR(60), @Precio03 DECIMAL(18,5), @co_precio04 CHAR(6), @des_precio04 VARCHAR(60), @Precio04 DECIMAL(18,5), @Precio05 DECIMAL(18,5), @Clasificado VARCHAR(60), @co_cat1 CHAR(6)
	
DECLARE @temp table (co_art CHAR(30), art_des VARCHAR(120), co_uni CHAR(6), des_uni VARCHAR(60), StockActual DECIMAL(18,5), co_lin CHAR(6), lin_des VARCHAR(60), co_subl CHAR(6), subl_des VARCHAR(60), co_cat CHAR(6) ,cat_des VARCHAR(60), FechaPrecio SMALLDATETIME, co_precio01 CHAR(6) , des_precio01 VARCHAR(60), 
Precio01 DECIMAL(18,5), co_precio02 CHAR(6), des_precio02 VARCHAR(60), Precio02 DECIMAL(18,5), co_precio03 CHAR(6), des_precio03 VARCHAR(60), Precio03 DECIMAL(18,5), co_precio04 CHAR(6), des_precio04 VARCHAR(60), Precio04 DECIMAL(18,5), Precio05 DECIMAL(18,5), Clasificado VARCHAR(60), co_cat1 CHAR(6))
	
OPEN pointer;

FETCH NEXT FROM pointer INTO @co_art, @art_des, @co_uni , @des_uni , @StockActual , @co_lin , @lin_des , @co_subl , @subl_des , @co_cat  ,@cat_des, @FechaPrecio , @co_precio01 , @des_precio01, 
@Precio01 , @co_precio02, @des_precio02, @Precio02, @co_precio03, @des_precio03, @Precio03, @co_precio04, @des_precio04, @Precio04, @Precio05, @Clasificado, @co_cat1

WHILE @@FETCH_STATUS = 0 

BEGIN 
		
	INSERT INTO @temp (co_art, art_des, co_uni , des_uni , StockActual , co_lin , lin_des , co_subl , subl_des , co_cat  ,cat_des, FechaPrecio , co_precio01 , des_precio01, 
Precio01 , co_precio02, des_precio02, Precio02, co_precio03, des_precio03, Precio03, co_precio04, des_precio04, Precio04, Precio05, Clasificado, co_cat1)
				values(@co_art, @art_des, @co_uni , @des_uni , @StockActual , @co_lin , @lin_des , @co_subl , @subl_des , @co_cat  ,@cat_des, @FechaPrecio , @co_precio01 , @des_precio01, 
@Precio01 , @co_precio02, @des_precio02, @Precio02, @co_precio03, @des_precio03, @Precio03, @co_precio04, @des_precio04, @Precio04, @Precio05, @Clasificado, @co_cat1)
	
	FETCH NEXT FROM pointer INTO @co_art, @art_des, @co_uni , @des_uni , @StockActual , @co_lin , @lin_des , @co_subl , @subl_des , @co_cat  ,@cat_des, @FechaPrecio , @co_precio01 , @des_precio01, 
@Precio01 , @co_precio02, @des_precio02, @Precio02, @co_precio03, @des_precio03, @Precio03, @co_precio04, @des_precio04, @Precio04, @Precio05, @Clasificado, @co_cat1
END

CLOSE pointer;  
DEALLOCATE pointer; 

SELECT T.co_art, T.art_des, T.co_uni, T.des_uni, T.StockActual, T.co_lin, T.lin_des, T.co_subl, T.subl_des, T.co_cat, T.cat_des, 
T.FechaPrecio,
T.co_precio01,T.des_precio01, T.Precio01,
T.co_precio02,T.des_precio02, T.Precio02,
T.co_precio03,T.des_precio03, T.Precio03,
CASE WHEN T.co_art IN 
(SELECT FVR.co_art FROM saFacturaCompra AS FC INNER JOIN saFacturaCompraReng AS FVR ON FVR.doc_num = FC.doc_num 
WHERE FC.anulado = 0 AND ( ( @sCo_FechaMRLL_d IS NULL OR dbo.FechaSimple(FC.fec_emis) >= @sCo_FechaMRLL_d ) AND ( @sCo_FechaMRLL_h IS NULL OR dbo.FechaSimple(FC.fec_emis) <= @sCo_FechaMRLL_h ))
GROUP BY FVR.co_art)
THEN '1' ELSE '0' END as co_precio04, 

@sCo_MostrarMRLL as des_precio04, T.Precio04,
T.Precio05, T.Clasificado, T.co_cat1, 
(select u.des_ubicacion from saUbicacion U where U.co_ubicacion = A.co_ubicacion) as ubi

FROM @temp AS T INNER JOIN saArticulo AS A ON A.co_art = T.co_art
WHERE 
(@sCo_NivelStock = 'TODO') -- TODOS
OR (@sCo_NivelStock = 'IAO' AND ISNULL(StockActual,0.00000)  = 0) -- IGUAL A CERO
OR (@sCo_NivelStock = 'MAY' AND ISNULL(StockActual,0.00000) > 0)-- MAYOR A CERO
OR (@sCo_NivelStock = 'MEN' AND ISNULL(StockActual,0.00000) < 0) -- MENOR A CERO
OR (@sCo_NivelStock = 'PMAX' AND ISNULL(StockActual,0.00000) >= A.stock_max) -- MAYOR A PUNTO MAXIMO
OR (@sCo_NivelStock = 'PMIN' AND ISNULL(StockActual,0.00000) <= A.stock_min) -- MENOR A PUNTO MINIMO
OR (@sCo_NivelStock = 'PPED' AND ISNULL(StockActual,0.00000) <= A.stock_pedido) -- MENOR A PUNTO DE PEDIDO
OR (@sCo_NivelStock = 'DIFE' AND ISNULL(StockActual,0.00000) <> 0) --Diferente a CERO
order by T.lin_des, T.subl_des
