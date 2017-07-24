DECLARE @Inicio_Incidente		bigint
DECLARE @Final_Incidente		bigint
DECLARE @Temp					varchar(20)
DECLARE @Desfase_SAF			int
DECLARE @Desfase_Reintento		int
DECLARE @T_SAFS					bigint
DECLARE @T_Reintento			bigint

DECLARE @Totales	AS TABLE(
organizacion smallint,
TOTAL int,
TBK int,
RBC int,
VISA int,
MC int,
AMEX int,
TOTAL_S int,
TBK_S int,
RBC_S int,
VISA_S int,
MC_S int,
AMEX_S int)

DECLARE @DATA		AS TABLE(
stan char(6),
tiempoauth bigint,
ctran smallint,
pcode char(6),
organizacion smallint,
pan numeric(19,0),
montofac money,
tasacambio char(8),
bloqueo varchar(100),
codrsp varchar(3),
razon varchar(80),
codaut varchar(6),
paisiata char(2),
rubro char(4),
nomloccomercio varchar(40),
pem varchar(12),
dseg1 varchar(76),
dseg2 varchar(37),
idcomercio varchar(15),
idterminal varchar(16),
securityinfo varchar(16),
adv varchar(5),
nacInt varchar(15),
estado char(2),
disponible varchar(14),
rMarca char(3),
rNexus char(3)
)

-----------------------VARIABLES A UTILIZAR-----------------------

SET		@Desfase_SAF		=	2				--modificar
SET		@Desfase_Reintento	=	4				--modificar
SET		@Inicio_Incidente	=	170716151100	--modificar
SET		@Final_Incidente	=	170716154100	--modificar
SET		@Temp				=	'20'+SUBSTRING(CONVERT(VARCHAR(12),@Final_Incidente),1,2) +
								'-' + SUBSTRING(CONVERT(VARCHAR(12),@Final_Incidente),3,2)
								+ '-' + SUBSTRING(CONVERT(VARCHAR(12),@Final_Incidente),5,2)
								+ ' ' + SUBSTRING(CONVERT(VARCHAR(12),@Final_Incidente),7,2)
								+ ':' + SUBSTRING(CONVERT(VARCHAR(12),@Final_Incidente),9,2)
								+ ':' + SUBSTRING(CONVERT(VARCHAR(12),@Final_Incidente),11,2)
SET  @T_SAFS				=	CAST(convert(varchar, DateADD(HH, @Desfase_SAF, @Temp), 12) + replace(convert(varchar, DateADD(HH, @Desfase_SAF, @Temp), 108), ':', '') as bigint  )
SET  @T_Reintento			=	CAST(convert(varchar, DateADD(HH, @Desfase_Reintento, @Temp), 12) + replace(convert(varchar, DateADD(HH, @Desfase_Reintento, @Temp), 108), ':', '') as bigint  )

-----------------------GENERACIÓN DE LA DATA -----------------------

--TABLA CON DATA DEL INCIDENTE
INSERT INTO @DATA
	select stan,tiempoauth,ctran,pcode,organizacion,pan,montoFac,tasacambio,bloqueo,codRsp,razon,codAut,paisIata,rubro,nomLocComercio,pem,dseg1,dseg2,idcomercio,idterminal,securityinfo,
	
	CASE		
		WHEN ctran = '220' and charindex(',p,',bloqueo)>0 THEN 'TBK'	
		WHEN ctran = '120' and ( charindex(',102',bloqueo)>0 or charindex(',111',bloqueo)>0  or charindex(',113',bloqueo)>0 or charindex(',109',bloqueo)>0 ) THEN 'MC'
		WHEN ctran = '120' and ( charindex(',9001',bloqueo)>0 or charindex(',9020',bloqueo)>0  or charindex(',9011',bloqueo)>0 ) THEN 'VISA'
		WHEN ctran in (1120,1220) THEN 'AMEX'
		WHEN ctran in ('220') and rubro in ('6011') THEN 'RBC'
		ELSE ''	
	END as 'ADV',
	
	IIF(paisiata='CL', 'Nacional', 'Internacional') as 'NacInt',
	
	right(left(origDataElem,3),2) estado,


	IIF(
		PATINDEX('%[1-9-]%',right(left(origDataElem,16),12))=0,cast(0 as bigint),
		substring(
			right(left(origDataElem,16),12),
			PATINDEX('%[1-9-]%',right(left(origDataElem,16),12)),
			13-PATINDEX('%[1-9-]%', right(left(origDataElem,16),12))
		)
	)disponible,

	IIF(
		IIF(
		PATINDEX('%[1-9-]%',right(left(origDataElem,16),12))=0,cast(0 as bigint),
			substring(
				right(left(origDataElem,16),12),
				PATINDEX('%[1-9-]%',right(left(origDataElem,16),12)),
				13-PATINDEX('%[1-9-]%', right(left(origDataElem,16),12))
			)
		)
		>=
		montofac and right(left(origDataElem,3),2)='01','Aut','Den'
	) rMarca,

	IIF(codrsp in ('00','85'),'Aut','Den') rNexus

	from IsoNxs.dbo.journal (nolock)
	where tiempoauth>=@Inicio_Incidente and tiempoauth<@Final_Incidente
	and ( (ctran = '120' and ( charindex(',102',bloqueo)>0 or charindex(',111',bloqueo)>0  or charindex(',113',bloqueo)>0 or charindex(',109',bloqueo)>0 )) or (ctran = '220' and charindex(',p,',bloqueo)>0) or (ctran = '120' and ( charindex(',9001',bloqueo)>0 or charindex(',9020',bloqueo)>0  or charindex(',9011',bloqueo)>0 )) or (ctran in (1120,1220)) or (ctran in ('220') and rubro in ('6011')))


--TABLA DE CONTEOS DEL INCIDENTE
INSERT INTO @Totales
	SELECT organizacion,
	--Flujos Totales
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@Final_Incidente) THEN 1 END) TOTAL,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@Final_Incidente) and right(posjnlaccess,1)='9' THEN 1 END) as TBK,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@Final_Incidente) and right(posjnlaccess,2)='10' THEN 1 END) as RBC,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@Final_Incidente) and right(posjnlaccess,2)='13' THEN 1 END) as VISA,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@Final_Incidente) and right(posjnlaccess,2)='12' THEN 1 END) as MC,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@Final_Incidente) and right(posjnlaccess,2)='11' THEN 1 END) as AMEX,
	--SAFS Totales
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@T_SAFS) and ( (ctran = '120' and ( charindex(',102',bloqueo)>0 or charindex(',111',bloqueo)>0  or charindex(',113',bloqueo)>0 or charindex(',109',bloqueo)>0 )) or (ctran = '220' and charindex(',p,',bloqueo)>0) or (ctran = '120' and ( charindex(',9001',bloqueo)>0 or charindex(',9020',bloqueo)>0  or charindex(',9011',bloqueo)>0 )) or (ctran in (1120,1220)) or (ctran in ('220') and rubro in ('6011'))) THEN 1 END) TOTAL_S,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@T_SAFS) and ctran = '220' and charindex(',p,',bloqueo)>0 THEN 1 END) as TBK_S,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@T_SAFS) and ctran in ('220') and rubro in ('6011') THEN 1 END) as RBC_S,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@T_SAFS) and ctran = '120' and ( charindex(',9001',bloqueo)>0 or charindex(',9020',bloqueo)>0  or charindex(',9011',bloqueo)>0) THEN 1 END) as VISA_S,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@T_SAFS) and ctran = '120' and ( charindex(',102',bloqueo)>0 or charindex(',111',bloqueo)>0  or charindex(',113',bloqueo)>0 or charindex(',109',bloqueo)>0 ) THEN 1 END) as MC_S,
	count(CASE WHEN (tiempoauth>=@Inicio_Incidente and tiempoauth<@T_SAFS) and ctran in (1120,1220) THEN 1 END) as AMEX_S

	from IsoNxs.dbo.journal (nolock)
	where tiempoauth>=@Inicio_Incidente and tiempoauth<@T_SAFS
	GROUP BY organizacion

--DATA TOTAL
select * from @DATA

--INFORMACION SAFS
select *
from @Totales
order by organizacion desc

--BUSQUEDA DE REINTENTOS PARA TARJETAS
SELECT T.organizacion,
count(*) Tarjetas,
count(CASE WHEN A.pan is not NULL THEN 1 END) Reintentos

FROM 
		(
			SELECT tiempoauth,pan,organizacion
			FROM (
				SELECT  *, ROW_NUMBER() OVER(PARTITION BY Pan ORDER BY pan DESC) rn 
				FROM @DATA
				) a 
			WHERE rn = 1
		) as T
left join
		(
			select tiempoauth,pan
			from IsoNxs.dbo.journal (nolock)
			where tiempoauth>=@Inicio_Incidente and tiempoauth<@T_Reintento
			and ctran in ('0100','0200')
		) as A
ON T.pan=A.pan and T.tiempoauth<A.tiempoauth

GROUP BY T.organizacion