UPDATE T1
SET 
T1.LCHDATE = T2.LCHDATE
FROM [ivy.mm.dim.posupc] T1 
INNER JOIN (
    SELECT UPC, MIN(ACT_DATE) AS LCHDATE
    FROM [dbo].[ivy.sd.fact.pos]
    GROUP BY UPC
	) AS T2 ON T1.upc = T2.upc
WHERE T1.upc = T2.upc;

UPDATE T1
SET 
T1.[description] = T2.fdesc
FROM [ivy.mm.dim.posupc] AS T1 
	INNER JOIN cr_upc_desc AS T2 ON T1.upc = T2.upc

UPDATE T1
SET 
T1.description = T2.description,
T1.kiss = T2.material,
T1.brand = T2.brand,
T1.division = T2.division,
T1.mg = T2.mg,
T1.company='KISS'
FROM [ivy.mm.dim.posupc] AS T1 
INNER JOIN (
	SELECT*
    FROM [ivy.mm.dim.mtrl]
    WHERE UPC IN (
		SELECT UPC
		FROM [ivy.mm.dim.mtrl]
		GROUP BY UPC
		HAVING COUNT(*)=1)
		) AS T2 ON T1.upc = T2.upc
WHERE T1.upc = T2.upc;
