WITH sub AS (
SELECT DISTINCT 
c.FirstName + c.LastName + c.Birth_Date__c AS 'ID', 
CASE 
	WHEN c.Gender__c = 'Man/Boy' THEN 'Male'
	WHEN c.Gender__c = 'Woman/Girl' THEN 'Female'
	WHEN c.Gender__c = 'A gender identity not listed' THEN 'Other'
	WHEN c.Gender__c = 'Declined to answer' THEN 'Unknown'
	WHEN c.Gender__c = '' THEN IIF(c.Sex_at_birth__c ='', 'Unknown', c.Sex_at_birth__c)
	ELSE c.Gender__c
END AS GENDER, 
c.Language_Preference__c AS 'PREFERRED LANGUAGE', 
CASE
	WHEN c.Ethinicity__c = '' THEN 'Unknown'
	WHEN c.Ethinicity__c LIKE '%Another group or groups. Please Specify%' THEN IIF(c.Specify_Ethinicity__c ='', 'Unknown', c.Specify_Ethinicity__c)
	WHEN c.Ethinicity__c LIKE '%I do not identify as any specific ethnic or cultural group%' THEN 'Unknown'
	WHEN c.Ethinicity__c LIKE '%Declined to answer%' THEN 'Unknown'
	ELSE c.Ethinicity__c
	END AS ETHNICITY,
IIF(c.Birth_Date__c='', '', year(GETDATE()) - year(c.Birth_Date__c)) AS 'AGE',
IIF(c.OtherPostalCode != '' AND c.OtherPostalCode != c.MailingPostalCode, c.OtherPostalCode, c.MailingPostalCode) AS [ZIP],
c.MailingPostalCode AS 'Permanent Zip',
c.OtherPostalCode AS 'Current Zip',
CASE 
	WHEN c.MailingPostalCode IN ('10457','10458','10459','10460','10462','10472','10475','10474','10473','10471','10470','10469','10468','10467',
								 '10466','10465','10464','10463','10461','10456','10455','10454','10453','10452','10451') THEN 'Bronx'
	WHEN c.MailingPostalCode IN ('10001','10002','10003','10004','10005','10006','10007','10009','10010','10011','10012','10013','10014','10015',
								 '10016','10017','10018','10019','10020','10021','10022','10023','10024','10025','10026','10027','10028','10029',
								 '10030','10031','10032','10033','10034','10035','10036','10037','10038','10039','10040','10041','10044','10045',
								 '10048','10055','10060','10069','10090','10095','10098','10099','10103','10104','10105','10106','10107','10110',
								 '10111','10112','10115','10118','10119','10120','10121','10122','10123','10128','10151','10152','10153','10154',
								 '10155','10158','10161','10162','10165','10166','10167','10168','10169','10170','10171','10172','10173','10174',
								 '10175','10176','10177','10178','10199','10270','10271','10278','10279','10280','10281','10282') THEN 'Manhattan'
	WHEN c.MailingPostalCode IN ('11201','11203','11204','11205','11206','11207','11208','11209','11210','11211','11212','11213','11214','11215',
								 '11216','11217','11218','11219','11220','11221','11222','11223','11224','11225','11226','11228','11229','11230',
								 '11231','11232','11233','11234','11235','11236','11237','11238','11239','11241','11242','11243','11249','11252',
								 '11256') THEN 'Brooklyn' 
	WHEN c.MailingPostalCode IN ('11004','11101','11102','11103','11104','11105','11106','11109','11351','11354','11355','11356','11357','11358',
								 '11359','11360','11361','11362','11363','11364','11365','11366','11367','11368','11369','11370','11371','11372',
								 '11373','11374','11375','11377','11378','11379','11385','11411','11412','11413','11414','11415','11416','11417',
								 '11418','11419','11420','11421','11422','11423','11426','11427','11428','11429','11430','11432','11433','11434',
								 '11435','11436','11691','11692','11693','11694','11697') THEN 'Queens'
	WHEN c.MailingPostalCode IN ('10301','10302','10303','10304','10305','10306','10307','10308','10309','10310','10311','10312','10314') THEN 'Staten Island'
	ELSE c.Permanent_Address_Borough__c END AS 'Permanent Borough',
CASE 
	WHEN c.OtherPostalCode IN ('10457','10458','10459','10460','10462','10472','10475','10474','10473','10471','10470','10469','10468','10467',
							   '10466','10465','10464','10463','10461','10456','10455','10454','10453','10452','10451') THEN 'Bronx'
	WHEN c.OtherPostalCode IN ('10001','10002','10003','10004','10005','10006','10007','10009','10010','10011','10012','10013','10014','10015',
							   '10016','10017','10018','10019','10020','10021','10022','10023','10024','10025','10026','10027','10028','10029',
							   '10030','10031','10032','10033','10034','10035','10036','10037','10038','10039','10040','10041','10044','10045',
							   '10048','10055','10060','10069','10090','10095','10098','10099','10103','10104','10105','10106','10107','10110',
							   '10111','10112','10115','10118','10119','10120','10121','10122','10123','10128','10151','10152','10153','10154',
							   '10155','10158','10161','10162','10165','10166','10167','10168','10169','10170','10171','10172','10173','10174',
							   '10175','10176','10177','10178','10199','10270','10271','10278','10279','10280','10281','10282') THEN 'Manhattan'
	WHEN c.OtherPostalCode IN ('11201','11203','11204','11205','11206','11207','11208','11209','11210','11211','11212','11213','11214','11215',
							   '11216','11217','11218','11219','11220','11221','11222','11223','11224','11225','11226','11228','11229','11230',
							   '11231','11232','11233','11234','11235','11236','11237','11238','11239','11241','11242','11243','11249','11252',
							   '11256') THEN 'Brooklyn'
	WHEN c.OtherPostalCode IN ('11004','11101','11102','11103','11104','11105','11106','11109','11351','11354','11355','11356','11357','11358',
							   '11359','11360','11361','11362','11363','11364','11365','11366','11367','11368','11369','11370','11371','11372',
							   '11373','11374','11375','11377','11378','11379','11385','11411','11412','11413','11414','11415','11416','11417',
							   '11418','11419','11420','11421','11422','11423','11426','11427','11428','11429','11430','11432','11433','11434',
							   '11435','11436','11691','11692','11693','11694','11697') THEN 'Queens'
	WHEN c.OtherPostalCode IN ('10301','10302','10303','10304','10305','10306','10307','10308','10309','10310','10311','10312','10314') THEN 'Staten Island'
	ELSE c.Permanent_Address_Borough__c END AS 'Current Borough',
	
r.Contact_Intake_Completed_Date__c AS 'Created Date',	
r.Receive_At_Home_Testing__c AS 'AT HOME TESTING',
CASE WHEN t.race_trunc = '' THEN 'Missing' ELSE  t.race_trunc END AS 'Race'
	
FROM ICS_COVID19_TRACE_DATASHARE.sf2.Contact c
LEFT JOIN ICS_COVID19_TRACE_DATASHARE.sf2.RESULT r ON r.SALESFORCE_CONTACT_ID = c.Salesforce_Id
LEFT JOIN ICS_COVID19_TRACE_DATASHARE.sf2.Interaction I ON  I.Salesforce_id = r.SALESFORCE_CONTACT_ID
LEFT JOIN ICS_COVID19_TRACE.sf2.contact_recode_new t ON t.salesforce_id = I.Salesforce_Id

WHERE 
MONTH(r.Contact_Intake_Completed_Date__c) = DATEPART(MONTH, DATEADD(MONTH, -1, GETDATE())) AND
YEAR(r.Contact_Intake_Completed_Date__c) = YEAR(GETDATE()) AND
--CAST(r.Contact_Intake_Completed_Date__c AS Date) >= '2021-03-01' AND 
I.RecordTypeId = 'Contact Intake' AND 
r.RecordType = 'Exposed'

)

SELECT DISTINCT [ID], [GENDER], [PREFERRED LANGUAGE], [ETHNICITY], [Race], [AGE], [ZIP], 
IIF([Current Zip]!= '' AND [Current Zip] != [Permanent Zip], [Current Borough], [Permanent Borough]) AS [BOROUGH], 
[AT HOME TESTING], [Created Date]
FROM sub
