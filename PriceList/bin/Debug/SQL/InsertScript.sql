IF NOT EXISTS (SELECT * FROM [@SML_PRCAUTH] where Code = 'orsan')
BEGIN
	INSERT INTO [@SML_PRCAUTH] (DocEntry,Code) VALUES (1,'orsan');
	
END;
IF NOT EXISTS (SELECT * FROM [@SML_PRCAUTH] where Code = 'manager')
BEGIN
	INSERT INTO [@SML_PRCAUTH] (DocEntry,Code) VALUES (2,'manager');
	
END;

