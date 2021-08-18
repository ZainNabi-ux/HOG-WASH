	DECLARE @Sql VARCHAR(255)
	
	SET @Sql =  ' INSERT INTO CRNoteTemp '									+
				' SELECT * '												+
				' FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', '			+
				' ''Excel 12.0; HDR=NO; Database='+ @FilePath + ';'' ,'   +
				' ''SELECT* FROM ['+ @CreditSheetName + @CreditSheetColumns +']'')' 

	EXECUTE sp_executesql @Sql



SET @Sql =  ' BULK INSERT  ##tmp_RepTargets ' +
			' ''FROM ['+ @FileName  +']' +
			' WITH ' +
			'  (   ' +
			'    FIRSTROW = 2, ' +
			' FIELDTERMINATOR = '','  +
			' ROWTERMINATOR = ''\n,'   +
			' TABLOCK ' +
			' ) ' 
