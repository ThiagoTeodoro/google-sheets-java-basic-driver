# google-sheets-java-basic-driver
Driver for connection and operation with google sheets. 

To use this driver, it is necessary to activate the google sheets api in the account where the driver will operate, when activated you will get a credentials.json file that must be placed in the src/main/resource folder.

In the first use you will receive a link that when you click it will take you to Google's oauth2 authentication page, so just upload and proceed.

Usage example


# Driver initialization

GoogleSheetsService googleSheetsService = GoogleSheetsService.initializer(
				"PATH_CREDENTIALS",
				"USER_TO_DISPLAY",
				"tokens",
				"APPLICATION_NAME",
				"ID_SPREEDSHEET");


# To update data into a specific range

googleSheetsService.updateValueInSheets(Arrays.asList("VALUE A1", "VALUE B1)"), "A1:B1");

# To read data in a Range.

String range = "NAME_SHEET!A1:B1";
List<List<Object>> data = googleSheetsService.readValueInSheets(range);

if(dados != null) {
	for(List<Object> row : data) {
		log.info(String.format("VALUE IN A1 : %s VALUE IN B1 : %s.", row.get(0), row.get(1)));
	}
}

# To insert a row in end of file

googleSheetsService.insertRowInSheets(Arrays.asList("VALUE A1", "VALUE B1", "VALUE B2"), "NAME_SHEET");


# To delete Row 

googleSheetsService.deleteRowsInSheets(INT ROW START, INT ROW END, ID SHEET);

