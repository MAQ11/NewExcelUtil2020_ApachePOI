# NewExcelUtil2020_ApachePOI
## This is latest Apache POI- Excel Util for read/write in Excel

# same code:

''' Xls_Reader reader = new Xls_Reader("src/main/java/com/excel/lib/util/SampleExcel.xlsx");
String data = reader.getCellData("login", "username", 3);
System.out.println(reader.getCellData("login", 1, 2));
System.out.println(data);
	
reader.setCellData("login", "status", 2, "pass");
reader.addSheet("Naveen2"); 
reader.removeColumn("login", 2); '''
