package com.excel.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.excel.util.Utility;
import com.excel.vo.InputVO;
import com.fasterxml.jackson.databind.ObjectMapper;

@RestController
@RequestMapping("/api/excel-to-json")
public class ExcelController {

	@PostMapping
	public ResponseEntity<Object> process(@RequestBody InputVO inputVO) {

		File outputDir = new File(inputVO.getResult_directory());
		if (!outputDir.exists() || !outputDir.isDirectory()) {
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("output directory does not exist");
		} else {
			if (Utility.getFileExtension(inputVO.getFile_name()).equalsIgnoreCase("xlsx")) {
				File fileToBeProcessed = new File(inputVO.getFile_path() + "/" + inputVO.getFile_name());
				if (fileToBeProcessed.exists()) {
					try {
						process(fileToBeProcessed, inputVO.getResult_directory());
					} catch (IOException ioe) {
						return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("error reading file");
					}
					return ResponseEntity.ok("file exists");
				} else {
					return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("file does not exist");
				}
			} else {
				return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("invalid file extention");
			}
		}
	}

	private void process(File file, String outputPath) throws FileNotFoundException, IOException {
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
		for (int i = 0; i < myWorkBook.getNumberOfSheets(); i++) {
			processSheets(myWorkBook.getSheetAt(i),myWorkBook.getSheetName(i),  outputPath);
		}
	}

	private void processSheets(XSSFSheet sheet,String sheetName,  String outputPath) throws IOException{
		Iterator<Row> rowIterator = sheet.iterator();
		List<String> keys = new ArrayList<>();
		List<List<String>> values = new ArrayList<List<String>>();
		List<Map<String, String>> finalJson = new ArrayList<Map<String,String>>();
		int rowIndex = 1;
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if(rowIndex == 1 ) {
				keys = processHeader(row);
			}else {
				values.add(processRow(row));
			}
			rowIndex++;
		}

		for(int i = 0; i< values.size(); i++) {
			Map<String, String> map = new HashedMap<>();
			for(int j = 0 ; j < keys.size() ; j++) {
				map.put(keys.get(j), values.get(i).get(j));
			}
			finalJson.add(map);
		}
		
		File outputFile = new File(outputPath+"/"+sheetName+".json");
		outputFile.createNewFile();
		writeJsonDataToFile(outputFile,finalJson);
	}
	
	private void writeJsonDataToFile(File file , List<Map<String, String>> data) throws IOException{
		ObjectMapper mapper = new ObjectMapper();
		
		mapper.writerWithDefaultPrettyPrinter().writeValue(file, data);
		
	}
	
	private List<String> processHeader(Row row){
		List<String> header = new ArrayList<>();
		Iterator<Cell> cellIterator = row.cellIterator();
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			switch (cell.getCellType()) {
			case STRING:
				header.add(cell.getStringCellValue());
				break;
			default:
			}
		}
		
		return header;
	}
	
	private List<String> processRow(Row row){
		List<String> excelRow = new ArrayList<>();
		Iterator<Cell> cellIterator = row.cellIterator();
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			switch (cell.getCellType()) {
			case STRING:
				excelRow.add(cell.getStringCellValue());
				break;
			case NUMERIC:
				excelRow.add(cell.getNumericCellValue()+"");
				break;
			case BOOLEAN:
				excelRow.add(cell.getBooleanCellValue() + "");
				break;
				
			default:
			}
		}
		
		return excelRow;
	}
}
