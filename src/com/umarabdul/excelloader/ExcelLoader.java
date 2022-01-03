package com.umarabdul.excelloader;

import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 *   A Java class for loading simple excel files using Apache POI.
 * With the help of theFormulaEvaluator class, cells whose values where
 * generated using a formula are dynamically resolved at runtime.
 * 
 * @author Umar Abdul
 * @version 1.0
 * Date: 27/Dec/2021
 */

public class ExcelLoader {
  
  HashMap<String, ArrayList<String>> records;
  private File excelFile;
  private int sheetIndex;
  private int colsCount;
  private int rowsCount;

  /**
   * Constructor. Does not parse the excel document.
   * @param filename The excel file to load.
   * @param sheetIndex The sheet index to parse.
   */
  public ExcelLoader(File filename, int sheetIndex){

    excelFile = filename;
    this.sheetIndex = sheetIndex;
    records = new HashMap<String, ArrayList<String>>();
  }

  /**
   * Actually parses the excel file.
   * @return {@code true} on success.
   * @throws IOException on I/O error.
   */
  public boolean parse() throws IOException{

    // First, load all cell values in a grid format.
    ArrayList<ArrayList<String>> grid = new ArrayList<ArrayList<String>>();
    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(excelFile));
    XSSFSheet sheet = wb.getSheetAt(sheetIndex);
    Iterator<Row> rows = sheet.iterator();
    FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
    while (rows.hasNext()){
      Row row = rows.next();
      Iterator<Cell> columns = row.cellIterator();
      ArrayList<String> data = new ArrayList<String>();
      while (columns.hasNext()){
        Cell cell = columns.next();
        String value = new DataFormatter().formatCellValue(cell, evaluator);
        if (value != null && value.length() != 0)
          data.add(value);
      }
      if (data.size() > 0)
        grid.add(data);
    }
    wb.close();
    // Now use the first row as column names, and map values in columns to their respective column names.
    if (grid.size() < 2) // At least 2 rows are required!
      return false;
    records.clear();
    // Set dimensions.
    rowsCount = grid.size() - 1;
    colsCount = grid.get(0).size();
    // Load the grid into a mapping of column names to values.
    ArrayList<String> columnNames = grid.get(0);
    for (String name : columnNames)
      records.put(name, new ArrayList<String>());
    for (int i = 1; i < grid.size(); i++){ // Loop over the remaining rows.
      ArrayList<String> row = grid.get(i);
      for (int j = 0; j < row.size(); j++)
        records.get(columnNames.get(j)).add(row.get(j));
    }
    // Done.
    return true;
  }

  /**
   * Get the number of rows loaded.
   * @return The number of rows (excluding the column names row).
   */
  public int getRowsCount(){
    return rowsCount;
  }

  /**
   * Get the number of columns loaded.
   * @return The number of columns loaded.
   */
  public int getColsCount(){
    return colsCount;
  }

  /**
   * Get an ArrayList containing name of columns loaded.
   * @return An ArrayList of column names.
   */
  public ArrayList<String> getColumnNames(){

    ArrayList<String> colNames = new ArrayList<String>();
    for (String name : records.keySet())
    colNames.add(name);
    return colNames;
  }

  /**
   * Return all the values in a column with the given name.
   * @param name The name of the column.
   * @return An ArrayList of the values under the column.
   */
  public ArrayList<String> getColumn(String name){
    return records.get(name);
  }

  /**
   * A simpler method of parsing an excel file.
   * Returns all cell values (including emtpy ones) in the given range.
   * @param excelFile The excel file to read from.
   * @param sheetIndex The index number of the target sheet.
   * @param topLeftAddr Cell address (inclusive) of the upper left cell to process, e.g: A1
   * @param bottomRightAddr Cell address (inclusive) of the bottom right cell to process, e.g: Z10
   * @throws IOException on I/O error.
   * @return A 2D ArrayList of loaded data (with emtpy cells set to null).
   */
  public static ArrayList<ArrayList<String>> slice(File excelFile, int sheetIndex, String topLeftAddr, String bottomRightAddr) throws IOException{

    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(excelFile));
    XSSFSheet sheet = wb.getSheetAt(sheetIndex);
    FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
    CellAddress startAddr = new CellAddress(topLeftAddr);
    CellAddress stopAddr = new CellAddress(bottomRightAddr);
    int startRow = startAddr.getRow();
    int startCol = startAddr.getColumn();
    int endRow = stopAddr.getRow();
    int endCol = stopAddr.getColumn();
    ArrayList<ArrayList<String>> data = new ArrayList<ArrayList<String>>();
    for (int currRow = startRow; currRow <= endRow; currRow++){
      ArrayList<String> rowData = new ArrayList<String>();
      for (int currCol = startCol; currCol <= endCol; currCol++){
        try{
          Cell cell = sheet.getRow(currRow).getCell(currCol);
          rowData.add((cell == null) ? null : new DataFormatter().formatCellValue(cell, evaluator));
        }catch(NullPointerException e){ // Happens if sheet.getRow() fails.
          rowData.add(null);
        }
      }
      data.add(rowData);
    }
    return data;
  }

}
