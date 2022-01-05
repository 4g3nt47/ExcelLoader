# ExcelLoader

ExcelLoader is a simple Java library I made to make reading data from excel files easier. It uses the powerful Apache POI library to load and process excel files. Cells whose values are dynamically generated using a formula are resolved at runtime.

## Installation

* `$ git clone https://github.com/UmarAbdul01/ExcelLoader.git`
* `$ cd ExcelLoader`
* Copy `excelloader.jar` to your classpath.

## Build from source

* `$ git clone https://github.com/UmarAbdul01/ExcelLoader.git`
* `$ cd ExcelLoader`
* Edit `build.xml` and replace `/opt/java/lib` with your classpath.
* `$ ant jar`
* Copy `excelloader.jar` to your classpath.

## Demo

```java
// Import
import com.umarabdul.excelloader.ExcelLoader;

// A simple demo of using the library.
ExcelLoader ld = new ExcelLoader(new File("sample.xlsx"), 0);
if (ld.parse()){
  System.out.println(String.format("Number of cols: %d\nNumber of rows: %d", ld.getColsCount(), ld.getRowsCount()));
  for (String col: ld.getColumnNames()){
    System.out.println("Column: " + col);
    System.out.println("Values: ");
    for (String val : ld.getColumn(col))
      System.out.print("\t" + val);
    System.out.println("");
  }
}else{
  System.out.println("Error loading excel document!");
}

// A second demo that uses the static ExcelLoader.slice() method to parse a file.
ArrayList<ArrayList<String>> data = ExcelLoader.slice(new File("sample.xlsx"), 0, "A1", "e7");
for (int row = 0; row < data.size(); row++){
  System.out.print(String.format("\nRow %d:  ", row));
  ArrayList<String> rowData = data.get(row);
  for (int col = 0; col < rowData.size(); col++){
    String val = rowData.get(col);
    if (val == null)
      val = "[null]";
    System.out.print(String.format("%20s", val));
  }
}
System.out.println("");
```
