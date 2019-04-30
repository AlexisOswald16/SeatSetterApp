package com.example.seatsetterapp;

import android.Manifest;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.content.Intent;
import android.net.Uri;
import android.util.Log;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.ListView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import android.database.Cursor;
import android.provider.MediaStore;
import android.content.CursorLoader;




public class FileChooserActivity extends AppCompatActivity {

    //Declaration of variables
    private static final String TAG = "MainActivity";

    private String[] FilePathStrings;
    private String[] FileNameStrings;
    private File[] listFile;
    File file;

    Button btnUpDirectory, btnSDCard;

    ArrayList<String> pathHistory;
    String lastDirectory;
    int count = 0;

    ArrayList<XValues> uploadData;      //arraylist<XValues> ??


    ListView lvInternalStorage;

    private void readExcelData(Uri selectedFile) {
        Log.d(TAG, "readExceldata: Reading Excel File.");
        //declare input file
        //File inputFile = new File(filePath);


        try {
            InputStream inputStream = getContentResolver().openInputStream(selectedFile);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowsCount = sheet.getPhysicalNumberOfRows();
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            StringBuilder sb = new StringBuilder();

            //outer loop, loops through rows
            for (int r = 1; r < rowsCount; r++) {
                Row row = sheet.getRow(r);
                int cellsCount = row.getPhysicalNumberOfCells();
                //inner loop, does columns, even though we know we have 2
                for (int c = 0; c < cellsCount; c++) {
                    //if there are too many columns
                    if (c > 2) {
                        Log.e(TAG, "readExcelData: ERROR - Excel file format is incorrect");
                        toastMessage("Error excel file format incorrect");
                        break;
                    } else {
                        String value = getCellAsString(row, c, formulaEvaluator);//writing getcellasstring later
                        String cellInfo = "r:" + r + "; c:" + c + "; v:" + value;
                        Log.d(TAG, "readExcelData: Data from row: " + cellInfo);
                        sb.append(value + ", ");
                    }
                }
                sb.append(";");
            }

            Log.d(TAG, "ReadExcelData: STRINGBUILDER: " + sb.toString());
            parseStringBuilder(sb);

        } catch (FileNotFoundException e) {
            Log.e(TAG, "readExcelData: FileNotFoundException " + e.getMessage());
        } catch (IOException e) {
            Log.e(TAG, "readExcelData: Error reading inputstream" + e.getMessage());
        }
    }

    public void parseStringBuilder(StringBuilder mStringBuilder) {
        Log.d(TAG, "parseStringBuilder: Started parsing ");


        //Add

        //splits sb into rows
        String[] rows = mStringBuilder.toString().split(":");

        //add to arraylist<Xvalues> row by row
        for (int i = 0; i < rows.length; i++) {
            String[] columns = rows[i].split(",");

            //use try catch to make sure no '' that try to parse into doubles
            try {
                String x = (columns[0]);
                String y = (columns[1]);

                String cellInfo = "(x.y): (" + x + "," + y + ")";
                Log.d(TAG, "ParseStringBuilder: Data from row: " + cellInfo);

                //add data to ArrayList
                uploadData.add(new XValues(x, y));
            } catch (NumberFormatException e) {
                Log.e(TAG, "parseStringuilder: NumberFormatException " + e.getMessage());
            }

        }
    }

    private void printDataToLog() {
        Log.d(TAG, "PrintDataToLog: Printing data to log....");

        for (int i = 0; i < uploadData.size(); i++) {
            String x = uploadData.get(i).getX();
            String y = uploadData.get(i).getY();
            Log.d(TAG, "printDataToLog: (x,y): (" + x + "," + y + ")");
        }
    }


    private String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator) {
        String value = "";
        try {
            Cell cell = row.getCell(c);
            CellValue cellValue = formulaEvaluator.evaluate(cell);
            switch (cellValue.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    value = "" + cellValue.getBooleanValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    double numericValue = cellValue.getNumberValue();
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        double date = cellValue.getNumberValue();
                        SimpleDateFormat formatter = new SimpleDateFormat("dd/mm/yy");
                        value = formatter.format(HSSFDateUtil.getJavaDate(date));
                    } else {
                        value = "" + numericValue;
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = "" + cellValue.getStringValue();
                    break;
                default:
            }
        } catch (NullPointerException e) {
            Log.e(TAG, "getCellString: NullPointerException: " + e.getMessage());
        }
        return value;
    }


    private void checkInternalStorage() {
        Log.d(TAG, "checkInternalStorage Started");
        try {
            if (!Environment.getExternalStorageState().equals(Environment.MEDIA_MOUNTED)) {
                toastMessage("No sd card found");
            } else {
                //locate image folder in your sd card
                file = new File((pathHistory.get(count)));
                Log.d(TAG, "checkInternalStorage:directory path: " + pathHistory.get(count));
            }

            listFile = file.listFiles();
            //Create string array for file path strings
            FilePathStrings = new String[listFile.length];

            //string array for file name strings
            FileNameStrings = new String[listFile.length];
            for (int i = 0; i < listFile.length; i++) {//path of image file
                FilePathStrings[i] = listFile[i].getAbsolutePath();
                //get name of image file
                FileNameStrings[i] = listFile[i].getName();
            }

            for (int i = 0; i < listFile.length; i++) {
                Log.d("Files", "FileName: " + listFile[i].getName());
            }
            ArrayAdapter<String> adapter = new ArrayAdapter<String>(this, android.R.layout.simple_expandable_list_item_1, FilePathStrings);
            lvInternalStorage.setAdapter(adapter);

        } catch (NullPointerException e) {
            Log.e(TAG, "checkInternalStorage: NULLPOINTER EXCEPTION " + e.getMessage());
        }
    }

    private void checkFilePermissions() {
        if (Build.VERSION.SDK_INT > Build.VERSION_CODES.LOLLIPOP) {
            int permissionCheck = this.checkSelfPermission("Manifest.permission.READ_EXTERNAL_STORAGE");
            permissionCheck = this.checkSelfPermission("Manifest.permission.WRITE_EXTERNAL_STORAGE");
            if (permissionCheck != 0) {
                this.requestPermissions(new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE, Manifest.permission.READ_EXTERNAL_STORAGE}, 1001);
            } else {
                Log.d(TAG, "checkBTPermissions: no need to check permissions. SDK version < LOLLIPOP.");
            }
        }
    }


    /*
    customizable toast
     */

    private void toastMessage(String message) {
        Toast.makeText(this, message, Toast.LENGTH_SHORT).show();
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.file_chooser_activity);

        Intent intent = new Intent()
                .setType("*/*")
                .setAction(Intent.ACTION_GET_CONTENT);

        startActivityForResult(Intent.createChooser(intent, "Select a file"), 123);
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == 123 && resultCode == RESULT_OK) {
            Uri selectedFile = data.getData(); //The uri with the location of the file
            String absolutePath = getContentResolver().getType(selectedFile);
            checkFilePermissions();

            readExcelData(selectedFile);
        }

    }


}