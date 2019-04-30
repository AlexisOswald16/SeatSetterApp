package com.example.seatsetterapp;

import android.Manifest;
import android.os.Build;
import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.content.Intent;
import android.net.Uri;
import android.util.Log;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;


public class FileChooserActivity extends AppCompatActivity {

    //Declaration of variables
    private static final String TAG = "MainActivity";
    List<String> allNames = new ArrayList<>();

    ArrayList<String> names = new ArrayList<>(); //ArrayList that contains all names from excel file

    private void readExcelData(Uri selectedFile) {
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
                        String value = getCellAsString(row, c, formulaEvaluator);
                        sb.append(value + ", ");
                    }
                }
            }

            Log.d(TAG, "ReadExcelData: STRINGBUILDER: " + sb.toString().replace(";",""));
            names.add(sb.toString().replace(";","")); // where the names are added into the arraylist
            Log.d(TAG, "NAMES HERE " + names.toString());

        } catch (FileNotFoundException e) {
            Log.e(TAG, "readExcelData: FileNotFoundException " + e.getMessage());
        } catch (IOException e) {
            Log.e(TAG, "readExcelData: Error reading inputstream" + e.getMessage());
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
    private void toastMessage(String message) {
        Toast.makeText(this, message, Toast.LENGTH_SHORT).show();
    }

    private List<String> parseArrayList(){
        String all = names.get(0);
        allNames = Arrays.asList(all.split("\\s*,\\s*"));
        return allNames;
    }
    private void generateSeatingChart(){
        parseArrayList();
        ArrayList<Integer> numberList = new ArrayList<>();
        for (int i = 1; i < allNames.size(); i++) {
            numberList.add(i);
        }

        Collections.shuffle(numberList); //generates random numbers to correspond with the names
        setAllTextViews(numberList);

    }

    private void setAllTextViews(ArrayList<Integer> numberList){
        
        TextView T1 = (TextView)findViewById(R.id.T1);
        TextView T2 = (TextView)findViewById(R.id.T2);
        TextView T3 = (TextView)findViewById(R.id.T3);
        TextView T4 = (TextView)findViewById(R.id.T4);
        TextView T5 = (TextView)findViewById(R.id.T5);
        TextView T6 = (TextView)findViewById(R.id.T6);
        TextView T7 = (TextView)findViewById(R.id.T7);
        TextView T8 = (TextView)findViewById(R.id.T8);
        TextView T9 = (TextView)findViewById(R.id.T9);
        TextView T10 = (TextView)findViewById(R.id.T10);
        TextView T11 = (TextView)findViewById(R.id.T11);
        TextView T12 = (TextView)findViewById(R.id.T12);
        TextView T13 = (TextView)findViewById(R.id.T13);
        TextView T14 = (TextView)findViewById(R.id.T14);
        TextView T15 = (TextView)findViewById(R.id.T15);
        TextView T16 = (TextView)findViewById(R.id.T16);
        TextView T17= (TextView)findViewById(R.id.T17);
        TextView T18 = (TextView)findViewById(R.id.T18);
        TextView T19 = (TextView)findViewById(R.id.T19);
        TextView T20 = (TextView)findViewById(R.id.T20);
        TextView T21 = (TextView)findViewById(R.id.T21);
        TextView T22 = (TextView)findViewById(R.id.T22);
        TextView T23 = (TextView)findViewById(R.id.T23);
        TextView T24 = (TextView)findViewById(R.id.T24);
        TextView T25= (TextView)findViewById(R.id.T25);
        TextView T26= (TextView)findViewById(R.id.T26);
        TextView T27= (TextView)findViewById(R.id.T27);
        TextView T28= (TextView)findViewById(R.id.T28);
        TextView T29= (TextView)findViewById(R.id.T29);
        TextView T30= (TextView)findViewById(R.id.T30);

       if(numberList.size() < 30){
           for(int i = numberList.size(); i < 30; i++){
               numberList.add(0);
           }
       }

        T1.setText("   "+ allNames.get(numberList.get(0)) + "   ");
        T2.setText("   "+ allNames.get(numberList.get(1)) + "   ");
        T3.setText("   "+ allNames.get(numberList.get(2)) + "   ");
        T4.setText("   "+ allNames.get(numberList.get(3)) + "   ");
        T5.setText("   "+ allNames.get(numberList.get(4)) + "   ");
        T6.setText("   "+ allNames.get(numberList.get(5)) + "   ");
        T7.setText("   "+ allNames.get(numberList.get(6)) + "   ");
        T8.setText("   "+ allNames.get(numberList.get(7)) + "   ");
        T9.setText("   "+ allNames.get(numberList.get(8)) + "   ");
        T10.setText("   "+ allNames.get(numberList.get(9)) + "   ");
        T11.setText("   "+ allNames.get(numberList.get(10)) + "   ");
        T12.setText("   "+ allNames.get(numberList.get(11)) + "   ");
        T13.setText("   "+ allNames.get(numberList.get(12)) + "   ");
        T14.setText("   "+ allNames.get(numberList.get(13)) + "   ");
        T15.setText("   "+ allNames.get(numberList.get(14)) + "   ");
        T16.setText("   "+ allNames.get(numberList.get(15)) + "   ");
        T17.setText("   "+ allNames.get(numberList.get(16)) + "   ");
        T18.setText("   "+ allNames.get(numberList.get(17)) + "   ");
        T19.setText("   "+ allNames.get(numberList.get(18)) + "   ");
        T20.setText("   "+ allNames.get(numberList.get(19)) + "   ");
        T21.setText("   "+ allNames.get(numberList.get(20)) + "   ");
        T22.setText("   "+ allNames.get(numberList.get(21)) + "   ");
        T23.setText("   "+ allNames.get(numberList.get(22)) + "   ");
        T24.setText("   "+ allNames.get(numberList.get(23)) + "   ");
        T25.setText("   "+ allNames.get(numberList.get(24)) + "   ");
        T26.setText("   "+ allNames.get(numberList.get(25)) + "   ");
        T27.setText("   "+ allNames.get(numberList.get(26)) + "   ");
        T28.setText("   "+ allNames.get(numberList.get(27)) + "   ");
        T29.setText("   "+ allNames.get(numberList.get(28)) + "   ");
        T30.setText("   "+ allNames.get(numberList.get(29)) + "   ");


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
            checkFilePermissions(); //checks permissions
            readExcelData(selectedFile); // reads the excel file
            generateSeatingChart();
        }

    }


}