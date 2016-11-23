package org.m8.toexcel;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;

import java.io.*;
import java.util.ArrayList;
import java.util.Locale;

/**
 * Created by Malintha on 11/13/2016.
 */

public class WriteExcel {
    private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    private String outputFile;

    public void setOutputFile(String outputFile) {
        this.outputFile = outputFile;
    }

    /***
     * read agent log and return an array of data
     *
     * agent name
     * A/B/C...
     * issues
     *
     * @return arrayList of arraylists
     */
    public ArrayList<ArrayList<String>> readAgentLog() {
        File agentLog = new File("C:\\Users\\Malintha\\Desktop\\paper\\data\\sim_grp2_dis_rv_10.txt");
        ArrayList<ArrayList<String>> outputArray = new ArrayList<>();
        try {
            BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(agentLog)));
            String line = br.readLine();
            while(line != null) {
                if (line.contains("sent the following offer")) {
                    ArrayList<String> outputArrayList = new ArrayList();
                    String[] firstLine = line.split(" ");
                    outputArrayList.add(firstLine[1]);
                    outputArrayList.add(String.valueOf(line.substring(line.indexOf("("),line.indexOf(")")).charAt(1)));
                    line = br.readLine();
                    String offer = line.substring(line.indexOf('['),line.indexOf(']'));
                    String[] issues = offer.split(",");
                    for(int i = 0; i<issues.length-1;i++) {
                        outputArrayList.add(issues[i].split(":")[1]);
                    }
                    outputArray.add(outputArrayList);
                }
                line = br.readLine();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return outputArray;
    }

    public void write() throws IOException, WriteException {
        File file = new File(outputFile);
        WorkbookSettings wbSettings = new WorkbookSettings();

        wbSettings.setLocale(new Locale("en", "EN"));

        WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Report", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        createLabel(excelSheet);
        ArrayList<ArrayList<String>> logs = readAgentLog();
        createContent(excelSheet, logs);

        workbook.write();
        workbook.close();
    }

    private void createLabel(WritableSheet sheet)
            throws WriteException {
        // Write headers
        addCaption(sheet, 0, 0, "agentName");
        addCaption(sheet, 1, 0, "agentCode");
        addCaption(sheet, 2, 0, "c1-i10");
        addCaption(sheet, 3, 0, "c1-i9");
        addCaption(sheet, 4, 0, "c1-i8");
        addCaption(sheet, 5, 0, "c1-i7");
        addCaption(sheet, 6, 0, "c1-i6");
        addCaption(sheet, 7, 0, "c1-i5");
        addCaption(sheet, 8, 0, "c1-i4");
        addCaption(sheet, 9, 0, "c1-i3");
        addCaption(sheet, 10, 0, "c1-i2");
        addCaption(sheet, 11, 0, "c1-i1");
    }

    private void createContent(WritableSheet sheet, ArrayList<ArrayList<String>> logs) throws WriteException {
        for(int i =0; i <logs.size(); i++) {
            ArrayList<String> log = logs.get(i);
            for(int j =0; j<log.size();j++) {
                addText(sheet, j, i+1, log.get(j));
            }

        }
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s)
            throws RowsExceededException, WriteException {
        Label label;
        label = new Label(column, row, s);
        sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row,
                           Integer integer) throws WriteException, RowsExceededException {
        Number number;
        number = new Number(column, row, integer, times);
        sheet.addCell(number);
    }

    private void addText(WritableSheet sheet, int column, int row, String s)
            throws WriteException, RowsExceededException {
        Label label;
        label = new Label(column, row, s);
        sheet.addCell(label);
    }

    public static void main(String[] args) throws WriteException, IOException {
        WriteExcel test = new WriteExcel();
        test.setOutputFile("C:\\Users\\Malintha\\Desktop\\paper\\GeniusLogToExcel\\output\\worksheet.xls");
        test.write();
        System.out.println("C:\\Users\\Malintha\\Desktop\\paper\\GeniusLogToExcel\\output\\worksheet.xls");
    }
}
