import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.nio.charset.Charset;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

/**
 * Created by Константин on 11.09.2016.
 */
public class Main {
    private static final String mainDir = "C:/Users/Константин/IdeaProjects/MtsData/";
    private static final String mtsReadScriptName = "MTS_read.exe";
    private static final String out = "out";
    private static final long startDateUnixTime = 315964800; //6 января 1980 года в Unixtime
    private static final Date startDate = new Date(startDateUnixTime);
    private static final long secondsInWeek = 604800; //число секунд в неделе
    private static final long secondsInHour = 3600;
    private static final DateFormat df = new SimpleDateFormat("dd/MM/yyyy HH:mm");
    private static final int[] means = {5, 10};
    private static int cellCount;

    public static void main(String[] args) throws IOException, ParseException {
        df.setTimeZone(TimeZone.getTimeZone("GMT"));
        File folder = new File(mainDir);
        File[] listOfFiles = folder.listFiles();

        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile()) {
                System.out.println("File: " + listOfFiles[i].getName());
            } else if (listOfFiles[i].isDirectory()) {
                System.out.println("Directory: " + listOfFiles[i].getName());
                if (listOfFiles[i].getName().startsWith("MTS")) {
                    File mtsDataFolder = new File(mainDir + listOfFiles[i].getName());
                    File[] dataFilesInMtsFolder = mtsDataFolder.listFiles();
                    File outDirectory = new File(mainDir + out);
                    if (!outDirectory.exists()) {
                        outDirectory.mkdir();//create "out" directory
                    }

                    File outMtsDataFolder = new File(mainDir + out + "/" + listOfFiles[i].getName() + "_" + out);
                    if (!outMtsDataFolder.exists()) {
                        outMtsDataFolder.mkdir();
                    }

                    for (int j=0; j < dataFilesInMtsFolder.length; j++){
                        String cmd = "cd " + mainDir + " && " + mtsReadScriptName + " "
                                + "./" + listOfFiles[i].getName() + "/" + dataFilesInMtsFolder[j].getName()
                                + " ./" + out + "/" + outMtsDataFolder.getName() + "/" + dataFilesInMtsFolder[j].getName().split("\\.")[0] + ".txt";
                        runCommand(cmd);
                    }
                    File[] txtFilesToXls = new File(mainDir + out + "/" + outMtsDataFolder.getName()).listFiles();
                    writeToExcel(mtsDataFolder.getName(), txtFilesToXls, mainDir + out + "/xls/" + mtsDataFolder.getName() + ".xls");
                }
            }
        }
    }

    private static void runCommand(String cmd) throws IOException {
        ProcessBuilder builder = new ProcessBuilder(
                "cmd.exe", "/c", cmd);
        builder.redirectErrorStream(true);
        Process p = builder.start();
        BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));
        String line;
        while (true) {
            line = r.readLine();
            if (line == null) {
                break;
            }
            //System.out.println(line);
        }
    }

    private static void writeToExcel(String folderName, File[] dataFiles, String outFileName) throws IOException, ParseException {

        int rowPosition = 1;

        File outFileDirectory = new File(mainDir + out + "/xls");
        if (!outFileDirectory.exists()) {
            outFileDirectory.mkdir();
        }

        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet(folderName);
        Row row;
        Row header = sheet.createRow(0);
        cellCount = 0;
        for (int meanFactor : means) {
            Cell date = header.createCell(cellCount);
            date.setCellValue("Дата");

            Cell magneticXTitle = header.createCell(1 + cellCount);
            magneticXTitle.setCellValue("magneticX");

            Cell magneticYTitle = header.createCell(2 + cellCount);
            magneticYTitle.setCellValue("magneticY");

            Cell magneticZTitle = header.createCell(3 + cellCount);
            magneticZTitle.setCellValue("magneticZ");

            header.createCell(4 + cellCount);

            Cell telluricXTitle = header.createCell(5 + cellCount);
            telluricXTitle.setCellValue("telluricX");

            Cell telluricYTitle = header.createCell(6 + cellCount);
            telluricYTitle.setCellValue("telluricY");

            Cell telluricZTitle = header.createCell(7 + cellCount);
            telluricZTitle.setCellValue("telluricZ");

            header.createCell(8 + cellCount);

            Cell seismicXTitle = header.createCell(9 + cellCount);
            seismicXTitle.setCellValue("seismicX");

            Cell seismicYTitle = header.createCell(10 + cellCount);
            seismicYTitle.setCellValue("seismicY");

            Cell seismicZTitle = header.createCell(11 + cellCount);
            seismicZTitle.setCellValue("seismicZ");

            for (File dataFile : dataFiles) {
                ArrayList<String> lines = new ArrayList<>();
                InputStream fis = new FileInputStream(dataFile.getAbsolutePath());
                InputStreamReader isr = new InputStreamReader(fis, Charset.forName("UTF-8"));
                BufferedReader br = new BufferedReader(isr);
                String line;
                while ((line = br.readLine()) != null) {//цикл по каждой строке в файле
                    lines.add(line);
                }
                String dateString = getDateFromWeekCount(folderName, dataFile.getName());
                switch (meanFactor) {
                    case 5: {
                        row = sheet.createRow(rowPosition);
                        row.createCell(cellCount).setCellValue(dateString);
                        rowPosition = writeMean(sheet, dateString, lines, rowPosition, 15000);//усреднение по 5 минутам.
                        break;
                    }
                    case 10: {
                        row = sheet.getRow(rowPosition);
                        row.createCell(cellCount).setCellValue(dateString);
                        rowPosition = writeMean(sheet, dateString, lines, rowPosition, 30000);//усреднение по 10 минутам.
                        break;
                    }
                }
            }
            rowPosition = 1;
            sheet.autoSizeColumn(cellCount);
            cellCount += 14;
        }
        File outFile = new File(outFileName);

        book.write(new FileOutputStream(outFile));
        book.close();
    }

    private static int writeMean(Sheet sheet, String dateString, ArrayList<String> lines, int rowPosition, int meanFactor) throws ParseException {

        double magneticXSum = 0;
        double magneticYSum = 0;
        double magneticZSum = 0;
        double telluricXSum = 0;
        double telluricYSum = 0;
        double telluricZSum = 0;
        double seismicXSum = 0;
        double seismicYSum = 0;
        double seismicZSum = 0;

        int lineCount = 1;//счетчик строк в исходном файле с данными
        int totalLineCount = rowPosition;//счетчик строк в ексель файле
        int meanCount = 0;//счетчик усреднений
        Row row;
        for (String line : lines) {
            String[] values = line.split("\\s+");
            magneticXSum += (Double.parseDouble(values[0]) / 3727);
            magneticYSum += (Double.parseDouble(values[1]) / 3593);
            magneticZSum += (Double.parseDouble(values[2]) / 3640);

            telluricXSum += (Double.parseDouble(values[3]) / 30003);
            telluricYSum += (Double.parseDouble(values[4]) / 29944);
            telluricZSum += (Double.parseDouble(values[5]) / 30021);

            seismicXSum += (Double.parseDouble(values[6]));
            seismicYSum += (Double.parseDouble(values[7]));
            seismicZSum += (Double.parseDouble(values[8]) / 139000);
            if (lineCount == meanFactor) {
                row = sheet.getRow(totalLineCount);
                if (row == null){
                    row = sheet.createRow(totalLineCount);
                }

                Date date = df.parse(dateString);
                Calendar calendar = Calendar.getInstance();
                calendar.setTimeInMillis(date.getTime() + meanCount * getMinuteFromMeanFactor(meanFactor) * 60 * 1000);
                row.createCell(cellCount).setCellValue(df.format(calendar.getTime()));

                Cell magneticX = row.createCell(1 + cellCount);
                magneticX.setCellValue((magneticXSum + Double.parseDouble(values[0]) / 3727) / meanFactor);
                Cell magneticY = row.createCell(2 + cellCount);
                magneticY.setCellValue((magneticYSum + Double.parseDouble(values[1]) / 3593) / meanFactor);
                Cell magneticZ = row.createCell(3 + cellCount);
                magneticZ.setCellValue((magneticZSum + Double.parseDouble(values[2]) / 3640) / meanFactor);

                row.createCell(4 + cellCount);

                Cell telluricX = row.createCell(5 + cellCount);
                telluricX.setCellValue((telluricXSum + Double.parseDouble(values[3]) / 30003) / meanFactor);
                Cell telluricY = row.createCell(6 + cellCount);
                telluricY.setCellValue((telluricYSum + Double.parseDouble(values[4]) / 29944) / meanFactor);
                Cell telluricZ = row.createCell(7 + cellCount);
                telluricZ.setCellValue((telluricZSum + Double.parseDouble(values[5]) / 30021) / meanFactor);

                row.createCell(8 + cellCount);

                Cell seismicX = row.createCell(9 + cellCount);
                seismicX.setCellValue(seismicXSum + Double.parseDouble(values[6]));
                Cell seismicY = row.createCell(10 + cellCount);
                seismicY.setCellValue(seismicYSum + Double.parseDouble(values[7]));
                Cell seismicZ = row.createCell(11 + cellCount);
                seismicZ.setCellValue((seismicZSum + Double.parseDouble(values[8]) / 139000) / meanFactor);
                magneticXSum = 0;
                magneticYSum = 0;
                magneticZSum = 0;

                telluricXSum = 0;
                telluricYSum = 0;
                telluricZSum = 0;

                seismicXSum = 0;
                seismicYSum = 0;
                seismicZSum = 0;

                lineCount = 0;
                totalLineCount++;
                meanCount++;
            }
            lineCount++;
        }
        return totalLineCount;
    }

    private static String getDateFromWeekCount(String folderName, String fileName) {
        int weekCount = Integer.parseInt(folderName.split("\\.")[0].split("MTS")[1]);
        int hourInWeek = Integer.parseInt(fileName.split("\\.")[0].split("HOUR")[1]);
        long weekInUnixTime = weekCount * secondsInWeek + hourInWeek * secondsInHour + startDateUnixTime;
        Calendar calendar = Calendar.getInstance();
        calendar.setTimeInMillis(weekInUnixTime * 1000);
        return df.format(calendar.getTime());
    }

    private static int getMinuteFromMeanFactor(int meanFactor) {
        int minute;
        switch (meanFactor) {
            case 15000: {
                minute = 5;
                break;
            }
            case 30000: {
                minute = 10;
                break;
            }
            default: {
                minute = 0;
                break;
            }
        }
        return minute;
    }
}
