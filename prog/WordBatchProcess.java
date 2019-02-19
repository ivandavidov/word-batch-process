package prog;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Created by Ivan on 2019-01-16.
 */
public class WordBatchProcess {
    private Map<String, List<String>> placeholders = new HashMap<>();
    private int numDocs = 0;
    private final String PLACEHOLDER_SEPARATOR = "oo";
    private final String RESULT_DIR = "result";
    private final DecimalFormat DECIMAL_FORMAT = new DecimalFormat("#.##");

    public static void main(String... args) throws Exception {
        WordBatchProcess wbc = new WordBatchProcess();
        wbc.reset();
        wbc.preparePlaceholders();
        wbc.processWordFiles();
    }

    private void reset() throws Exception {
        File dir = new File(".");
        String dirName = dir.getCanonicalPath();
        File resultDir = Paths.get(dirName, "result").toFile();
        FileUtils.deleteDirectory(resultDir);
    }

    private void preparePlaceholders() throws Exception {
        File dir = new File(".");

        File[] excelFiles = dir.listFiles((directory, name) -> name.endsWith(".xlsx"));

        if(excelFiles == null || excelFiles.length <= 0) {
            throw new RuntimeException("The current directory contains no MS Excel files. Cannot continue.");
        }

        if (excelFiles.length > 1) {
            throw new RuntimeException("The current directory contains more than one MS Excel file. Cannot continue.");
        }

        File excelPlaceholdersFile = excelFiles[0];
        String excelPlaceholdersFileName = excelPlaceholdersFile.getCanonicalPath();

        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(excelPlaceholdersFileName));
        XSSFSheet sheet = workbook.getSheetAt(0);

        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();

        XSSFRow keysRow = sheet.getRow(firstRowNum);

        int firstCellNum = keysRow.getFirstCellNum();
        int lastCellNum = keysRow.getLastCellNum();
        int rowLength = keysRow.getPhysicalNumberOfCells();

        String[] keys = new String[rowLength];

        for(int i = firstCellNum; i < lastCellNum; i++) {
            XSSFCell cell = keysRow.getCell(i);
            String key = cell.getStringCellValue();
            keys[i - firstCellNum] = key;
        }

        for(int i = firstRowNum + 1; i <= lastRowNum; i++) {
            numDocs++;
            XSSFRow row = sheet.getRow(i);

            firstCellNum = row.getFirstCellNum();
            lastCellNum = row.getLastCellNum();

            for(int j = firstCellNum; j < lastCellNum; j++) {
                XSSFCell cell = row.getCell(j);
                CellType cellType = cell.getCellType();
                String value;
                switch(cellType) {
                    case STRING:
                        value = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        double dval =  cell.getNumericCellValue();
                        value = DECIMAL_FORMAT.format(dval);
                        break;
                    default:
                        value = cell.getRawValue();
                }

                String key = keys[j - firstCellNum];

                List<String> values;
                if(placeholders.containsKey(key)) {
                    values = placeholders.get(key);
                    if(values == null) {
                        values = new ArrayList<>();
                    }
                } else {
                    values = new ArrayList<>();
                    placeholders.put(key, values);
                }
                values.add(value);
            }
        }

        workbook.close();
    }

    private void processWordFiles() throws Exception {
        File dir = new File(".");

        File[] wordFiles = dir.listFiles((directory, name) -> name.endsWith(".docx"));

        if(wordFiles == null || wordFiles.length <= 0) {
            throw new RuntimeException("The current directory contains no MS Word files. Cannot continue.");
        }

        if (wordFiles.length > 1) {
            throw new RuntimeException("The current directory contains more than one MS Word file. Cannot continue.");
        }

        Paths.get(".", RESULT_DIR).toFile().mkdirs();

        File wordTemplateFile = wordFiles[0];

        Set<String> keys = placeholders.keySet();

        for(int i = 0; i < numDocs; i++) {
            File backupFile = new File("backup.bck");
            FileUtils.copyFile(wordTemplateFile, backupFile);

            XWPFDocument wordTemplate = new XWPFDocument(OPCPackage.open(backupFile));
            List<XWPFParagraph> paragraphs = wordTemplate.getParagraphs();

            for (XWPFParagraph paragraph : paragraphs) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null) {
                        if(i == 0) {
                            System.out.println(text);
                        }

                        for (String key : keys) {
                            List<String> values = placeholders.get(key);
                            String placeholder = PLACEHOLDER_SEPARATOR + key + PLACEHOLDER_SEPARATOR;
                            text = text.replace(placeholder, values.get(i));
                            run.setText(text, 0);
                        }
                    }
                }
            }

            String outFileName = Paths.get(".", RESULT_DIR, wordTemplateFile.getName()).toFile().getCanonicalPath();

            for(String key : keys) {
                List<String> values = placeholders.get(key);
                String placeholder = PLACEHOLDER_SEPARATOR + key + PLACEHOLDER_SEPARATOR;
                outFileName = outFileName.replace(placeholder, values.get(i));
            }

            wordTemplate.write(new FileOutputStream(outFileName));
            wordTemplate.close();

            backupFile.delete();
        }
    }
}
