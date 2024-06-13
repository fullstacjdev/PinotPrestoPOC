import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.concurrent.CountDownLatch;

import static Framework_Methods.Generic_Methods.*;
import static Framework_Methods.Generic_Methods.replaceTextInHTMLReport;

import java.awt.AWTException;
import java.awt.Dimension;
import java.awt.Robot;
import java.awt.Toolkit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import com.microsoft.playwright.Browser;
import com.microsoft.playwright.BrowserContext;
import com.microsoft.playwright.BrowserType;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.Playwright;

public class Framework_Main extends Thread {
    private static CountDownLatch latch;
    String browserName;
    static String ExecutionReportFolderName;
    static String executionDataSheetPathAndName;
    static String ExecutionScriptName;
    static String reportFileName;
    static String excelReportFileName;
    static String htmlReportFileName;
    private static final ThreadLocal<SimpleDateFormat> dateFormatThreadLocal = ThreadLocal.withInitial(() -> new SimpleDateFormat("HH:mm:ss"));

    //Constructor
    public Framework_Main(String browserName) {
        this.browserName = browserName;
        this.latch = latch;
    }
    //Thread run method
    public void run() {
        boolean strMethodsReturnStatus = false;
        long threadId = Thread.currentThread().getId();
        //String executionStartTime = String.valueOf(new SimpleDateFormat("HH:mm:ss").format(new Date()));
        String executionStartTime  = getCurrentTime();
        Playwright playwright = Playwright.create();
        Browser browser = null;
        Page page = null;
        if (browserName.equalsIgnoreCase("Chrome")) {
            page = getBrowser(playwright, browserName).launch(new BrowserType.LaunchOptions().setHeadless(false)
                            .setChannel("chrome").setArgs(List.of("--start-maximized")))
                    .newContext(new Browser.NewContextOptions().setViewportSize(null))
                    .newPage();
        } else if (browserName.equalsIgnoreCase("Firefox")) {
            browser = getBrowser(playwright, browserName).launch(new BrowserType.LaunchOptions().setHeadless(false));
            Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
            int width = (int) screenSize.getWidth();
            int height = (int) screenSize.getHeight();
            BrowserContext browserContext = browser.newContext(new Browser.NewContextOptions().setViewportSize(width - 30, height));
            page = browserContext.newPage();
        } else if (browserName.equalsIgnoreCase("Edge")) {
            page = getBrowser(playwright, browserName).launch(new BrowserType.LaunchOptions().setHeadless(false)
                            .setChannel("msedge").setArgs(List.of("--start-maximized")))
                    .newContext(new Browser.NewContextOptions().setViewportSize(null))
                    .newPage();
        } else if (browserName.equalsIgnoreCase("Chrome_Headless")) {
            browser = getBrowser(playwright, browserName).launch(new BrowserType.LaunchOptions().setHeadless(true).setChannel("chrome"));
            Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
            int width = (int) screenSize.getWidth();
            int height = (int) screenSize.getHeight();
            BrowserContext browserContext = browser.newContext(new Browser.NewContextOptions().setViewportSize(width - 30, height));
            page = browserContext.newPage();
        } else if (browserName.equalsIgnoreCase("Firefox_Headless")) {
            browser = getBrowser(playwright, browserName).launch(new BrowserType.LaunchOptions().setHeadless(true));
            Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
            int width = (int) screenSize.getWidth();
            int height = (int) screenSize.getHeight();
            BrowserContext browserContext = browser.newContext(new Browser.NewContextOptions().setViewportSize(width - 30, height));
            page = browserContext.newPage();
        } else if (browserName.equalsIgnoreCase("Edge_Headless")) {
            browser = getBrowser(playwright, browserName).launch(new BrowserType.LaunchOptions().setHeadless(true).setChannel("msedge"));
            Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
            int width = (int) screenSize.getWidth();
            int height = (int) screenSize.getHeight();
            BrowserContext browserContext = browser.newContext(new Browser.NewContextOptions().setViewportSize(width - 30, height));
            page = browserContext.newPage();
        }
        assert browser != null;
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(executionDataSheetPathAndName);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        XSSFWorkbook wb = null;
        try {
            wb = new XSSFWorkbook(fis);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        XSSFSheet xs = wb.getSheet(ExecutionScriptName);
        int lastRowNumber = xs.getLastRowNum();
        HashMap<String, String> executionPerameters = new HashMap<>();
        for (int rowCount = 1; rowCount <= lastRowNumber; rowCount++) {
            Row rowForColumnHeaders = xs.getRow(0);
            Row rowForValues = xs.getRow(rowCount);
            for (int columnCount = 0; columnCount < rowForColumnHeaders.getLastCellNum(); columnCount++) {
                Cell cellHeaderName = rowForColumnHeaders.getCell(columnCount);
                String strCellHeaderValue = cellHeaderName.toString();
                Cell cellValue = rowForValues.getCell(columnCount);
                String formattedValue;
                switch (cellValue.getCellType()) {
                    case STRING:
                        formattedValue = cellValue.getStringCellValue();
                        executionPerameters.put(strCellHeaderValue, formattedValue);
                        break;
                    case NUMERIC:
                        formattedValue = String.valueOf((long) cellValue.getNumericCellValue());
                        executionPerameters.put(strCellHeaderValue, formattedValue);
                        break;
                    case FORMULA:
                        formattedValue = evaluateFormulaCell(cellValue);
                        executionPerameters.put(strCellHeaderValue, formattedValue);
                        break;
                    default:
                        executionPerameters.put(strCellHeaderValue, "");
                }
            }
            String methodName = executionPerameters.get("Control_Action");
            String className = "Framework_Methods.Web_Control_Methods";
            Class<?> c = null;
            try {
                c = Class.forName(className);
            } catch (ClassNotFoundException e) {
                e.printStackTrace();
                throw new RuntimeException(e);
            }
            Class[] argTypes = new Class[]{int.class, Page.class, HashMap.class,String.class};
            Method main;
            try {
                main = c.getDeclaredMethod(methodName, argTypes);
            } catch (NoSuchMethodException e) {
                break;
            }
            try {
                strMethodsReturnStatus = (boolean) main.invoke(methodName, rowCount, page, executionPerameters,excelReportFileName);
                if (strMethodsReturnStatus) {
                    writeLogsInfoIntoExcel(excelReportFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Method returned status: " + methodName + "   " + strMethodsReturnStatus);
                } else {
                    writeLogsInfoIntoExcel(excelReportFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Method returned status:===================================== " + methodName + "   " + strMethodsReturnStatus);
                    break;
                }
            } catch (Exception e) {
                try {
                    writeLogsInfoIntoExcel(excelReportFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Method returned status:===================================== " + methodName + "   " + strMethodsReturnStatus);
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
            }
        }
        assert page != null;
        page.close();
        playwright.close();
        //String executionEndTime = new SimpleDateFormat("HH:mm:ss").format(new Date());
        String executionEndTime  = getCurrentTime();
        String diff=null;
        try {
            diff=executionTimeDiffCalculation(executionStartTime,executionEndTime);
        } catch (IOException | NoSuchMethodException | NoSuchFieldException | IllegalAccessException |
                 ClassNotFoundException | InstantiationException | ParseException e) {
            throw new RuntimeException(e);
        }
        String strExecutionStatus=null;
        if(strMethodsReturnStatus)
        {
            strExecutionStatus="PASSED";
        }else{
            strExecutionStatus="FAILED";
        }
        String[] arrSummaryInfo = null;
        arrSummaryInfo = new String[]{String.valueOf(threadId), ExecutionScriptName, browserName, executionStartTime, executionEndTime, diff, strExecutionStatus};
        try {
            writeSummaryInfoIntoExcel(excelReportFileName,"Summary_Info", arrSummaryInfo);
        } catch (Exception e) {
            try {
                writeLogsInfoIntoExcel(excelReportFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
            try {
                writeLogsInfoIntoExcel(excelReportFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
        }
        latch.countDown();
    }


    public static String getCurrentTime() {
        SimpleDateFormat dateFormat = dateFormatThreadLocal.get();
        return dateFormat.format(new Date());
    }

    public static String formatExecutionTime(long milliseconds) {
        long seconds = (milliseconds / 1000) % 60;
        long minutes = (milliseconds / (1000 * 60)) % 60;
        long hours = (milliseconds / (1000 * 60 * 60)) % 24;
        return String.format("%02d:%02d:%02d", hours, minutes, seconds);
    }

    //Main method
    public static void main(String[] args) throws IOException, NoSuchFieldException, ClassNotFoundException, InvocationTargetException, NoSuchMethodException, IllegalAccessException, InstantiationException {
        String[] filenames = ListFilesExample();

        System.out.println(filenames.length);
        for (int fileCount = 0; fileCount < filenames.length; fileCount++) {


            int lastIndex = filenames[fileCount].lastIndexOf('.');
            ;


            System.out.println("TestDataFileName " + filenames[fileCount].substring(0, lastIndex));
            //CreateExecutionReportFolder();
            reportFileName = filenames[fileCount].substring(0, lastIndex)+"_Execution_Report_" + customDateFormat();
            excelReportFileName = reportFileName + ".xlsx";
            htmlReportFileName = reportFileName + ".html";
            htmlReportCreate(htmlReportFileName);
            createExcelLogFile(excelReportFileName);

            executionDataSheetPathAndName = System.getProperty("user.dir") + "\\src\\test\\java\\Test_Data\\"+filenames[fileCount];
            String arr[] = readAllExecutionMethodsFromExecutionExcelFile(executionDataSheetPathAndName);
            HashMap<String, String> executionPerameters = new HashMap<>();
            if (arr == null) {
                System.out.println("Execution sheet not found in Execution Data File");
            } else if (arr.length == 0) {
                // writeLog("Info", "Headers not found in Execution Data File");
                System.out.println("Headers not found in Execution Data File");
            } else if (arr.length == 1) {
                //writeLog("Info", "Methods not found");
                System.out.println("Methods not found");
            } else {
                FileInputStream fis = new FileInputStream(executionDataSheetPathAndName);
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                XSSFSheet xs = wb.getSheet("Executions");
                for (int methodsCount = 1; methodsCount < arr.length; methodsCount++) {
                    Row row = xs.getRow(methodsCount);
                    Row rowKey = xs.getRow(0);
                    for (int cell = 0; cell < rowKey.getLastCellNum(); cell++) {
                        Cell cellValue = row.getCell(cell);
                        String v = cellValue.toString();
                        //System.out.println(v);
                        Cell keyValue = rowKey.getCell(cell);
                        String k = keyValue.toString();
                        //System.out.println(k);
                        executionPerameters.put(k, v);
                    }
                    latch = new CountDownLatch(0);
                    if (executionPerameters.get("Execution_Required").equalsIgnoreCase("Yes")) {
                        if (executionPerameters.get("Browsers_Required_For_Execution").contains(",")) {
                            String[] browserArray = executionPerameters.get("Browsers_Required_For_Execution").split(",");
                            latch = new CountDownLatch(browserArray.length);
                            ExecutionScriptName = executionPerameters.get("Scripts");
                            for (String s : browserArray) {
                                Framework_Main temp = new Framework_Main(s.trim());
                                temp.setDaemon(false);
                                temp.start();
                            }
                        } else if (executionPerameters.get("Browsers_Required_For_Execution") == null) {
                            latch = new CountDownLatch(1);
                            ExecutionScriptName = executionPerameters.get("Scripts");
                            Framework_Main temp = new Framework_Main(executionPerameters.get("Chrome").trim());
                            temp.setDaemon(false);
                            temp.start();
                        } else {
                            latch = new CountDownLatch(1);
                            ExecutionScriptName = executionPerameters.get("Scripts");
                            Framework_Main temp = new Framework_Main(executionPerameters.get("Browsers_Required_For_Execution").trim());
                            temp.setDaemon(false);
                            temp.start();
                        }
                    }
                    try {
                        assert latch != null;
                        latch.await(); // Wait for all threads in the set to finish
                        //strSensitiveUserName = null;
                        //strSensitivePassword = null;
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                }
            }
            //HTML Report builder
            writeSummaryInfoIntoHtmlReport(excelReportFileName, htmlReportFileName);
            writeTestCasesInfoIntoHtmlReport(excelReportFileName, htmlReportFileName);
            writeLogsInfoIntoHtmlReport(excelReportFileName, htmlReportFileName);
            replaceTextInHTMLReport(htmlReportFileName, "summarytablesreplacetext", "");
            replaceTextInHTMLReport(htmlReportFileName, "testcasetablereplace", "");
            replaceTextInHTMLReport(htmlReportFileName, "logtablereplace", "");
        }
    }
    public BrowserType getBrowser(Playwright playwright, String browserName) {
        BrowserType browserType;
        switch (browserName.toUpperCase().trim()) {
            case "CHROME":
            case "EDGE":
            case "CHROME_HEADLESS":
            case "EDGE_HEADLESS":
                browserType = playwright.chromium();
                break;
            case "FIREFOX":
            case "FIREFOX_HEADLESS":
                browserType = playwright.firefox();
                break;
            default:
                browserType = playwright.webkit();
        }
        return browserType;
    }

    public static void BrowserMaximized() throws AWTException, InterruptedException {
        Robot robot = new Robot();
        // Press Alt
        robot.keyPress(KeyEvent.VK_ALT);
        // Press Space
        robot.keyPress(KeyEvent.VK_SPACE);
        Thread.sleep(500);
        // Release Space
        robot.keyPress(KeyEvent.VK_X);
        // Press X
        robot.keyRelease(KeyEvent.VK_X);
        // Release X
        robot.keyRelease(KeyEvent.VK_SPACE);
        // Release Alt
        robot.keyRelease(KeyEvent.VK_ALT);
        Thread.sleep(1000);
    }


    public static void CreateExecutionReportFolder1(String browserName) throws IOException, NoSuchFieldException, ClassNotFoundException, InvocationTargetException, NoSuchMethodException, IllegalAccessException, InstantiationException {
        // Create a unique file name for the execution report based on date and time
        String FolderName = ExecutionReportFolderDateAndTime();

        String dataPath = System.getProperty("user.dir");
        // Create the path to the Excel file based on the report file name
        dataPath = dataPath + "\\src\\Execution_Reports\\" + ExecutionReportFolderName + "\\" + browserName + "_" + "_" + ExecutionScriptName + "_" + FolderName;
        //writeValueInPropertyFile("ExecutionFolderName",FolderName);
        // Create a File object representing the folder
        File folder = new File(dataPath);

        // Check if the folder already exists
        if (!folder.exists()) {
            // Attempt to create the folder
            boolean success = folder.mkdirs();

            if (success) {
                System.out.println("Folder created successfully.");
            } else {
                System.err.println("Failed to create folder.");
            }
        } else {
            System.out.println("Folder already exists.");
        }

    }

    public static void CreateExecutionReportFolder() throws IOException, NoSuchFieldException, ClassNotFoundException, InvocationTargetException, NoSuchMethodException, IllegalAccessException, InstantiationException {
        // Create a unique file name for the execution report based on date and time
        String FolderName = ExecutionReportFolderDateAndTime();
        ExecutionReportFolderName = FolderName;
        String dataPath = System.getProperty("user.dir");
        // Create the path to the Excel file based on the report file name
        dataPath = dataPath + "\\src\\Execution_Reports\\" + FolderName;
        //writeValueInPropertyFile("ExecutionFolderName",FolderName);
        // Create a File object representing the folder
        File folder = new File(dataPath);

        // Check if the folder already exists
        if (!folder.exists()) {
            // Attempt to create the folder
            boolean success = folder.mkdirs();

            if (success) {
                System.out.println("Folder created successfully.");
            } else {
                System.err.println("Failed to create folder.");
            }
        } else {
            System.out.println("Folder already exists.");
        }

    }

    public static String ExecutionReportFolderDateAndTime() throws IOException, NoSuchMethodException, NoSuchFieldException, IllegalAccessException, ClassNotFoundException, InvocationTargetException, InstantiationException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy_MMM_dd_HH_mm_ss_sss");
        return sdf.format(Calendar.getInstance().getTime());
    }

    public static String[] readAllExecutionMethodsFromExecutionExcelFile(String filePath) throws IOException {
        FileInputStream file = new FileInputStream(filePath);
        XSSFWorkbook xwb = new XSSFWorkbook(file);
        //Need to change below sheetName to executedScriptName
        String[] arr_ExecutionMethods = new String[0];
        try {
            XSSFSheet xs = xwb.getSheet("Executions");
            arr_ExecutionMethods = new String[xs.getLastRowNum() + 1];
            for (int i = 0; i <= xs.getLastRowNum(); i++) {
                Row row = xs.getRow(i);
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    arr_ExecutionMethods[i] = String.valueOf(row.getCell(j));
                }
            }
        } catch (Exception e) {
            System.out.println(e);
            //writeLog("Error", "Execution sheet not found in Execution Data File");
            System.out.println("Execution sheet not found in Execution Data File");
            //writeLog("Error", String.valueOf(e));
            System.out.println(String.valueOf(e));
            arr_ExecutionMethods = null;
        }
        return arr_ExecutionMethods;
    }


    public static String evaluateFormulaCell(Cell cell) {
        Workbook workbook = cell.getSheet().getWorkbook();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        CellValue cellValue = evaluator.evaluate(cell);
        switch (cellValue.getCellType()) {
            case STRING:
                return cellValue.getStringValue();
            case NUMERIC:
                return String.valueOf(cellValue.getNumberValue());
            case BOOLEAN:
                return String.valueOf(cellValue.getBooleanValue());
            default:
                return ""; // or handle as needed
        }
    }
}
