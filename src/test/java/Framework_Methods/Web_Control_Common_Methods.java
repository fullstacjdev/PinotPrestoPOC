package Framework_Methods;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.LoadState;
import org.openqa.selenium.JavascriptExecutor;

import javax.swing.*;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Scanner;

import static Framework_Methods.Generic_Methods.writeLogsInfoIntoExcel;
import static Framework_Methods.Generic_Methods.writeTestCasesInfoIntoExcel;

public class Web_Control_Common_Methods {
    public static void waitForElementsEnabled(Page page) {
        String script = "(() => { " +
                "const allElements = document.querySelectorAll('*');" +
                "for (let element of allElements) {" +
                "if (!element.disabled && element.offsetWidth > 0 && element.offsetHeight > 0) {" +
                "return true;" +
                "}" +
                "}" +
                "return false;" +
                "})()";
        try {
            page.waitForFunction(script, new Page.WaitForFunctionOptions().setTimeout(10000));
        } catch (PlaywrightException e) {

        }
    }



    public static boolean waitForPageContentEnabled(Page page) {
        String script = "document.readyState === 'complete' && !document.querySelector('body').hasAttribute('disabled');";
        try {
            page.waitForFunction(script, new Page.WaitForFunctionOptions().setTimeout(12000));
            return true;
        } catch (PlaywrightException e) {
            System.out.println("Error waiting for page content to be enabled: " + e.getMessage());
            return false;
        }
    }

    private static boolean waitForLoadStateWithRetry(Page page) {
        int retryCount = 3;
        int currentAttempt = 0;
        while (currentAttempt < retryCount) {
            try {
                page.waitForLoadState(LoadState.DOMCONTENTLOADED);
                page.waitForLoadState(LoadState.DOMCONTENTLOADED);
                page.waitForLoadState(LoadState.NETWORKIDLE);
                return true;
            } catch (Exception e) {
                System.err.println("Attempt " + (currentAttempt + 1) + " failed: " + e.getMessage());
            }
            currentAttempt++;
        }
        return false;
    }

    public static void waitForPageLoad(Page page) {
        Page.WaitForLoadStateOptions options = new Page.WaitForLoadStateOptions().setTimeout(120000);
        page.waitForLoadState(LoadState.LOAD, options);
        page.waitForLoadState(LoadState.NETWORKIDLE, options);
        page.waitForLoadState(LoadState.DOMCONTENTLOADED, options);
    }

    public static ElementHandle waitForAnElement(Page page, String elementString, int timeoutMillis) {
        long startTime = System.currentTimeMillis();
        ElementHandle element = null;
        while (System.currentTimeMillis() - startTime < timeoutMillis) {
            try {
                element = page.querySelector(elementString);
                if (element != null && element.isVisible()) {
                    return element;
                }
            } catch (PlaywrightException e) {
            }
        }
        return null;
    }


    public static String ScreenshotToBase64(Page page, String selector, String color) throws IOException {
        SimpleDateFormat uniqueImageName = new SimpleDateFormat("yyyy_MMM_dd_HH_mm_ss_sss");
        String imageFilePathAndName = System.getProperty("user.dir") + "\\src\\test\\java\\Images\\Image_" + uniqueImageName.format(Calendar.getInstance().getTime()) + ".PNG";
        waitForPageLoad(page);
        ElementHandle element = page.waitForSelector(selector);
        page.evaluate("element => { element.style.outline = '4px solid " + color + "'; }", element);
        try {
            Thread.sleep(1000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        byte[] screenshotBytes = page.screenshot(new Page.ScreenshotOptions());
        try {
            Thread.sleep(1000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        Path screenshotPath = Paths.get(imageFilePathAndName);
        Files.write(screenshotPath, screenshotBytes);
        page.evaluate("element => { element.style.outline = ''; }", element);
        return imageFilePathAndName;
    }

    public static String ScreenshotToBase64(Page page) throws IOException {
        SimpleDateFormat uniqueImageName = new SimpleDateFormat("yyyy_MMM_dd_HH_mm_ss_sss");
        String ImageFilePathAndName = System.getProperty("user.dir") + "\\src\\test\\java\\Images\\Image_" + uniqueImageName.format(Calendar.getInstance().getTime()) + ".PNG";
        waitForPageLoad(page);
        try {
            Thread.sleep(1000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        byte[] screenshotBytes = page.screenshot(new Page.ScreenshotOptions());
        Path screenshotPath = Paths.get(ImageFilePathAndName);
        Files.write(screenshotPath, screenshotBytes);
        return ImageFilePathAndName;
    }

    public static void writeLogsAndTCsInfoIntoReportFile(Page page, String reportLogFileName, HashMap map, boolean methodStatus) throws IOException {
        if (methodStatus) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Page has been loaded successfully");
            if (((!map.get("TEST_CASES").toString().isEmpty()) && (map.get("ACTION_FLOW").toString().equals("Assertion"))) && ((map.get("SHOULD_TAKE_SCREENSHOT").toString().equals("") || map.get("SHOULD_TAKE_SCREENSHOT").toString().equalsIgnoreCase("YES")))) {
                writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("TEST_CASES").toString(), ScreenshotToBase64(page), "Pass");
            } else if (((!map.get("TEST_CASES").toString().isEmpty()) && (map.get("ACTION_FLOW").toString().equals("Assertion"))) && ((map.get("SHOULD_TAKE_SCREENSHOT").toString().equals("") || map.get("SHOULD_TAKE_SCREENSHOT").toString().equalsIgnoreCase("NO")))) {
                writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("TEST_CASES").toString(), "", "Pass");
            }
        } else {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Page is not loaded");
            if (((!map.get("TEST_CASES").toString().isEmpty()) && (map.get("ACTION_FLOW").toString().equals("Assertion"))) && ((map.get("SHOULD_TAKE_SCREENSHOT").toString().equals("") || map.get("SHOULD_TAKE_SCREENSHOT").toString().equalsIgnoreCase("YES")))) {
                writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("TEST_CASES").toString(), ScreenshotToBase64(page), "Fail");
            } else if (((!map.get("TEST_CASES").toString().isEmpty()) && (map.get("ACTION_FLOW").toString().equals("Assertion"))) && ((map.get("SHOULD_TAKE_SCREENSHOT").toString().equals("") || map.get("SHOULD_TAKE_SCREENSHOT").toString().equalsIgnoreCase("NO")))) {
                writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("TEST_CASES").toString(), "", "Fail");
            }
        }
    }

    public static void writeLogsAndTCsInfoIntoReportFileWithElementMarking(Page page, String reportLogFileName, HashMap map, boolean methodStatus) throws IOException {
        if (methodStatus) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Page has been loaded successfully");
            if (((!map.get("TEST_CASES").toString().isEmpty()) && (map.get("ACTION_FLOW").toString().equals("Assertion"))) && ((map.get("SHOULD_TAKE_SCREENSHOT").toString().equals("") || map.get("SHOULD_TAKE_SCREENSHOT").toString().equalsIgnoreCase("YES")))) {
                writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("TEST_CASES").toString(), ScreenshotToBase64(page, map.get("ATTRIBUTE_VALUE").toString(), "green"), "Pass");
            } else if (((!map.get("TEST_CASES").toString().isEmpty()) && (map.get("ACTION_FLOW").toString().equals("Assertion"))) && ((map.get("SHOULD_TAKE_SCREENSHOT").toString().equals("") || map.get("SHOULD_TAKE_SCREENSHOT").toString().equalsIgnoreCase("NO")))) {
                writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("TEST_CASES").toString(), "", "Pass");
            }
        } else {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Page is not loaded");
            if (((!map.get("TEST_CASES").toString().isEmpty()) && (map.get("ACTION_FLOW").toString().equals("Assertion"))) && ((map.get("SHOULD_TAKE_SCREENSHOT").toString().equals("") || map.get("SHOULD_TAKE_SCREENSHOT").toString().equalsIgnoreCase("YES")))) {
                writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("TEST_CASES").toString(), ScreenshotToBase64(page, map.get("ATTRIBUTE_VALUE").toString(), "red"), "Fail");
            } else if (((!map.get("TEST_CASES").toString().isEmpty()) && (map.get("ACTION_FLOW").toString().equals("Assertion"))) && ((map.get("SHOULD_TAKE_SCREENSHOT").toString().equals("") || map.get("SHOULD_TAKE_SCREENSHOT").toString().equalsIgnoreCase("NO")))) {
                writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("TEST_CASES").toString(), "", "Fail");
            }
        }
    }


    // Function to check if the element is visible
    private static void isElementVisible(Page page, ElementHandle element) {
        JavascriptExecutor js = (JavascriptExecutor) page;
        boolean elementFound = false;
        int scrollStartPosition = 0;
        int scrollEndPosition = 0;
        int maxScroll = 5000;
        while (!elementFound && scrollStartPosition <= maxScroll) {
            js.executeScript("window.scrollTo(scrollStartPosition, " + scrollEndPosition + ")");
            if (element.isVisible()) {
                elementFound = true;
            }
            scrollStartPosition = scrollStartPosition + scrollEndPosition;
            scrollEndPosition += 500;
        }
        if (elementFound) {
            System.out.println("Element found!");
        } else {
            System.out.println("Element not found after scrolling.");
        }
    }

    //Page_Down
    public static boolean Page_Down(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            Keyboard keyboard = page.keyboard();
            while (true) {
                keyboard.press("PageDown");
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                if (newScrollY.equals(initialScrollY)) {
                    return methodStatus = true;
                }
                initialScrollY = newScrollY;
            }
        } catch (Exception e) {
            System.out.println();
        }
        return methodStatus;
    }


    //Wait_For_Short_Time
    public static boolean Wait_For_Short_Time(int rownumber, Page page, HashMap map, String reportLogFileName) {
        boolean methodStatus = false;
        try {
            Thread.sleep(5000);
            methodStatus = true;
        } catch (Exception e) {
            System.out.println(e.toString());
        }
        return methodStatus;
    }

    //Wait_For_Medium_Time
    public static boolean Wait_For_Medium_Time(int rownumber, Page page, HashMap map, String reportLogFileName) {
        boolean methodStatus = false;
        try {
            Thread.sleep(7500);
            methodStatus = true;
        } catch (Exception e) {
            System.out.println(e.toString());
        }
        return methodStatus;
    }

    //Wait_For_Long_Time
    public static boolean Wait_For_Long_Time(int rownumber, Page page, HashMap map, String reportLogFileName) {
        boolean methodStatus = false;
        try {
            Thread.sleep(10000);
            methodStatus = true;
        } catch (Exception e) {
            System.out.println(e.toString());
        }
        return methodStatus;
    }


}
