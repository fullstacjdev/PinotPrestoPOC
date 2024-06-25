package Framework_Methods;

import com.microsoft.playwright.*;
import com.microsoft.playwright.assertions.PlaywrightAssertions;

import javax.swing.*;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Scanner;

import static Framework_Methods.Generic_Methods.writeLogsInfoIntoExcel;
import static Framework_Methods.Web_Control_Common_Methods.*;
import static com.microsoft.playwright.assertions.PlaywrightAssertions.assertThat;

public class Web_Page {

    public static void verifyPageTitle(Page page, String expectedTitle, int timeoutInSeconds) {
        long startTime = System.currentTimeMillis();
        boolean isTitleMatched = false;

        while (!isTitleMatched && (System.currentTimeMillis() - startTime) < timeoutInSeconds * 1000) {
            String actualTitle = page.title();
            if (actualTitle.contains(expectedTitle)) {
                System.out.println("Page title matched: " + actualTitle);
                isTitleMatched = true;
            } else {
                try {
                    Thread.sleep(200);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
        }
        if (!isTitleMatched) {
            System.out.println("Page title didn't match within the timeout.");
        }
    }


    //Page_Load
    public static boolean Page_Load(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            page.navigate(map.get("CONTROL_VALUE").toString());
            waitForPageLoad(page);
            page.waitForLoadState();
            methodStatus = true;
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", page.title()+" has been launched successfully");
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }


    //Get_Page_Title_Equals
    public static boolean Get_Page_Title_Equals(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            String expectedTitle = map.get("CONTROL_VALUE").toString().trim();
            verifyPageTitle(page, map.get("CONTROL_VALUE").toString().trim(), 30);
            assertThat(page).hasTitle(map.get("CONTROL_VALUE").toString().trim());
            String actualPageTitle = page.title().trim();
            if (actualPageTitle.equals(expectedTitle)) {
                methodStatus = true;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Get_Page_Title_Contains
    public static boolean Get_Page_Title_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            waitForElementsEnabled(page);
            String expectedTitle = map.get("CONTROL_VALUE").toString().trim();
            verifyPageTitle(page, map.get("CONTROL_VALUE").toString().trim(), 30);
            String actualPageTitle = page.title().trim();
            if (actualPageTitle.contains(expectedTitle)) {
                methodStatus = true;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Get_Page_Title_Equals
    public static boolean Get_Page_Title_Equals_Ignore_Case(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            String expectedTitle = map.get("CONTROL_VALUE").toString().trim();
            verifyPageTitle(page, expectedTitle, 30);
            String actualPageTitle = page.title().trim();
            if (actualPageTitle.equalsIgnoreCase(expectedTitle)) {
                methodStatus = true;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Get_Page_Title_Contains
    public static boolean Get_Page_Title_Contains_Ignore_Case(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            waitForElementsEnabled(page);
            String expectedTitle = map.get("CONTROL_VALUE").toString().toUpperCase().trim();
            verifyPageTitle(page, expectedTitle, 30);
            String actualPageTitle = page.title().trim().toUpperCase();
            if (actualPageTitle.contains(expectedTitle)) {
                methodStatus = true;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

}
