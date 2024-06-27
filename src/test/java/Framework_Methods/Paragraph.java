package Framework_Methods;

import com.microsoft.playwright.Locator;
import com.microsoft.playwright.Page;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;

import static Framework_Methods.Generic_Methods.writeLogsInfoIntoExcel;
import static Framework_Methods.Web_Control_Common_Methods.*;

public class Paragraph {
    //Verify_Paragraph_Text_Equals
    public static boolean Verify_Paragraph_Text_Equals(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        String paragraphName = "";
        String expectedHeaderName = map.get("CONTROL_VALUE").toString();
        Locator paragraphLocator = null;
        try {
            waitForPageLoad(page);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                paragraphLocator = page.locator(String.format(map.get("ATTRIBUTE_VALUE").toString(), map.get("CONTROL_VALUE").toString()));
            } else {
                paragraphLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            }
            methodStatus = waitForPageLocator(page, paragraphLocator);
            if (methodStatus) {
                if (paragraphLocator.isVisible()) {
                    paragraphName = paragraphLocator.textContent();
                    System.out.println("Inovar =======>"+paragraphName);
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", paragraphName + " " + "is Visible");
                    if (paragraphName.equals(expectedHeaderName)) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Actual and Expected Paragraph Names are matched");
                    } else {
                        methodStatus = false;
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Actual and Expected Paragraph Names are not matched");
                    }
                } else {
                    methodStatus = false;
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", paragraphName + " " + "is not Visible");
                }
                //writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", paragraphName + " " + "Button name verification is successful");
            } else {
                //writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", headerName + " " + "Button name verification was unsuccessful");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert paragraphLocator != null;
            int elementCount = paragraphLocator.count();
            if (elementCount > 0) {
                writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
            } else {
                writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
            }
        }
        if (map.get("VALIDATION_TYPE").equals("Partial_Assertion")) {
            methodStatus = true;
        }
        return methodStatus;
    }

    //Verify_Paragraph_Text_Equals_Ignore_Case
    public static boolean Verify_Paragraph_Text_Equals_Ignore_Case(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
            boolean methodStatus = false;
            String paragraphName = "";
            String expectedHeaderName = map.get("CONTROL_VALUE").toString();
            Locator paragraphLocator = null;
            try {
                waitForPageLoad(page);
                if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                    paragraphLocator = page.locator(String.format(map.get("ATTRIBUTE_VALUE").toString(), map.get("CONTROL_VALUE").toString()));
                } else {
                    paragraphLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
                }
                methodStatus = waitForPageLocator(page, paragraphLocator);
                if (methodStatus) {
                    if (paragraphLocator.isVisible()) {
                        paragraphName = paragraphLocator.textContent();
                        System.out.println("Inovar =======>"+paragraphName);
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", paragraphName + " " + "is Visible");
                        if (paragraphName.equalsIgnoreCase(expectedHeaderName)) {
                            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Actual and Expected Paragraph Names are matched");
                        } else {
                            methodStatus = false;
                            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Actual and Expected Paragraph Names are not matched");
                        }
                    } else {
                        methodStatus = false;
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", paragraphName + " " + "is not Visible");
                    }
                    //writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", paragraphName + " " + "Button name verification is successful");
                } else {
                    //writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", headerName + " " + "Button name verification was unsuccessful");
                }
            } catch (Exception e) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
            } finally {
                assert paragraphLocator != null;
                int elementCount = paragraphLocator.count();
                if (elementCount > 0) {
                    writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
                } else {
                    writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
                }
            }
            if (map.get("VALIDATION_TYPE").equals("Partial_Assertion")) {
                methodStatus = true;
            }
            return methodStatus;
        }

    //Verify_Paragraph_Text_Contains
    public static boolean Verify_Paragraph_Text_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        String paragraphName = "";
        String expectedHeaderName = map.get("CONTROL_VALUE").toString();
        Locator paragraphLocator = null;
        try {
            waitForPageLoad(page);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                paragraphLocator = page.locator(String.format(map.get("ATTRIBUTE_VALUE").toString(), map.get("CONTROL_VALUE").toString()));
            } else {
                paragraphLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            }
            methodStatus = waitForPageLocator(page, paragraphLocator);
            if (methodStatus) {
                if (paragraphLocator.isVisible()) {
                    paragraphName = paragraphLocator.textContent();
                    System.out.println("Inovar =======>"+paragraphName);
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", paragraphName + " " + "is Visible");
                    if (paragraphName.equalsIgnoreCase(expectedHeaderName)) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Actual and Expected Paragraph Names are matched");
                    } else {
                        methodStatus = false;
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Actual and Expected Paragraph Names are not matched");
                    }
                } else {
                    methodStatus = false;
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", paragraphName + " " + "is not Visible");
                }
                //writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", paragraphName + " " + "Button name verification is successful");
            } else {
                //writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", headerName + " " + "Button name verification was unsuccessful");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert paragraphLocator != null;
            int elementCount = paragraphLocator.count();
            if (elementCount > 0) {
                writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
            } else {
                writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
            }
        }
        if (map.get("VALIDATION_TYPE").equals("Partial_Assertion")) {
            methodStatus = true;
        }
        return methodStatus;
    }
}
