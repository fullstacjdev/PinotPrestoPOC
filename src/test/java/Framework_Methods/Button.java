package Framework_Methods;
import com.microsoft.playwright.*;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import static Framework_Methods.Generic_Methods.writeLogsInfoIntoExcel;
import static Framework_Methods.Web_Control_Common_Methods.*;
public class Button {
    //Button_Click
    public static boolean Button_Click(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        String buttonName = "";
        Locator buttonLocator = null;
        try {
            waitForPageLoad(page);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                buttonLocator = page.locator(String.format(map.get("ATTRIBUTE_VALUE").toString(), map.get("CONTROL_VALUE").toString()));
            } else {
                buttonLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            }
            methodStatus = waitForPageLocator(page, buttonLocator);
            if (methodStatus) {
                if (buttonLocator.isVisible()) {
                    buttonName = buttonLocator.textContent();
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", buttonName + " " + "Button is Visible");
                    if (buttonLocator.isEnabled()) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", buttonName + " " + "Button is Enabled");
                        buttonLocator.click();
                    } else {
                        methodStatus = false;
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", buttonName + " " + "Button is not Enabled");
                    }
                } else {
                    methodStatus = false;
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", buttonName + " " + "Button is not Visible");
                }
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", buttonName + " " + "Button has been clicked successfully");
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", buttonName + " " + "Button click was unsuccessful");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert buttonLocator != null;
            int elementCount = buttonLocator.count();
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
    //Verify_Button_Name_Equals
    public static boolean Verify_Button_Name_Equals(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        String buttonName = "";
        String expectedButtonName = map.get("CONTROL_VALUE").toString();
        Locator buttonLocator = null;
        try {
            waitForPageLoad(page);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                buttonLocator = page.locator(String.format(map.get("ATTRIBUTE_VALUE").toString(), map.get("CONTROL_VALUE").toString()));
            } else {
                buttonLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            }
            methodStatus = waitForPageLocator(page, buttonLocator);
            if (methodStatus) {
                if (buttonLocator.isVisible()) {
                    buttonName = buttonLocator.textContent();
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", buttonName + " " + "Button is Visible");
                    if (buttonName.equals(expectedButtonName)) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Actual and Expected Button Names are matched");
                    } else {
                        methodStatus = false;
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Actual and Expected Button Names are not matched");
                    }
                } else {
                    methodStatus = false;
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", buttonName + " " + "Button is not Visible");
                }
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", buttonName + " " + "Button name verification is successful");
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", buttonName + " " + "Button name verification was unsuccessful");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert buttonLocator != null;
            int elementCount = buttonLocator.count();
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

    //Verify_Button_Name_Equals_Ignore_Case
    public static boolean Verify_Button_Name_Equals_Ignore_Case(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        String buttonName = "";
        String expectedButtonName = map.get("CONTROL_VALUE").toString();
        Locator buttonLocator = null;
        try {
            waitForPageLoad(page);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                buttonLocator = page.locator(String.format(map.get("ATTRIBUTE_VALUE").toString(), map.get("CONTROL_VALUE").toString()));
            } else {
                buttonLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            }
            methodStatus = waitForPageLocator(page, buttonLocator);
            if (methodStatus) {
                if (buttonLocator.isVisible()) {
                    buttonName = buttonLocator.textContent();
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", buttonName + " " + "Button is Visible");
                    if (buttonName.equalsIgnoreCase(expectedButtonName)) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Actual and Expected Button Names are matched");
                    } else {
                        methodStatus = false;
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Actual and Expected Button Names are not matched");
                    }
                } else {
                    methodStatus = false;
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", buttonName + " " + "Button is not Visible");
                }
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", buttonName + " " + "Button name verification is successful");
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", buttonName + " " + "Button name verification was unsuccessful");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert buttonLocator != null;
            int elementCount = buttonLocator.count();
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

    //Verify_Button_Name_Contains
    public static boolean Verify_Button_Name_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        String buttonName = "";
        String expectedButtonName = map.get("CONTROL_VALUE").toString();
        Locator buttonLocator = null;
        try {
            waitForPageLoad(page);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                buttonLocator = page.locator(String.format(map.get("ATTRIBUTE_VALUE").toString(), map.get("CONTROL_VALUE").toString()));
            } else {
                buttonLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            }
            methodStatus = waitForPageLocator(page, buttonLocator);
            if (methodStatus) {
                if (buttonLocator.isVisible()) {
                    buttonName = buttonLocator.textContent();
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", buttonName + " " + "Button is Visible");
                    if (buttonName.contains(expectedButtonName)) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Actual and Expected Button Names are matched");
                    } else {
                        methodStatus = false;
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Actual and Expected Button Names are not matched");
                    }
                } else {
                    methodStatus = false;
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", buttonName + " " + "Button is not Visible");
                }
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", buttonName + " " + "Button name verification is successful");
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", buttonName + " " + "Button name verification was unsuccessful");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert buttonLocator != null;
            int elementCount = buttonLocator.count();
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


    //Need to update below methods


    //Button_Enable_Status
    public static boolean Button_Enable_Status(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementButton = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 12000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementButton = page.waitForSelector(elementPath);
            } else {
                elementButton = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString());
            }
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            Keyboard keyboard = page.keyboard();
            while (true) {
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                if (newScrollY.equals(initialScrollY)) {
                    if (elementButton.isEnabled()) {
                        methodStatus = true;
                        break;
                    }
                }
                keyboard.press("PageDown");
                initialScrollY = newScrollY;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Button_Visible_Status
    public static boolean Button_Visible_Status(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            //methodStatus = waitForPageLocator(page, map.get("ATTRIBUTE_VALUE").toString());

            ElementHandle elementButton = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 12000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementButton = page.waitForSelector(elementPath);
            } else {
                elementButton = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString());
            }
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            Keyboard keyboard = page.keyboard();
            while (true) {
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                if (newScrollY.equals(initialScrollY)) {
                    if (elementButton.isVisible()) {
                        methodStatus = true;
                        break;
                    }
                }
                keyboard.press("PageDown");
                initialScrollY = newScrollY;
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
