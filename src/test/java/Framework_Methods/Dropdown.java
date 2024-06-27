package Framework_Methods;

import com.microsoft.playwright.*;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;

import static Framework_Methods.Generic_Methods.writeLogsInfoIntoExcel;
import static Framework_Methods.Web_Control_Common_Methods.*;

public class Dropdown {
    //Select_Value_From_Dropdown
    public static boolean Select_Value_From_Dropdown(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        //String dropdownName = "";
        Locator dropdownLocator = null;
        try {
            waitForPageLoad(page);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                dropdownLocator = page.locator(String.format(map.get("ATTRIBUTE_VALUE").toString(), map.get("CONTROL_VALUE").toString()));
            } else {
                dropdownLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            }
            methodStatus = waitForPageLocator(page, dropdownLocator);
            if (methodStatus) {
                if (dropdownLocator.isVisible()) {
                    //dropdownName = dropdownLocator.textContent();
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Dropdown is Visible");
                    if (dropdownLocator.isEnabled()) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Dropdown is Enabled");
                        dropdownLocator.selectOption(map.get("CONTROL_VALUE").toString());
                    } else {
                        methodStatus = false;
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Dropdown is not Enabled");
                    }
                } else {
                    methodStatus = false;
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Dropdown is not Visible");
                }
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Value from Dropdown has been selected successfully");
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to select the value from dropdown");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert dropdownLocator != null;
            int elementCount = dropdownLocator.count();
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

    //Select_Value_From_Web_Dropdown
    public static boolean Select_Value_From_Web_Dropdown(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            String[] dropdownLocators = map.get("ATTRIBUTE_VALUE").toString().split("<<=>>");
            String dropdownLocator = dropdownLocators[0].trim();
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementDropdown = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            String dropdownSelectValue = dropdownLocators[1].trim();
            String elementPath;
            if (dropdownSelectValue.contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                elementPath = (dropdownSelectValue).replace("%s", dynamicValue);
            } else {
                elementPath = dropdownSelectValue;
            }
            boolean isDisplayed = elementDropdown.isVisible();
            boolean isEnabled = elementDropdown.isEnabled();
            if (isDisplayed && isEnabled) {
                elementDropdown.click();
                page.waitForTimeout(1000);
                ElementHandle option = page.querySelector(elementPath);
                Thread.sleep(1000);
                option.click();
                methodStatus = true;
            } else {
                System.out.println("Control is not displayed on the page");
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
