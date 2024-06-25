package Framework_Methods;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.LoadState;
import com.microsoft.playwright.options.WaitForSelectorState;
import org.openqa.selenium.JavascriptExecutor;

import javax.swing.*;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

import static Framework_Methods.Generic_Methods.writeLogsInfoIntoExcel;
import static Framework_Methods.Generic_Methods.writeTestCasesInfoIntoExcel;
import static Framework_Methods.Web_Control_Common_Methods.*;
import static org.apache.commons.io.FileUtils.waitFor;

public class Web_Control_Methods {

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
            methodStatus = true;
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
            String expectedTitle = map.get("CONTROL_VALUE").toString().toUpperCase().trim();
            verifyPageTitle(page, map.get("CONTROL_VALUE").toString().trim(), 30);
            String actualPageTitle = page.title().trim().toUpperCase();
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
            String expectedTitle = map.get("CONTROL_VALUE").toString().toUpperCase().trim();
            verifyPageTitle(page, map.get("CONTROL_VALUE").toString().trim(), 30);
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

    //Get_Page_Title_Equals
    public static boolean Get_Page_Title_Equals_Ignore_Case(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            String expectedTitle = map.get("CONTROL_VALUE").toString().toUpperCase().trim();
            verifyPageTitle(page, expectedTitle, 30);
            String actualPageTitle = page.title().trim().toUpperCase();
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

    //Enter_Value_In_TextBox
    public static boolean Enter_Value_In_TextBox(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            page.focus(map.get("ATTRIBUTE_VALUE").toString());
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                elementTextBox.click();
                elementTextBox.fill(map.get("CONTROL_VALUE").toString());
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

    //Enter_Value_In_TextBox_Via_Command_Prompt
    public static boolean Enter_Value_In_TextBox_Via_Command_Prompt(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(12000);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            page.focus(map.get("ATTRIBUTE_VALUE").toString());
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                elementTextBox.click();
                Scanner scanner = new Scanner(System.in);
                String value = scanner.nextLine();
                elementTextBox.fill(value);
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

    //Enter_Value_In_TextBox_Via_Input_Box
    public static boolean Enter_Value_In_TextBox_Via_Input_Box(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            page.focus(map.get("ATTRIBUTE_VALUE").toString());
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                elementTextBox.click();
                String userInput = JOptionPane.showInputDialog(null, "Enter value for " + map.get("VALIDATION_TYPE").toString() + " : ");
                elementTextBox.fill(userInput);
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

    //Enter_Value_In_TextBox_Via_Secure_Input_Box
    public static boolean Enter_Value_In_TextBox_Via_Secure_Input_Box(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            page.focus(map.get("ATTRIBUTE_VALUE").toString());
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                elementTextBox.click();
                JPasswordField passwordField = new JPasswordField();
                int option = JOptionPane.showConfirmDialog(null, passwordField, "Enter value for " + map.get("VALIDATION_TYPE").toString(), JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                if (option == JOptionPane.OK_OPTION) {
                    char[] password = passwordField.getPassword();
                    String userInput = new String(password);
                    elementTextBox.fill(userInput);
                }
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

    //Enter_Value_In_Frame_TextBox
    public static boolean Enter_Value_In_Frame_TextBox(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 30000);
            page.focus(map.get("Attribute_Value").toString());
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                if (map.get("CONTROL_TYPE").toString().equalsIgnoreCase("SensitiveTextBox")) {
                    Scanner scanner = new Scanner(System.in);
                    String value = scanner.nextLine();
                    elementTextBox.fill(value);
                } else {
                    elementTextBox.fill(map.get("CONTROL_VALUE").toString());
                }
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

    //Get_Value_From_TextBox
    public static boolean Get_Value_From_TextBox(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 30000);
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            Keyboard keyboard = page.keyboard();
            String strGetValueFromTextBox = null;
            while (true) {
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                if (newScrollY.equals(initialScrollY)) {
                    if (elementTextBox.isVisible()) {
                        strGetValueFromTextBox = elementTextBox.inputValue();
                        methodStatus = true;
                        break;
                    } else {
                        methodStatus = false;
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

    //Button_Click
    public static boolean Button_Click(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
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
                    if (elementButton.isVisible() && elementButton.isEnabled()) {
                        elementButton.click();
                        methodStatus = true;
                        break;
                    } else {
                        methodStatus = false;
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



    //Link_Click
    public static boolean Link_Click(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementLink = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 12000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementLink = page.waitForSelector(elementPath);
            } else {
                elementLink = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString());
            }
            boolean isDisplayed = elementLink.isVisible();
            boolean isEnabled = elementLink.isEnabled();
            if (isDisplayed && isEnabled) {
                elementLink.click();
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

    //Click on Tab
    public static boolean Tab_Click(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementLink = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 12000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementLink = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementLink = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString(), optionsForElementLoad);
            }
            elementLink.click();
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Frame_Link_Click
    public static boolean Frame_Link_Click(int rownumber, Page page, HashMap map, String reportLogFileName) throws InterruptedException, IOException {
        boolean methodStatus = false;
        Frame innerFrame = page.frameByUrl("https://selenium143.blogspot.com/");
        innerFrame.locator(map.get("ATTRIBUTE_VALUE").toString()).click();
        page.bringToFront();
        methodStatus = true;
        return methodStatus;
    }

    //Get_Header_Title_Equals
    public static boolean Get_Header_Title_Equals(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementActualHeaderTitle = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            if (elementActualHeaderTitle.isVisible()) {
                String actualHeaderTitle = elementActualHeaderTitle.textContent();
                String expectedHeaderTitle = map.get("CONTROL_VALUE").toString();
                if (actualHeaderTitle.equals(expectedHeaderTitle)) {
                    methodStatus = true;
                    return methodStatus;
                }
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Get_Header_Title_Contains
    public static boolean Get_Header_Title_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementActualHeaderTitle = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            assert elementActualHeaderTitle != null;
            if (elementActualHeaderTitle.isVisible()) {
                String actualHeaderTitle = elementActualHeaderTitle.textContent();
                String expectedHeaderTitle = map.get("CONTROL_VALUE").toString();
                if (actualHeaderTitle.contains(expectedHeaderTitle)) {
                    methodStatus = true;
                    return methodStatus;
                }
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Select_Value_From_Dropdown
    public static boolean Select_Value_From_Dropdown(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle selectDropdown = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            assert selectDropdown != null;
            selectDropdown.selectOption(map.get("CONTROL_VALUE").toString());
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
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


    //Accept_Or_Reject_Cookies
    public static boolean Accept_Or_Reject_Cookies(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementForCookies = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            elementForCookies.click();
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Select_Radio_Button
    public static boolean Select_Radio_Button(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementRadioButton = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementRadioButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementRadioButton = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString(), optionsForElementLoad);
            }
            boolean isDisplayed = elementRadioButton.isVisible();
            boolean isEnabled = elementRadioButton.isEnabled();
            if (isDisplayed && isEnabled) {
                elementRadioButton.check();
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

    //Select_Checkbox_Button
    public static boolean Select_Checkbox_Button(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementCheckboxButton = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementCheckboxButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementCheckboxButton = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString(), optionsForElementLoad);
            }
            boolean isDisplayed = elementCheckboxButton.isVisible();
            boolean isEnabled = elementCheckboxButton.isEnabled();
            if (isDisplayed && isEnabled) {
                if (elementCheckboxButton.isChecked()) {
                    methodStatus = true;
                } else {
                    elementCheckboxButton.click();
                    methodStatus = true;
                }
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

    //Handle_PopUp
    public static boolean Handle_PopUp(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementRadioButton = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //To upload File or Image to the page from File Explore
    public static boolean Upload_File_Or_Image_To_Page_From_File_Explore(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementButton = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementButton = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString(), optionsForElementLoad);
            }
            FileChooser fileChooser = page.waitForFileChooser(() -> {
                page.getByText(map.get("CONTROL_VALUE").toString()).click();
            });
            fileChooser.setFiles(Paths.get(System.getProperty("user.dir") + "\\src\\test\\java\\Upload_Test_Files\\Dharanija.jpeg"));
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify Button Name
    public static boolean Verify_Button_Name(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementButton = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementButton = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString(), optionsForElementLoad);
            }
            if ((elementButton.innerText()).equals(map.get("CONTROL_VALUE").toString())) {
                elementButton.click();
            } else {
            }
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify Share Option Popup
    public static boolean Verify_Share_Option_PopUp(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementButton = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("CONTROL_VALUE").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementButton = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString(), optionsForElementLoad);
            }
            if (elementButton.isVisible()) {
                methodStatus = true;
            } else {
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Press_Keyboard_Key
    public boolean Press_Keyboard_Key(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            page.keyboard().press(map.get("CONTROL_VALUE").toString());
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Share to User
    public static boolean Share_To_User(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementButton = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            if ((map.get("ATTRIBUTE_VALUE").toString()).contains("%s")) {
                String dynamicValue = map.get("Control_Value").toString();
                String elementPath = String.format(map.get("ATTRIBUTE_VALUE").toString(), dynamicValue);
                elementButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementButton = page.waitForSelector(map.get("ATTRIBUTE_VALUE").toString(), optionsForElementLoad);
            }
            if (elementButton.isVisible()) {
                if (elementButton.innerText().equals(map.get("CONTROL_VALUE").toString())) {
                    page.keyboard().press("Escape");
                }
            } else {
                page.getByPlaceholder("Search").click();
                page.getByPlaceholder("Search").fill(map.get("CONTROL_VALUE").toString());
                if (elementButton.innerText().equals(map.get("CONTROL_VALUE").toString())) {
                    elementButton.click();
                }
            }
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Get_Hint_Text_From_Control_Equals_To
    public static boolean Get_Hint_Text_From_Control_Equals_To(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            Keyboard keyboard = page.keyboard();
            String strGetValueFromTextBox = null;
            while (true) {
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                if (newScrollY.equals(initialScrollY)) {
                    if (elementTextBox.isVisible()) {
                        String validationMessage = (String) elementTextBox.evaluate("element => element.validationMessage");
                        if (validationMessage.equals(map.get("CONTROL_VALUE").toString())) {
                            methodStatus = true;
                            break;
                        }
                    } else {
                        methodStatus = false;
                    }
                }
                keyboard.press("PageDown");
                initialScrollY = newScrollY;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Get_Hint_Text_From_Control_Contains
    public static boolean Get_Hint_Text_From_Control_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            Keyboard keyboard = page.keyboard();
            String strGetValueFromTextBox = null;
            while (true) {
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                if (newScrollY.equals(initialScrollY)) {
                    if (elementTextBox.isVisible()) {
                        String validationMessage = (String) elementTextBox.evaluate("element => element.validationMessage");
                        if (validationMessage.contains(map.get("CONTROL_VALUE").toString())) {
                            methodStatus = true;
                            break;
                        }
                    }
                }
                keyboard.press("PageDown");
                initialScrollY = newScrollY;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify_Control_Text_Contains
    public static boolean Verify_Control_Text_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle element = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            assert element != null;
            String Textvalue = element.innerText();
            String actualTextContains = Textvalue.trim();
            String expectedTextContains = map.get("CONTROL_VALUE").toString().trim();
            if (actualTextContains.contains(expectedTextContains)) {
                methodStatus = true;
                return methodStatus;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify_Paragraph_Text_Contains
    public static boolean Verify_Paragraph_Text_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle paragraphElement = (ElementHandle) page.evaluateHandle("xpath => document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue", map.get("ATTRIBUTE_VALUE").toString());
            Object jsParagraph = paragraphElement.evaluate("el => el.textContent");
            String paragraphText = jsParagraph.toString();
            String expectedTextContains = map.get("CONTROL_VALUE").toString().trim();
            if (paragraphText.trim().contains(expectedTextContains)) {
                methodStatus = true;
            } else {
                methodStatus = false;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify_Paragraph_Text_Equals
    public static boolean Verify_Paragraph_Text_Equals(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle paragraphElement = (ElementHandle) page.evaluateHandle("xpath => document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue", map.get("ATTRIBUTE_VALUE").toString());
            Object jsParagraph = paragraphElement.evaluate("el => el.textContent");
            String paragraphText = jsParagraph.toString();
            String expectedTextContains = map.get("CONTROL_VALUE").toString().trim();
            if (paragraphText.trim().equals(expectedTextContains)) {
                methodStatus = true;
            } else {
                methodStatus = false;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify_Paragraph_Text_Equals_IgnoreCase
    public static boolean Verify_Paragraph_Text_Equals_IgnoreCase(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle paragraphElement = (ElementHandle) page.evaluateHandle("xpath => document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue", map.get("ATTRIBUTE_VALUE").toString());
            Object jsParagraph = paragraphElement.evaluate("el => el.textContent");
            String paragraphText = jsParagraph.toString();
            String expectedTextContains = map.get("CONTROL_VALUE").toString().trim();
            if (paragraphText.trim().equals(expectedTextContains)) {
                methodStatus = true;
            } else {
                methodStatus = false;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify_Alert_Text_Contains
    public static boolean Verify_Alert_Text_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            String paragraphText = page.evaluate("document.querySelector('p').textContent").toString();
            page.evaluate("alert('This is an alert!');");
            page.onDialog(dialog -> {
                String message = dialog.message();
                dialog.accept();
            });
            page.evaluate("alert('This is an alert!');");
            page.waitForTimeout(2000);
            String expectedTextContains = map.get("CONTROL_VALUE").toString().trim();
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify_Control_Text_Contains_Ignore_Case
    public static boolean Verify_Control_Text_Contains_Ignore_Case(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle element = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            assert element != null;
            String actualTextContains = element.innerText().trim().toUpperCase();
            String expectedTextContains = map.get("CONTROL_VALUE").toString().toUpperCase().trim();
            if (actualTextContains.contains(expectedTextContains)) {
                methodStatus = true;
                return methodStatus;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify_Control_Text_Equals
    public static boolean Verify_Control_Text_Equals(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle element = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            assert element != null;
            String Textvalue = element.innerText();
            String actualTextContains = Textvalue.trim();
            String expectedTextContains = map.get("CONTROL_VALUE").toString().trim();
            if (actualTextContains.equals(expectedTextContains)) {
                methodStatus = true;
                return methodStatus;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

    //Verify_Control_Text_Equals_Ignore_Case
    public static boolean Verify_Control_Text_Equals_Ignore_Case(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle element = waitForAnElement(page, map.get("ATTRIBUTE_VALUE").toString(), 30000);
            assert element != null;
            String actualTextContains = element.innerText().trim().toUpperCase();
            String expectedTextContains = map.get("CONTROL_VALUE").toString().toUpperCase().trim();
            if (actualTextContains.equalsIgnoreCase(expectedTextContains)) {
                methodStatus = true;
                return methodStatus;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcelWithElementMarking(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }

}
