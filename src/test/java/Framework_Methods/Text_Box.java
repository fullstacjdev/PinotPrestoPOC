package Framework_Methods;

import com.microsoft.playwright.*;

import javax.swing.*;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Scanner;

import static Framework_Methods.Generic_Methods.writeLogsInfoIntoExcel;
import static Framework_Methods.Web_Control_Common_Methods.*;

public class Text_Box {
    //Enter_Value_In_TextBox
    public static boolean Enter_Value_In_TextBox(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        Locator textBoxLocator = null;
        try {
            waitForPageLoad(page);
            textBoxLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            methodStatus = waitForPageLocator(page, textBoxLocator);
            if (methodStatus) {
                if (textBoxLocator.isVisible()) {
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "TextBox is Visible");
                    if (textBoxLocator.isEnabled()) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "TextBox is Enabled");
                        textBoxLocator.fill(map.get("CONTROL_VALUE").toString());
                    } else {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "TextBox is not Enabled");
                    }
                } else {
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "TextBox Link is not Visible");
                }
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Data has been successfully entered into TextBox");
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to enter data into TextBox");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert textBoxLocator != null;
            int elementCount = textBoxLocator.count();
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

    //Enter_Value_In_TextBox_Via_Command_Prompt
    public static boolean Enter_Value_In_TextBox_Via_Command_Prompt(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        Locator textBoxLocator = null;
        try {
            waitForPageLoad(page);
            textBoxLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            methodStatus = waitForPageLocator(page, textBoxLocator);
            if (methodStatus) {
                if (textBoxLocator.isVisible()) {
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "TextBox is Visible");
                    if (textBoxLocator.isEnabled()) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "TextBox is Enabled");
                        Scanner scanner = new Scanner(System.in);
                        String userEnteredInputFromCommandPrompt = scanner.nextLine();
                        textBoxLocator.fill(userEnteredInputFromCommandPrompt);
                    } else {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "TextBox is not Enabled");
                    }
                } else {
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "TextBox Link is not Visible");
                }
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Data has been successfully entered into TextBox from Command Prompt");
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to enter data into TextBox");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert textBoxLocator != null;
            int elementCount = textBoxLocator.count();
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

    //Enter_Value_In_TextBox_Via_Input_Box
    public static boolean Enter_Value_In_TextBox_Via_Input_Box(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        Locator textBoxLocator = null;
        try {
            waitForPageLoad(page);
            textBoxLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            methodStatus = waitForPageLocator(page, textBoxLocator);
            if (methodStatus) {
                if (textBoxLocator.isVisible()) {
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "TextBox is Visible");
                    if (textBoxLocator.isEnabled()) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "TextBox is Enabled");
                        String userEnteredInputFromInputBox = JOptionPane.showInputDialog(null, "Enter value:");
                        textBoxLocator.fill(userEnteredInputFromInputBox);
                    } else {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "TextBox is not Enabled");
                    }
                } else {
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "TextBox Link is not Visible");
                }
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Data has been successfully entered into TextBox from Input Box");
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to enter data into TextBox");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert textBoxLocator != null;
            int elementCount = textBoxLocator.count();
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

    //Enter_Value_In_TextBox_Via_Secure_Input_Box
    public static boolean Enter_Value_In_TextBox_Via_Secure_Input_Box(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        Locator textBoxLocator = null;
        try {
            waitForPageLoad(page);
            textBoxLocator = page.locator(map.get("ATTRIBUTE_VALUE").toString());
            methodStatus = waitForPageLocator(page, textBoxLocator);
            if (methodStatus) {
                if (textBoxLocator.isVisible()) {
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "TextBox is Visible");
                    if (textBoxLocator.isEnabled()) {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "TextBox is Enabled");
                        JPasswordField passwordField = new JPasswordField();
                        int option = JOptionPane.showConfirmDialog(null, passwordField, "Enter value:", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                        if (option == JOptionPane.OK_OPTION) {
                            char[] password = passwordField.getPassword();
                            String userEnteredInputFromInputBox = new String(password);
                            textBoxLocator.fill(userEnteredInputFromInputBox);
                        }
                    } else {
                        writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "TextBox is not Enabled");
                    }
                } else {
                    writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "TextBox Link is not Visible");
                }
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Data has been successfully entered into TextBox from Input Box");
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to enter data into TextBox");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            assert textBoxLocator != null;
            int elementCount = textBoxLocator.count();
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
}
