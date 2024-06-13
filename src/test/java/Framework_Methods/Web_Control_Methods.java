package Framework_Methods;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.ElementState;
import com.microsoft.playwright.options.LoadState;
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

public class Web_Control_Methods {


    public static void waitForElementsEnabled(Page page) {
        // Define a JavaScript function to check if at least some elements are enabled
        String script = "(() => { " +
                "const allElements = document.querySelectorAll('*');" +
                "for (let element of allElements) {" +
                "if (!element.disabled && element.offsetWidth > 0 && element.offsetHeight > 0) {" +
                "return true;" +
                "}" +
                "}" +
                "return false;" +
                "})()";

        // Wait for the defined condition to become true with a timeout of 10 seconds
        try {
            page.waitForFunction(script, new Page.WaitForFunctionOptions().setTimeout(10000));
        } catch (PlaywrightException e) {
            //System.out.println("Timeout occurred while waiting for some elements to be enabled.");
        }

    }

    public static void verifyPageTitle(Page page, String expectedTitle, int timeoutInSeconds) throws InterruptedException {
        long startTime = System.currentTimeMillis();
        boolean isTitleMatched = false;

        while (!isTitleMatched && (System.currentTimeMillis() - startTime) < timeoutInSeconds * 1000L) {
            String actualTitle = page.title();
            if (actualTitle.contains(expectedTitle)) {
                //System.out.println("Page title matched: " + actualTitle);
                isTitleMatched = true;
            } else {
                //System.out.println("Page title not matched: " + actualTitle);
                // Sleep for a short interval before checking again

            }
            Thread.sleep(200); // 1 second interval
        }
        if (!isTitleMatched) {
            System.out.println("Page title didn't match within the timeout.");
        }
    }

    public static boolean waitForPageContentEnabled(Page page) {
        // Define a JavaScript function to check if the page content is enabled
        String script = "document.readyState === 'complete' && !document.querySelector('body').hasAttribute('disabled');";

        try {
            // Wait for the defined condition to become true
            page.waitForFunction(script, new Page.WaitForFunctionOptions().setTimeout(12000));
            return true; // Page content is enabled
        } catch (PlaywrightException e) {
            System.out.println("Error waiting for page content to be enabled: " + e.getMessage());
            return false; // Page content is not enabled within the timeout
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
                return true; // Load state reached successfully
            } catch (Exception e) {
                // Log or handle the error
                System.err.println("Attempt " + (currentAttempt + 1) + " failed: " + e.getMessage());
            }
            currentAttempt++; // Increment attempt counter
        }
        return false; // Load state not reached after retries
    }

    public static void waitForPageLoad(Page page) {
        Page.WaitForLoadStateOptions options = new Page.WaitForLoadStateOptions().setTimeout(120000); // Set a timeout of 30 seconds
        page.waitForLoadState(LoadState.LOAD, options);
        /*page.onPageError(event -> {
            System.out.println("Page error event triggered");
            System.out.println("Event object: " + event);
        });
        page.onConsoleMessage(message -> {
            System.out.println("Console message: " + message.text());
        });*/
    }

    // Wait for an element
    public static ElementHandle waitForAnElement(Page page, String elementString,int timeoutMillis) {
        long startTime = System.currentTimeMillis();
        ElementHandle element = null;
        while (System.currentTimeMillis() - startTime < timeoutMillis) {
            try {
                // Attempt to find the element
                element = page.querySelector(elementString);
                // If element is found and visible, return it
                if (element != null && element.isVisible()) {
                    //System.out.println("Element found!");
                    return element;
                }
            } catch (PlaywrightException e) {
                // Ignore exception if element is not found within the timeout
            }
            // Add a short delay before attempting to find the element again
            try {
                Thread.sleep(500); // Adjust the delay as needed
            } catch (InterruptedException e) {
                // Handle interruption if needed
            }
        }
        System.out.println("Element not found within the timeout.");
        return null;
    }
    public static String ScreenshotToBase64(Page page) throws IOException {
        SimpleDateFormat uniqueImageName = new SimpleDateFormat("yyyy_MMM_dd_HH_mm_ss_sss");
        String ImageFilePathAndName = System.getProperty("user.dir") + "\\src\\test\\java\\Images\\Image_" + uniqueImageName.format(Calendar.getInstance().getTime()) + ".PNG";
        // Capture screenshot
        byte[] screenshotBytes = page.screenshot();
        // Save screenshot to file
        Path screenshotPath = Paths.get(ImageFilePathAndName);
        Files.write(screenshotPath, screenshotBytes);
        return ImageFilePathAndName;
    }

    //Page_Load
    public static boolean Page_Load(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            page.navigate(map.get("Control_Value").toString());
            waitForPageLoad(page);
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Page has been loaded successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Page is not loaded");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Get_Page_Title_Equals
    public static boolean Get_Page_Title_Equals(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for all elements to be enabled
            waitForElementsEnabled(page);
            // Expected page title
            String expectedTitle = map.get("Control_Value").toString().toUpperCase().trim();
            // Verify page title within a time limit
            verifyPageTitle(page, expectedTitle, 30); // 30 seconds timeout
            // Get the page title
            String actualPageTitle = page.title().trim().toUpperCase();
            if (actualPageTitle.equals(expectedTitle)) {
                methodStatus = true;
                //return methodStatus;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Page title matched successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "The page is not matched correctly");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Get_Page_Title_Contains
    public static boolean Get_Page_Title_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for all elements to be enabled
            waitForElementsEnabled(page);
            // Expected page title
            String expectedTitle = map.get("Control_Value").toString().toUpperCase().trim();
            // Verify page title within a time limit
            verifyPageTitle(page, expectedTitle, 30); // 30 seconds timeout
            // Get the page title
            String actualPageTitle = page.title().trim().toUpperCase();
            if (actualPageTitle.contains(expectedTitle)) {
                methodStatus = true;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Page title matched successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "The page is not matched correctly");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Enter_Value_In_TextBox
    public static boolean Enter_Value_In_TextBox(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            //Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            //Focus on element
            page.focus(map.get("Attribute_Value").toString());
            //Check whether textbox control is displayed and enabled
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                // Click on the text box
                elementTextBox.click();
                //Value from Excel sheet
                elementTextBox.type(map.get("Control_Value").toString());
                //return method status
                methodStatus = true;
                //System.out.println("Control is displayed on the page");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Value successfully entered in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to enter the value in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Enter_Value_In_TextBox_Via_Command_Prompt
    public static boolean Enter_Value_In_TextBox_Via_Command_Prompt(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(12000); // Set a timeout of 10 seconds
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            //Focus on element
            page.focus(map.get("Attribute_Value").toString());
            //Check whether textbox control is displayed and enabled
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                // Click on the text box
                elementTextBox.click();
                // Create a scanner object to read input from the command line
                Scanner scanner = new Scanner(System.in);
                // Prompt the user to enter a value
                //System.out.print("Enter value for " + map.get("Action_Flow").toString() + " : ");
                String value = scanner.nextLine();
                //Value from Command Prompt
                elementTextBox.type(value);
                //return method status
                methodStatus = true;
                //System.out.println("Control is displayed on the page");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Value successfully entered in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to enter the value in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Enter_Value_In_TextBox_Via_Input_Box
    public static boolean Enter_Value_In_TextBox_Via_Input_Box(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            //Focus on element
            page.focus(map.get("Attribute_Value").toString());
            //Check whether textbox control is displayed and enabled
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                // Click on the text box
                elementTextBox.click();
                // Display input dialog and get the user input
                String userInput = JOptionPane.showInputDialog(null, "Enter value for " + map.get("Action_Flow").toString() + " : ");
                //Value from Input Box
                elementTextBox.type(userInput);
                //return method status
                methodStatus = true;
                //System.out.println("Control is displayed on the page");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Value successfully entered in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to enter the value in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Enter_Value_In_TextBox_Via_Secure_Input_Box
    public static boolean Enter_Value_In_TextBox_Via_Secure_Input_Box(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            page.focus(map.get("Attribute_Value").toString());
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                // Click on the text box
                elementTextBox.click();
                // Create a password input dialog
                JPasswordField passwordField = new JPasswordField();
                int option = JOptionPane.showConfirmDialog(null, passwordField, "Enter value for " + map.get("Action_Flow").toString(), JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                // Check if user clicked OK
                if (option == JOptionPane.OK_OPTION) {
                    // Get the password entered by the user
                    char[] password = passwordField.getPassword();
                    // Convert char array to String if needed
                    String userInput = new String(password);
                    // Now userInput contains the secure input from the user
                    //Value from Input Box
                    elementTextBox.type(userInput);
                }
                //return method status
                methodStatus = true;
                //System.out.println("Control is displayed on the page");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Value successfully entered in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to enter the value in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) && ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Enter_Value_In_Frame_TextBox
    public static boolean Enter_Value_In_Frame_TextBox(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            //Focus on element
            page.focus(map.get("Attribute_Value").toString());
            //Check whether textbox control is displayed and enabled
            if (elementTextBox.isVisible() && elementTextBox.isEnabled()) {
                // Click on the text box
                //elementTextBox.click();
                if (map.get("Control_Type").toString().equalsIgnoreCase("SensitiveTextBox")) {
                    // Create a scanner object to read input from the command line
                    Scanner scanner = new Scanner(System.in);
                    // Prompt the user to enter a value
                    //System.out.print("Enter value for " + map.get("Action_Flow").toString() + " : ");
                    String value = scanner.nextLine();
                    //Value from Command Prompt
                    elementTextBox.type(value);
                } else {
                    //Value from Excel sheet
                    elementTextBox.type(map.get("Control_Value").toString());
                }
                //return method status
                methodStatus = true;
                //System.out.println("Control is displayed on the page");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Value successfully entered in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to enter the value in textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Get_Value_From_TextBox
    public static boolean Get_Value_From_TextBox(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            // Get the initial scroll position
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            // Simulate pressing the "Page Down" key until we reach the bottom of the page
            Keyboard keyboard = page.keyboard();
            String strGetValueFromTextBox = null;
            while (true) {
                // Wait for a short moment to let the page scroll
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                // Check if the scroll position hasn't changed
                if (newScrollY.equals(initialScrollY)) {
                    //Check whether textbox control is displayed and enabled
                    if (elementTextBox.isVisible()) {
                        strGetValueFromTextBox = elementTextBox.inputValue();
                        methodStatus = true;
                        break;
                    } else {
                        //System.out.println("Control is not displayed on the page");
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
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Able to get value from textbox successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to get the value from textbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Button_Enable_Status
    public static boolean Button_Enable_Status(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementButton = waitForAnElement(page, map.get("Attribute_Value").toString(), 12000);
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementButton = page.waitForSelector(elementPath);
            } else {
                elementButton = page.waitForSelector(map.get("Attribute_Value").toString());
            }
            // Get the initial scroll position
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            // Simulate pressing the "Page Down" key until we reach the bottom of the page
            Keyboard keyboard = page.keyboard();
            while (true) {
                // Wait for a short moment to let the page scroll
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                // Check if the scroll position hasn't changed
                if (newScrollY.equals(initialScrollY)) {
                    //Check whether textbox control is displayed and enabled
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
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Button has been enabled");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Button is not enabled");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Button_Visible_Status
    public static boolean Button_Visible_Status(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementButton = waitForAnElement(page, map.get("Attribute_Value").toString(), 12000);
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementButton = page.waitForSelector(elementPath);
            } else {
                elementButton = page.waitForSelector(map.get("Attribute_Value").toString());
            }
            // Get the initial scroll position
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            // Simulate pressing the "Page Down" key until we reach the bottom of the page
            Keyboard keyboard = page.keyboard();
            while (true) {
                // Wait for a short moment to let the page scroll
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                // Check if the scroll position hasn't changed
                if (newScrollY.equals(initialScrollY)) {
                    //Check whether textbox control is displayed and enabled
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
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Able to see the button");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to see the button");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Button_Click
    public static boolean Button_Click(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementButton = waitForAnElement(page, map.get("Attribute_Value").toString(), 12000);
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementButton = page.waitForSelector(elementPath);
            } else {
                elementButton = page.waitForSelector(map.get("Attribute_Value").toString());
            }
            // Get the initial scroll position
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            // Simulate pressing the "Page Down" key until we reach the bottom of the page
            Keyboard keyboard = page.keyboard();
            while (true) {
                // Wait for a short moment to let the page scroll
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                // Check if the scroll position hasn't changed
                if (newScrollY.equals(initialScrollY)) {
                    //Check whether textbox control is displayed and enabled
                    if (elementButton.isVisible() && elementButton.isEnabled()) {
                        // Click on the text box
                        elementButton.click();
                        //System.out.println("Control is displayed on the page");
                        //return method status
                        methodStatus = true;
                        break;
                    } else {
                        //System.out.println("Control is not displayed on the page");
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
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Button has been clicked successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to click the button");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    // Function to check if the element is visible
    private static void isElementVisible(Page page, ElementHandle element) {
        // Create an instance of JavascriptExecutor
        JavascriptExecutor js = (JavascriptExecutor) page;
        // Scroll until the element is found or until reaching the end of the page
        // WebElement element = null;
        boolean elementFound = false;
        int scrollStartPosition = 0;
        int scrollEndPosition = 0;
        int maxScroll = 5000; // Maximum scroll position to prevent infinite scrolling
        while (!elementFound && scrollStartPosition <= maxScroll) {
            // Scroll down the page
            js.executeScript("window.scrollTo(scrollStartPosition, " + scrollEndPosition + ")");
            // Check if the element is available
            // element = driver.findElement(By.id("your-element-id")); // Change to your element's locator
            if (element.isVisible()) {
                elementFound = true;
            }
            // Increment scroll position for the next scroll
            scrollStartPosition = scrollStartPosition + scrollEndPosition;
            scrollEndPosition += 500; // You can adjust the scroll amount as needed
        }
        // Perform actions with the element if found
        if (elementFound) {
            System.out.println("Element found!");
            // Perform actions with the element
        } else {
            System.out.println("Element not found after scrolling.");
        }
    }

    //Page_Down
    public static boolean Page_Down(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Get the initial scroll position
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            // Simulate pressing the "Page Down" key until we reach the bottom of the page
            Keyboard keyboard = page.keyboard();
            while (true) {
                keyboard.press("PageDown");
                // Wait for a short moment to let the page scroll
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                // Check if the scroll position hasn't changed
                if (newScrollY.equals(initialScrollY)) {
                    //page.waitForTimeout(1000);
                    return methodStatus = true;
                }
                initialScrollY = newScrollY;
            }
        } catch (Exception e) {
            System.out.println();
        }

        return methodStatus;
    }

    //Link_Click
    public static boolean Link_Click(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            ElementHandle elementLink = waitForAnElement(page, map.get("Attribute_Value").toString(), 12000);
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementLink = page.waitForSelector(elementPath);
            } else {
                elementLink = page.waitForSelector(map.get("Attribute_Value").toString());
            }
            // Verify if the control is displayed
            boolean isDisplayed = elementLink.isVisible();
            //Verify if the control is enabled
            boolean isEnabled = elementLink.isEnabled();
            //Check whether textbox control is displayed and enabled
            if (isDisplayed && isEnabled) {
                // Click on the text box
                elementLink.click();
                //return method status
                methodStatus = true;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Link has been clicked successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to click the link");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Click on Tab
    public static boolean Tab_Click(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds

            ElementHandle elementLink=waitForAnElement(page, map.get("Attribute_Value").toString(), 12000);

            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementLink = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementLink = page.waitForSelector(map.get("Attribute_Value").toString(), optionsForElementLoad);
            }
            // Click on the text box
            elementLink.click();
            //return method status
            methodStatus = true;
            //System.out.println("Control is displayed on the page");
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Tab has been clicked successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to click the tab");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Frame_Link_Click
    public static boolean Frame_Link_Click(int rownumber, Page page, HashMap map, String reportLogFileName) throws InterruptedException, IOException {
        //Method return status default value
        boolean methodStatus = false;
        //Thread.sleep(5000);
        System.out.println(page.title());
        Frame innerFrame = page.frameByUrl("https://selenium143.blogspot.com/");
        innerFrame.locator(map.get("Attribute_Value").toString()).click();
        //newTab.close();
        page.bringToFront();
        methodStatus = true;
        return methodStatus;
    }

    //Get_Header_Title_Equals
    public static boolean Get_Header_Title_Equals(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds

            ElementHandle elementActualHeaderTitle = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            if (elementActualHeaderTitle.isVisible()) {
                // Get the page title
                String actualHeaderTitle = elementActualHeaderTitle.textContent();
                // Expected page title
                String expectedHeaderTitle = map.get("Control_Value").toString();
                if (actualHeaderTitle.equals(expectedHeaderTitle)) {
                    methodStatus = true;
                    return methodStatus;
                }
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Get text successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to get the text");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Get_Header_Title_Contains
    public static boolean Get_Header_Title_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            ElementHandle elementActualHeaderTitle = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            assert elementActualHeaderTitle != null;
            if (elementActualHeaderTitle.isVisible()) {
                // Get the page title
                String actualHeaderTitle = elementActualHeaderTitle.textContent();
                // Expected page title
                String expectedHeaderTitle = map.get("Control_Value").toString();
                if (actualHeaderTitle.contains(expectedHeaderTitle)) {
                    methodStatus = true;
                    return methodStatus;
                }
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Get the text successfully from control");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to get the text from control");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Select_Value_From_Dropdown
    public static boolean Select_Value_From_Dropdown(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
           // WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle selectDropdown = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            // Dropdown element
            //Locator selectDropdown = page.locator(map.get("Attribute_Value").toString());
            //Select required value from dropdown
            assert selectDropdown != null;
            selectDropdown.selectOption(map.get("Control_Value").toString());
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Able to select the value from dropdown successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to select the value from dropdown");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Select_Value_From_Web_Dropdown
    public static boolean Select_Value_From_Web_Dropdown(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
           // WaiForAnElement(page, map.get("Attribute_Value").toString());
            //Dropdown_Value_Locator=
            String[] dropdownLocators = map.get("Attribute_Value").toString().split("<<=>>");
            String dropdownLocator = dropdownLocators[0].trim();
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            ElementHandle elementDropdown = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            String dropdownSelectValue = dropdownLocators[1].trim();
            //Dynamic element preparation
            String elementPath;
            if (dropdownSelectValue.contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                elementPath = (dropdownSelectValue).replace("%s", dynamicValue);
            } else {
                // Example dynamic value
                elementPath = dropdownSelectValue;
            }
            // Verify if the control is displayed
            boolean isDisplayed = elementDropdown.isVisible();
            //Verify if the control is enabled
            boolean isEnabled = elementDropdown.isEnabled();
            //Check whether textbox control is displayed and enabled
            if (isDisplayed && isEnabled) {
                // Click on the dropdown to open it
                elementDropdown.click();
                // Wait for a brief moment to ensure the dropdown options are visible
                page.waitForTimeout(1000); // Adjust the timeout value as needed
                // Select the desired option from the dropdown
                ElementHandle option = page.querySelector(elementPath);
                Thread.sleep(1000);
                option.click();
                methodStatus = true;
                //System.out.println("Control is displayed on the page");
            } else {
                System.out.println("Control is not displayed on the page");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Able to select the value from dropdown successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to select the value from dropdown");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Wait_For_Short_Time
    public static boolean Wait_For_Short_Time(int rownumber, Page page, HashMap map, String reportLogFileName) {
        //Method return status default value
        boolean methodStatus = false;
        try {
            Thread.sleep(5000);
            methodStatus = true;
        } catch (Exception e) {
            System.out.println(e.toString());
        }
        return methodStatus;
    }

    //Wait_For_Short_Time
    public static boolean Wait_For_Medium_Time(int rownumber, Page page, HashMap map, String reportLogFileName) {
        //Method return status default value
        boolean methodStatus = false;
        try {
            Thread.sleep(7500);
            methodStatus = true;
        } catch (Exception e) {
            System.out.println(e.toString());
        }
        return methodStatus;
    }

    //Wait_For_Short_Time
    public static boolean Wait_For_Long_Time(int rownumber, Page page, HashMap map, String reportLogFileName) {
        //Method return status default value
        boolean methodStatus = false;
        try {
            Thread.sleep(10000);
            methodStatus = true;
        } catch (Exception e) {
            System.out.println(e.toString());
        }
        return methodStatus;
    }

    //Accept_Or_Reject_Cookies
    public static boolean Accept_Or_Reject_Cookies(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementForCookies = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            elementForCookies.click();
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Cookies selected successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to select the Cookies");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Select_Radio_Button
    public static boolean Select_Radio_Button(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementRadioButton=waitForAnElement(page, map.get("Attribute_Value").toString(), 30000);;
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementRadioButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementRadioButton = page.waitForSelector(map.get("Attribute_Value").toString(), optionsForElementLoad);
            }
            // Verify if the control is displayed
            boolean isDisplayed = elementRadioButton.isVisible();
            //Verify if the control is enabled
            boolean isEnabled = elementRadioButton.isEnabled();
            //Check whether textbox control is displayed and enabled
            if (isDisplayed && isEnabled) {
                // Click on the text box
                elementRadioButton.check();
                //return method status
                methodStatus = true;
                //System.out.println("Control is displayed on the page");
            } else {
                System.out.println("Control is not displayed on the page");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Check the Radio button successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to check the Radio button");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Select_Checkbox_Button
    public static boolean Select_Checkbox_Button(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementCheckboxButton=waitForAnElement(page, map.get("Attribute_Value").toString(), 30000);;
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementCheckboxButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementCheckboxButton = page.waitForSelector(map.get("Attribute_Value").toString(), optionsForElementLoad);
            }
            // Verify if the control is displayed
            boolean isDisplayed = elementCheckboxButton.isVisible();
            //Verify if the control is enabled
            boolean isEnabled = elementCheckboxButton.isEnabled();
            //Check whether textbox control is displayed and enabled
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
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Able to check Checkbox successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to check the Checkbox");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Handle_PopUp
    public static boolean Handle_PopUp(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementRadioButton=waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);;
            methodStatus = true;
            System.out.println("Handle_PopUp");
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Handled Pop Up successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to handle Pop Up");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //To upload File or Image to the page from File Explore
    public static boolean Upload_File_Or_Image_To_Page_From_File_Explore(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementButton=waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);;
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementButton = page.waitForSelector(map.get("Attribute_Value").toString(), optionsForElementLoad);
            }
            FileChooser fileChooser = page.waitForFileChooser(() -> {
                page.getByText(map.get("Control_Value").toString()).click();
            });
            fileChooser.setFiles(Paths.get(System.getProperty("user.dir") + "\\src\\test\\java\\Upload_Test_Files\\Dharanija.jpeg"));
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "File uploaded successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to upload the file");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Verify Button Name
    public static boolean Verify_Button_Name(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementButton=waitForAnElement(page, map.get("Attribute_Value").toString(), 30000);;
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementButton = page.waitForSelector(map.get("Attribute_Value").toString(), optionsForElementLoad);
            }
            if ((elementButton.innerText()).equals(map.get("Control_Value").toString())) {
                elementButton.click();
            } else {
                //System.out.println("The video is already saved");
            }
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Button name checked successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Unable to check the button name");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Verify Share Option Popup
    public static boolean Verify_Share_Option_PopUp(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementButton=waitForAnElement(page, map.get("Attribute_Value").toString(), 30000);;
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementButton = page.waitForSelector(map.get("Attribute_Value").toString(), optionsForElementLoad);
            }
            if (elementButton.isVisible()) {
                //System.out.println("Share PopUp is Visible");
                methodStatus = true;
            } else {
                //System.out.println("Share option is not showing");
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Page has been loaded successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Page is not loaded");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Share to User
    public static boolean Share_To_User(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementButton=waitForAnElement(page, map.get("Attribute_Value").toString(), 30000);;
            if ((map.get("Attribute_Value").toString()).contains("%s")) {
                // Example dynamic value
                String dynamicValue = map.get("Control_Value").toString();
                // Construct the XPath with dynamic value
                String elementPath = String.format(map.get("Attribute_Value").toString(), dynamicValue);
                //String elementPath = (map.get("Attribute_Value").toString()).replace("%s", dynamicValue);
                elementButton = page.waitForSelector(elementPath, optionsForElementLoad);
            } else {
                elementButton = page.waitForSelector(map.get("Attribute_Value").toString(), optionsForElementLoad);
            }
            if (elementButton.isVisible()) {
                if (elementButton.innerText().equals(map.get("Control_Value").toString())) {
                    //System.out.println("Video or Course is already shared with the user");
                    page.keyboard().press("Escape");
                }
            } else {
                page.getByPlaceholder("Search").click();
                page.getByPlaceholder("Search").fill(map.get("Control_Value").toString());
                if (elementButton.innerText().equals(map.get("Control_Value").toString())) {
                    elementButton.click();
                }
            }
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Page has been loaded successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Page is not loaded");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Get_Hint_Text_From_Control_Equals_To
    public static boolean Get_Hint_Text_From_Control_Equals_To(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            // Get the initial scroll position
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            // Simulate pressing the "Page Down" key until we reach the bottom of the page
            Keyboard keyboard = page.keyboard();
            String strGetValueFromTextBox = null;
            while (true) {
                // Wait for a short moment to let the page scroll
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                // Check if the scroll position hasn't changed
                if (newScrollY.equals(initialScrollY)) {
                    //Check whether textbox control is displayed and enabled
                    if (elementTextBox.isVisible()) {
                        // strGetValueFromTextBox = elementTextBox.getAttribute("validationMessage");
                        // Execute JavaScript to get the validationMessage attribute
                        String validationMessage = (String) elementTextBox.evaluate("element => element.validationMessage");
                        if (validationMessage.equals(map.get("Control_Value").toString())) {
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
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Hint value matched successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Hint value unmatched");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Get_Hint_Text_From_Control_Contains
    public static boolean Get_Hint_Text_From_Control_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        //Method return status default value
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for an element
            Page.WaitForSelectorOptions optionsForElementLoad = new Page.WaitForSelectorOptions().setTimeout(60000); // Set a timeout of 10 seconds
            //WaiForAnElement(page, map.get("Attribute_Value").toString());
            ElementHandle elementTextBox = waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            // Get the initial scroll position
            Number initialScrollY = (Number) page.evaluate("() => window.scrollY");
            // Simulate pressing the "Page Down" key until we reach the bottom of the page
            Keyboard keyboard = page.keyboard();
            String strGetValueFromTextBox = null;
            while (true) {
                // Wait for a short moment to let the page scroll
                page.waitForTimeout(1000);
                Number newScrollY = (Number) page.evaluate("() => window.scrollY");
                // Check if the scroll position hasn't changed
                if (newScrollY.equals(initialScrollY)) {
                    //Check whether textbox control is displayed and enabled
                    if (elementTextBox.isVisible()) {
                        // strGetValueFromTextBox = elementTextBox.getAttribute("validationMessage");
                        // Execute JavaScript to get the validationMessage attribute
                        String validationMessage = (String) elementTextBox.evaluate("element => element.validationMessage");
                        if (validationMessage.contains(map.get("Control_Value").toString())) {
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
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Hint value matched successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Hint value unmatched");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }

    //Get_Text_Contains
    public static boolean Get_Text_Contains(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
        boolean methodStatus = false;
        try {
            waitForPageLoad(page);
            // Wait for all elements to be enabled
            //waitForElementsEnabled(page);
            // Expected page title
            //String expectedTextContains = map.get("Control_Value").toString().toUpperCase().trim();
            // Verify page title within a time limit
            ElementHandle element =    waitForAnElement(page, map.get("Attribute_Value").toString(), 10000);
            //page.waitForTimeout(2000);
            // Get the Text Contains
            assert element != null;
            String Textvalue = element.innerText();
            String actualTextContains = Textvalue.trim().toUpperCase();
            // Expected Text Contains
            String expectedTextContains = map.get("Control_Value").toString().toUpperCase().trim();
            if (actualTextContains.contains(expectedTextContains)) {
                methodStatus = true;
                return methodStatus;
            }
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            if (methodStatus) {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Info", "Get value matched successfully");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Pass");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Pass");
                }
            } else {
                writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", "Get value unmatched");
                if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("YES")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), ScreenshotToBase64(page), "Fail");
                } else if (((!map.get("Test_Cases").toString().isEmpty()) && (map.get("Action_Flow").toString().equals("Assertion"))) & ((map.get("Should_Take_ScreenShot").toString().equals("") || map.get("Should_Take_ScreenShot").toString().equalsIgnoreCase("NO")))) {
                    writeTestCasesInfoIntoExcel(reportLogFileName, "TestCases_Info", String.valueOf(Thread.currentThread().getId()), map.get("Test_Cases").toString(), "", "Fail");
                }
            }
        }
        return methodStatus;
    }
}
