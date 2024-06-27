package Framework_Methods;

import com.microsoft.playwright.ElementHandle;
import com.microsoft.playwright.FileChooser;
import com.microsoft.playwright.Locator;
import com.microsoft.playwright.Page;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;

import static Framework_Methods.Generic_Methods.writeLogsInfoIntoExcel;
import static Framework_Methods.Web_Control_Common_Methods.*;

public class Upload {

    //Upload_File
    public static boolean Upload_File(int rownumber, Page page, HashMap map, String reportLogFileName) throws IOException {
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
            fileChooser.setFiles(Paths.get(System.getProperty("user.dir") + "\\src\\test\\java\\Upload_Test_Files\\Sample_Image_File.jpeg"));
            methodStatus = true;
        } catch (Exception e) {
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", e.toString());
            writeLogsInfoIntoExcel(reportLogFileName, "Logs_Info", String.valueOf(Thread.currentThread().getId()), "Error", Arrays.toString(e.getStackTrace()));
        } finally {
            writeTCsInfoIntoExcel(page, reportLogFileName, map, methodStatus);
        }
        return methodStatus;
    }



}
