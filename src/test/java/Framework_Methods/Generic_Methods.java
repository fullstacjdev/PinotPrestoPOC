package Framework_Methods;

import com.microsoft.playwright.BrowserType;
import com.microsoft.playwright.Playwright;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.awt.*;
import java.awt.event.KeyEvent;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Generic_Methods {

    public static String[] ListFilesExample() {
        // Specify the path to the folder
        String folderPath = System.getProperty("user.dir") + "\\src\\test\\java\\Test_Data";
        // Create a File object representing the folder
        File folder = new File(folderPath);
        // Check if the folder exists
        String[] filenames = new String[0];
        if (folder.exists() && folder.isDirectory()) {
            // List all files in the folder
            File[] files = folder.listFiles();
            if (files != null) {
                // Extract filenames from File objects and put them into an array
                filenames = Arrays.stream(files)
                        .map(File::getName)
                        .toArray(String[]::new);
                // Print the filenames or perform any other operation
                //System.out.println("Filenames in the folder:");
                for (String filename : filenames) {
                    //System.out.println(filename);
                }
            } else {
                System.out.println("No files found in the folder.");
            }
        } else {
            System.out.println("Invalid folder path or folder does not exist.");
        }
        return filenames;
    }

    public static void htmlReportCreate(String fileName) {
        String htmlContent = "<!DOCTYPE html>\n" +
                "<html lang=\"en\">\n" +
                "<head>\n" +
                "    <meta charset=\"UTF-8\">\n" +
                "    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n" +
                "    <title>Automation_Execution_Report</title>\n" +
                "    <script src=\"https://cdn.jsdelivr.net/npm/chart.js\"></script>\n" +
                "    <script src=\"https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js\"></script>\n" +
                "    <script src=\"https://cdn.jsdelivr.net/npm/chartjs-plugin-chartjs-3d\"></script>\n" +
                "    <script type=\"text/javascript\" src=\"https://www.gstatic.com/charts/loader.js\"></script>\n" +
                "    <style>\n" +
                "        body, html {zoom: 100%;height: 99%;margin: 0;padding: 0;overflow: hidden; /* Disable scrolling on body */}\n" +
                "        .container {display: flex;height: 100%;}\n" +
                "        .right-panel {height: 68%;flex: 1;border: 0px solid #ccc;padding: 10px;}\n" +
                "        .left-panel {height: 68%;flex: 1;border: 0px solid #ccc;padding: 10px;overflow-y: hidden;}\n" +
                "        .horizontal-top-panel {height: 1%; /* Fixed height for the horizontal panel */ /* background-color: #f0f0f0;*/border-bottom: 0px solid #ccc;display: flex;justify-content: left;align-items: center;}\n" +
                "        .horizontal-bottom-panel {height: 28%; /* Fixed height for the horizontal panel */ /* background-color: #f0f0f0;*/border-bottom: 0px solid #ccc;display: flex;justify-content: left;align-items: center;}\n" +
                "        .horizontal-bottom-panel-first {width: 1%; /* Fixed height for the horizontal panel */ /* background-color: #f0f0f0;*/ /*border-bottom: 1px solid #ccc;*/display: flex;justify-content: left;align-items: center;}\n" +
                "        .horizontal-bottom-panel-second {width: 29%; /* Fixed height for the horizontal panel */ /* background-color: #f0f0f0;*/ /*border-bottom: 1px solid #ccc;*/display: flex;justify-content: left;align-items: center;}\n" +
                "        .horizontal-bottom-panel-third {width: 15%; /* Fixed height for the horizontal panel */ /*background-color: #f0f0f0;*/ /*border-bottom: 1px solid #ccc;*/display: flex;justify-content: left;align-items: center;}\n" +
                "        .horizontal-bottom-panel-fourth {width: 15%; /* Fixed height for the horizontal panel */ /* background-color: #f0f0f0;*/ /*border-bottom: 1px solid #ccc;*/display: flex;justify-content: left;align-items: center;}\n" +
                "        .horizontal-last-panel {width: 40%; /* Fixed height for the horizontal panel */ /* background-color: #f0f0f0;*/ /*border-bottom: 1px solid #ccc;*/display: flex;flex-direction: column;align-items: center;}\n" +
                "        .tab {\n" +
                "            overflow-y: auto;\n" +
                "            border: 0px solid #ccc;\n" +
                "            background-color: transparent;\n" +
                "        }\n" +
                "\n" +
                "        .tab button {\n" +
                "            background-color: darkblue;\n" +
                "            border: none;\n" +
                "            outline: none;\n" +
                "            cursor: pointer;\n" +
                "            padding: 6px 9px;\n" +
                "            margin-right: 5px;\n" +
                "            font-size: 10px;\n" +
                "            color: white; /* Text color */\n" +
                "            border-radius: 8px; /* Rounded corners */\n" +
                "            position: relative;\n" +
                "            transition: color 0.3s ease, background-color 0.3s ease;\n" +
                "        }\n" +
                "\n" +
                "        .tab button.active {\n" +
                "            background-color: #007bff; /* Active background color */\n" +
                "            color: white; /* Active text color */\n" +
                "            box-shadow: 0 2px 2px rgba(0, 0, 0, 0.1); /* Shadow effect */\n" +
                "        }\n" +
                "\n" +
                "        .tab button:not(.active):hover {\n" +
                "            background-color: #f0f0f0; /* Hover background color */\n" +
                "            color: #333; /* Hover text color */\n" +
                "        }\n" +
                "\n" +
                "        .tab button.active::before {\n" +
                "            content: '';\n" +
                "            position: absolute;\n" +
                "            bottom: 0;\n" +
                "            left: 50%;\n" +
                "            transform: translateX(-50%);\n" +
                "            width: 50%;\n" +
                "            height: 3px;\n" +
                "            background-color: #fff; /* Indicator line color */\n" +
                "        }\n" +
                "\n" +
                "        .tabcontent {\n" +
                "            display: block; /* Ensure the content is displayed */\n" +
                "            width: 100%; /* Ensure the content occupies 100% of the available width */\n" +
                "            max-height: 96%; /* Adjust 100px as needed */\n" +
                "            overflow-y: auto; /* Enable vertical scrolling when content exceeds max-height */\n" +
                "            padding: 1px; /* Adjust padding as needed */\n" +
                "            border: 0px solid #ccc;\n" +
                "            border-top: none;\n" +
                "            box-sizing: border-box; /* Ensure padding and border are included in width calculation */\n" +
                "        }\n" +
                "\n" +
                "        /* Additional specificity to ensure styles are applied */\n" +
                "        body .tabcontent {\n" +
                "            display: none;\n" +
                "        }\n" +
                "        body .tabcontent {display: none;}\n" +
                "        table {width: 99%;border-collapse: collapse;border: 2px solid #ddd; /* Border around the table */font-family: 'Verdana', serif; /* Specify the font family */font-size: 12px; /* Set the font size */margin-bottom: 05px; /* Adjust the value as needed */}\n" +
                "        th {\n" +
                "            border: 1px solid #ddd; /* Border around header cells */\n" +
                "            padding: 5px;\n" +
                "            text-align: center;\n" +
                "            font-family: 'Verdana', serif; /* Specify the font family */\n" +
                "            font-size: 11px; /* Set the font size */\n" +
                "            background-color: #00008B; /* Background color for header cells */\n" +
                "            color: white; /* Text color for header cells */\n" +
                "        }\n" +
                "        td {border: 1px solid #ddd; /* Border around data cells */padding: 5px;text-align: center;}\n" +
                "        a {text-decoration: none; /* Remove underline */}\n" +
                "        .toggle-table {display: none; /* Hide by default */}\n" +
                "        .pagination {display: inline-block;margin-top: 0px;font-family: Arial, sans-serif;}\n" +
                "        .pagination a {color: black;float: left;padding: 0px 7px;text-decoration: none;transition: background-color .3s;border: 1px solid #ddd;font-weight: bold;}\n" +
                "        .pagination a.active {background-color: #4CAF50;color: white;border: 1px solid #4CAF50;}\n" +
                "        .pagination a:hover:not(.active) {background-color: #ddd;}\n" +
                "        .horizontal-table .header {background-color: #00008B;color: white;padding: 8px;text-align: center;flex: 1;border-right: 1px solid #ddd;}\n" +
                "        .horizontal-table .data {flex: 2;padding: 8px;text-align: center;border-right: 1px solid #ddd;}\n" +
                "        .horizontal-table .data:last-child {border-right: none; /* Remove border for last data column */}\n" +
                "h1 {\n" +
                "            font-size: 24px; /* Increased font size for better prominence */\n" +
                "            font-weight: bold;\n" +
                "            color: darkblue; /* Adjusted color to a shade of blue */\n" +
                "            text-transform: uppercase;\n" +
                "            letter-spacing: 2px;\n" +
                "            text-align: center;\n" +
                "            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.4); /* Enhanced soft shadow for depth */\n" +
                "            margin: 0;\n" +
                "            padding: 16px 24px; /* Increased padding for better spacing */\n" +
                "            /*background-color: #fff; /* White background for the header */\n" +
                "            /*border-radius: 12px; /* Slightly increased border radius for smoother corners */\n" +
                "            /*box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Soft shadow for depth */\n" +
                "            transition: all 0.3s ease; /* Smooth transition for hover effects */\n" +
                "        }\n" +
                "\n" +
                "        h1:hover {\n" +
                "            transform: scale(1.05); /* Scale effect on hover for a subtle interactive feel */\n" +
                "            /*box-shadow: 0 8px 12px rgba(0, 0, 0, 0.2); /* Enhanced shadow on hover */\n" +
                "        }\n"+
                ".hint {\n" +
                "            position: relative;\n" +
                "            cursor: pointer;\n" +
                "        }\n" +
                "\n" +
                "        .hint::after {\n" +
                "            content: attr(data-hint);\n" +
                "            position: fixed;\n" +
                "            background-color: dimgray;\n" +
                "            color: #fff;\n" +
                "            padding: 9px;\n" +
                "            border-radius: 9px;\n" +
                "            bottom: 70%;\n" +
                "            left: 15.5%;\n" +
                "            transform: translateX(-30%);\n" +
                "            opacity: 0;\n" +
                "            white-space: initial; /* Allow text to wrap */\n" +
                "            width: 1500px; /* Adjust width based on content */\n" +
                "            max-width: 680px; /* Limit maximum width */\n" +
                "            font-size: 10px;\n" +
                "            line-height: 2; /* Adjust line height for readability */\n" +
                "            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);\n" +
                "            z-index: 5; /* Ensure tooltip is above other elements */\n" +
                "        }\n" +
                "\n" +
                "        .hint:hover::after {\n" +
                "            opacity: 1; /* Show tooltip on hover */\n" +
                "        }\n"+
                "    </style>\n" +
                "</head>\n" +
                "<body>\n" +
                "<div class=\"horizontal-top-panel\"></div>\n" +
                "<div class=\"horizontal-bottom-panel\">\n" +
                "    <div class=\"horizontal-bottom-panel-first\"></div>\n" +
                "    <div class=\"horizontal-bottom-panel-second\">\n" +
                "        <table>\n" +
                "            <tr>\n" +
                "                <th style=\"width: 100px; text-align: left; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">E</span>XECUTED\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">B</span>Y\n" +
                "                </th>\n" +
                "                <td style=\"width: 100px; text-align: left;\">Executed_By</td>\n" +
                "            </tr>\n" +
                "            <tr>\n" +
                "                <th style=\"width: 100px; text-align: left; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">E</span>XECUTION\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">D</span>ATE\n" +
                "                </th>\n" +
                "                <td style=\"width: 100px; text-align: left;\">Execution_Date</td>\n" +
                "            </tr>\n" +
                "            <tr>\n" +
                "                <th style=\"width: 100px; text-align: left; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">T</span>OTAL\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">E</span>XECUTIONS\n" +
                "                </th>\n" +
                "                <td style=\"width: 100px; text-align: left;\">Total_Summary_Executions</td>\n" +
                "            </tr>\n" +
                "            <tr>\n" +
                "                <th style=\"width: 100px; text-align: left; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">P</span>ASSED\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">E</span>XECUTIONS\n" +
                "                </th>\n" +
                "                <td style=\"width: 100px; text-align: left;\">Passed_Summary_Executions</td>\n" +
                "            </tr>\n" +
                "            <tr>\n" +
                "                <th style=\"width: 100px; text-align: left; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">F</span>AILED\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">E</span>XECUTIONS\n" +
                "                </th>\n" +
                "                <td style=\"width: 100px; text-align: left;\">Failed_Summary_Executions</td>\n" +
                "            </tr>\n" +
                "            <tr>\n" +
                "                <th style=\"width: 100px; text-align: left; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">E</span>XPORT\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">E</span>XECUTIONS\n" +
                "                </th>\n" +
                "                <td style=\"width: 100px; text-align: left;\">\n" +
                "                    <button id=\"exportButton\">Export to Excel</button>\n" +
                "                </td>\n" +
                "            </tr>\n" +
                "        </table>\n" +
                "    </div>\n" +
                "    <div class=\"horizontal-bottom-panel-third\">\n" +
                "        <div id=\"piechart_3d\" style=\"width: 300px; height: 200px;\"></div>\n" +
                "        <script type=\"text/javascript\">\n" +
                "            google.charts.load('current', {'packages': ['corechart']});\n" +
                "            google.charts.setOnLoadCallback(drawChart);\n" +
                "            function drawChart() {\n" +
                "                var data = new google.visualization.DataTable();\n" +
                "                data.addColumn('string', 'Color');\n" +
                "                data.addColumn('number', 'Value');\n" +
                "                data.addRows([\n" +
                "                    ['Pass', Passed_Summary_Executions],\n" +
                "                    ['Fail', Failed_Summary_Executions]\n" +
                "                ]);\n" +
                "                var options = {\n" +
                "                    width: 300,\n" +
                "                    height: 200,\n" +
                "                    is3D: true, // Enable 3D effect\n" +
                "                    colors: ['green', 'red'], // Green for Pass, Red for Fail\n" +
                "                    legend: {\n" +
                "                        position: 'bottom' // Display legend below the chart\n" +
                "                    }\n" +
                "                };\n" +
                "                // Instantiate and draw the chart, passing in some options.\n" +
                "                var chart = new google.visualization.PieChart(document.getElementById('piechart_3d'));\n" +
                "                chart.draw(data, options);\n" +
                "            }\n" +
                "        </script>\n" +
                "    </div>\n" +
                "    <div class=\"horizontal-bottom-panel-fourth\">\n" +
                "        <div id=\"columnchart_2d\" style=\"width: 300px; height: 225px;\"></div>\n" +
                "        <script type=\"text/javascript\">\n" +
                "            google.charts.load(\"current\", { packages: ['corechart'] });\n" +
                "            google.charts.setOnLoadCallback(drawChart);\n" +
                "            function drawChart() {\n" +
                "                var data = google.visualization.arrayToDataTable([\n" +
                "                    [\"Element\", \"Count\", { role: \"style\" }, { role: \"annotation\" }],\n" +
                "                    [\"Pass\", Passed_Summary_Executions, \"green\", Passed_Summary_Executions],\n" +
                "                    [\"Fail\", Failed_Summary_Executions, \"red\", Failed_Summary_Executions]\n" +
                "                ]);\n" +
                "                var options = {\n" +
                "                    width: 300,\n" +
                "                    height: 200,\n" +
                "                    bar: { groupWidth: \"50%\" },\n" +
                "                    legend: { position: \"none\" },\n" +
                "                    tooltip: { isHtml: true }, // Enable HTML content in tooltips\n" +
                "                    hAxis: { title: 'Status' }, // Label for horizontal axis\n" +
                "                    vAxis: { title: 'Count' }, // Label for vertical axis\n" +
                "                    annotations: {\n" +
                "                        alwaysOutside: false, // Place annotations inside the bars\n" +
                "                        textStyle: {\n" +
                "                            fontSize: 12,\n" +
                "                            bold: true,\n" +
                "                            italic: false,\n" +
                "                            color: '#333'\n" +
                "                        }\n" +
                "                    }\n" +
                "                };\n" +
                "                var chart = new google.visualization.ColumnChart(document.getElementById(\"columnchart_2d\"));\n" +
                "                chart.draw(data, options);\n" +
                "                // After the chart renders, adjust the annotations position\n" +
                "                setTimeout(function () {\n" +
                "                    var chartContainer = document.getElementById('columnchart_2d');\n" +
                "                    var labels = chartContainer.getElementsByTagName('text');\n" +
                "                    for (var i = 0; i < labels.length; i++) {\n" +
                "                        if (labels[i].getAttribute('text-anchor') === 'middle') {\n" +
                "                            var y = parseFloat(labels[i].getAttribute('y'));\n" +
                "                            labels[i].setAttribute('y', y + 15); // Adjust Y position as needed\n" +
                "                        }\n" +
                "                    }\n" +
                "                }, 100); // Adjust the delay time as needed\n" +
                "            }\n" +
                "        </script>\n" +
                "    </div>\n" +
                "    <div class=\"horizontal-last-panel\">\n" +
                "        <!--<img src=\"https://picsum.photos/800/600\" alt=\"Lorem Picsum Placeholder\" width=\"200\" height=\"100\" style=\"margin-top: 20px;\">-->\n" +
                "        <h1><span style=\"text-transform: uppercase; font-size: 1.25em;\">&emsp;&emsp;T</span>EST<span style=\"text-transform: uppercase; font-size: 1.25em;\">&ensp;A</span>UTOMATION<span style=\"text-transform: uppercase; font-size: 1.25em;\">&ensp;S</span>UMMARY&ensp;&ensp;</h1>\n" +
                "    </div>\n" +
                "</div>\n" +
                "<div class=\"container\">\n" +
                "    <div class=\"left-panel\">\n" +
                "        <table id=\"Summary_Table\">\n" +
                "            <tr>\n" +
                "                <th style=\"width: 100px; text-align: center; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">W</span>ORKFLOWS\n" +
                "                </th>\n" +
                "                <th style=\"width: 100px; text-align: center; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">B</span>ROWSER\n" +
                "                </th>\n" +
                "                <th style=\"width: 100px; text-align: center; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">S</span>TART\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">T</span>IME\n" +
                "                </th>\n" +
                "                <th style=\"width: 100px; text-align: center; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">E</span>ND\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">T</span>IME\n" +
                "                </th>\n" +
                "                <th style=\"width: 100px; text-align: center; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">D</span>URATION\n" +
                "                </th>\n" +
                "                <th style=\"width: 100px; text-align: center; font-size: 10px;\">\n" +
                "                    <span style=\"text-transform: uppercase; font-size: 1.25em;\">S</span>TATUS\n" +
                "                </th>\n" +
                "            </tr>\n" +
                "            summarytablesreplacetext\n" +
                "        </table>\n" +
                "        <div class=\"pagination\" id=\"pagination\"></div>\n" +
                "    </div>\n" +
                "    <div class=\"right-panel\">\n" +
                "        <div class=\"tab\">\n" +
                "            <button class=\"tablinks\" onclick=\"openTab(event, 'testCases')\"><span style=\"text-transform: uppercase; font-size: 1.5em;\">T</span>EST \n" +
                "    <span style=\"text-transform: uppercase; font-size: 1.5em;\">C</span>ASES</button>\n" +
                "            <button class=\"tablinks\" onclick=\"openTab(event, 'logs')\"><span style=\"text-transform: uppercase; font-size: 1.5em;\">L</span>OGS</button>\n" +
                "        </div>\n" +
                "        <div id=\"testCases\" class=\"tabcontent\">\n" +
                "            testcasetablereplace\n" +
                "        </div>\n" +
                "        <div id=\"logs\" class=\"tabcontent\">\n" +
                "            logtablereplace\n" +
                "        </div>\n" +
                "    </div>\n" +
                "</div>\n" +
                "<script>\n" +
                "    var passCountsTestCases = 0;\n" +
                "    var failCountsTestCases = 0;\n" +
                "</script>\n" +
                "<script>\n" +
                "    function toggleTable(...ids) {\n" +
                "        var tables = document.querySelectorAll('.toggle-table');\n" +
                "        tables.forEach(function (table) {\n" +
                "            table.style.display = 'none';\n" +
                "        });\n" +
                "        ids.forEach(function (id) {\n" +
                "            var tableToShow = document.getElementById(id);\n" +
                "            if (tableToShow) {\n" +
                "                tableToShow.style.display = 'table';\n" +
                "            }\n" +
                "        });\n" +
                "    }\n" +
                "</script>\n" +
                "<script>\n" +
                "    // Function to initialize pagination for tables with IDs starting with a given prefix\n" +
                "    function initializePagination(tableIdPrefix, rowsPerPage, paginationClass) {\n" +
                "        var tables = document.querySelectorAll('[id^=\"' + tableIdPrefix + '\"]');\n" +
                "        tables.forEach(function (table) {\n" +
                "            var rows = table.rows.length - 1; // Exclude header row\n" +
                "            var pageCount = Math.ceil(rows / rowsPerPage);\n" +
                "            var pagination = document.createElement('div');\n" +
                "            pagination.classList.add(paginationClass);\n" +
                "            table.parentNode.insertBefore(pagination, table.nextSibling);\n" +
                "            createPaginationButtons();\n" +
                "            function createPaginationButtons() {\n" +
                "                var firstButton = createButton('First', function () {\n" +
                "                    showPage(1);\n" +
                "                });\n" +
                "                pagination.appendChild(firstButton);\n" +
                "                var prevButton = createButton('Previous', function () {\n" +
                "                    var currentPage = getCurrentPage();\n" +
                "                    if (currentPage > 1) {\n" +
                "                        showPage(currentPage - 1);\n" +
                "                    }\n" +
                "                });\n" +
                "                pagination.appendChild(prevButton);\n" +
                "                for (var i = 1; i <= pageCount; i++) {\n" +
                "                    var link = createButton(i, function () {\n" +
                "                        var pageNumber = parseInt(this.innerHTML);\n" +
                "                        showPage(pageNumber);\n" +
                "                    });\n" +
                "                    pagination.appendChild(link);\n" +
                "                }\n" +
                "                var nextButton = createButton('Next', function () {\n" +
                "                    var currentPage = getCurrentPage();\n" +
                "                    if (currentPage < pageCount) {\n" +
                "                        showPage(currentPage + 1);\n" +
                "                    }\n" +
                "                });\n" +
                "                pagination.appendChild(nextButton);\n" +
                "                var lastButton = createButton('Last', function () {\n" +
                "                    showPage(pageCount);\n" +
                "                });\n" +
                "                pagination.appendChild(lastButton);\n" +
                "            }\n" +
                "            function createButton(label, onclick) {\n" +
                "                var button = document.createElement('a');\n" +
                "                button.href = '#';\n" +
                "                button.innerHTML = label;\n" +
                "                button.addEventListener('click', onclick);\n" +
                "                return button;\n" +
                "            }\n" +
                "            function getCurrentPage() {\n" +
                "                var activeLink = pagination.querySelector('.active');\n" +
                "                return parseInt(activeLink.innerHTML);\n" +
                "            }\n" +
                "            function showPage(pageNumber) {\n" +
                "                var start = (pageNumber - 1) * rowsPerPage + 1;\n" +
                "                var end = Math.min(start + rowsPerPage - 1, rows);\n" +
                "\n" +
                "                for (var i = 1; i <= rows; i++) {\n" +
                "                    var row = table.rows[i];\n" +
                "                    if (i >= start && i <= end) {\n" +
                "                        row.style.display = 'table-row';\n" +
                "                    } else {\n" +
                "                        row.style.display = 'none';\n" +
                "                    }\n" +
                "                }\n" +
                "                updateActiveLink(pageNumber);\n" +
                "            }\n" +
                "            function updateActiveLink(pageNumber) {\n" +
                "                var links = pagination.getElementsByTagName('a');\n" +
                "                for (var i = 0; i < links.length; i++) {\n" +
                "                    links[i].classList.remove('active');\n" +
                "                    if (parseInt(links[i].innerHTML) === pageNumber) {\n" +
                "                        links[i].classList.add('active');\n" +
                "                    }\n" +
                "                }\n" +
                "            }\n" +
                "            showPage(1); // Show first page initially\n" +
                "        });\n" +
                "    }\n" +
                "    // Initialize pagination for tables starting with 'Table_For_TestCases_'\n" +
                "    //initializePagination('Table_For_TestCases_', 10, 'pagination');\n" +
                "    // Initialize pagination for tables starting with 'Table_For_TestCases_'\n" +
                "    //initializePagination('Table_For_Logs_', 10, 'pagination');\n" +
                "    // Initialize pagination for tables starting with 'Table_For_TestCases_'\n" +
                "    initializePagination('Summary_Table', 18, 'pagination');\n" +
                "    // You can initialize pagination for other table prefixes or adjust rows per page as needed\n" +
                "</script>\n" +
                "<script>\n" +
                "    function toggleTable(testCasesTableId, logsTableId) {\n" +
                "        var tables = document.querySelectorAll('.toggle-table');\n" +
                "        tables.forEach(function (table) {\n" +
                "            table.style.display = 'none';\n" +
                "        });\n" +
                "        // Show the specified tables\n" +
                "        document.getElementById(testCasesTableId).style.display = 'table';\n" +
                "        document.getElementById(logsTableId).style.display = 'table';\n" +
                "        // Call countTextInColumns for test cases table\n" +
                "        passCountsTestCases = countTextInColumns(testCasesTableId, \"Pass\", \"passCountsTestCases\");\n" +
                "        failCountsTestCases = countTextInColumns(testCasesTableId, \"Fail\", \"failCountsTestCases\");\n" +
                "        // Call countTextInColumns for logs table\n" +
                "        var passCountsLogs = countTextInColumns(logsTableId, \"Pass\", \"passCountsLogs\");\n" +
                "        var failCountsLogs = countTextInColumns(logsTableId, \"Fail\", \"failCountsLogs\");\n" +
                "    }\n" +
                "</script>\n" +
                "<script>\n" +
                "    function countTextInColumns(tableName, searchText, targetDivId) {\n" +
                "        var table = document.getElementById(tableName);\n" +
                "        var counts = [0, 0, 0]; // Initialize counts for each column\n" +
                "        // Loop through each row of the table\n" +
                "        for (var i = 0; i < table.rows.length; i++) {\n" +
                "            // Loop through each cell of the row\n" +
                "            for (var j = 0; j < table.rows[i].cells.length; j++) {\n" +
                "                // Check if the cell text matches the search text\n" +
                "                if (table.rows[i].cells[j].innerText.trim() === searchText) {\n" +
                "                    // Increment count for the current column\n" +
                "                    counts[j]++;\n" +
                "                }\n" +
                "            }\n" +
                "        }\n" +
                "// Print counts to the HTML page\n" +
                "        var targetDiv = document.getElementById(targetDivId);\n" +
                "        targetDiv.innerHTML = \"\"; // Clear previous results\n" +
                "        var heading = document.createElement(\"h1\");\n" +
                "        heading.textContent = searchText + \"ed Test Cases\";\n" +
                "        targetDiv.appendChild(heading);\n" +
                "        for (var i = 2; i < counts.length; i++) {\n" +
                "            var heading = document.createElement(\"h2\");\n" +
                "            heading.style.textAlign = \"center\";\n" +
                "            heading.textContent = counts[i];\n" +
                "            targetDiv.appendChild(heading);\n" +
                "        }\n" +
                "        return counts;\n" +
                "    }\n" +
                "</script>\n" +
                "<script>\n" +
                "    function openTab(evt, tabName) {\n" +
                "        var i, tabcontent, tablinks;\n" +
                "        tabcontent = document.getElementsByClassName(\"tabcontent\");\n" +
                "        for (i = 0; i < tabcontent.length; i++) {\n" +
                "            tabcontent[i].style.display = \"none\";\n" +
                "        }\n" +
                "        tablinks = document.getElementsByClassName(\"tablinks\");\n" +
                "        for (i = 0; i < tablinks.length; i++) {\n" +
                "            tablinks[i].className = tablinks[i].className.replace(\" active\", \"\");\n" +
                "        }\n" +
                "        document.getElementById(tabName).style.display = \"block\";\n" +
                "        evt.currentTarget.className += \" active\";\n" +
                "    }\n" +
                "</script>\n" +
                "<script>\n" +
                "    function displayImage(base64) {\n" +
                "        var width = 1000;\n" +
                "        var height = 500;\n" +
                "        var left = (window.innerWidth - width) / 2;\n" +
                "        var top = (window.innerHeight - height) / 2;\n" +
                "        var newWindow = window.open(\"\", \"_blank\", \"width=\" + width + \",height=\" + height + \",left=\" + left + \",top=\" + top + \",toolbar=no,location=no,menubar=no,status=no,titlebar=no\");\n" +
                "        newWindow.document.write(\"<style>body {margin: 0;}</style><img src='\" + base64 + \"' style='max-width: 100%; max-height: 100%;'>\");\n" +
                "    }\n" +
                "</script>\n" +
                "<script>\n" +
                "    document.getElementById(\"exportButton\").addEventListener(\"click\", function () {\n" +
                "        var wb = XLSX.utils.book_new();\n" +
                "        exportTableToExcel(\"Summary_Table\", \"Summary\", wb);\n" +
                "        exportTablesWithPrefixToExcel(\"Table_For_Logs_\", \"Logs\", wb);\n" +
                "        exportTablesWithPrefixToExcel(\"Table_For_TestCases_\", \"TCs\", wb);\n" +
                "        XLSX.writeFile(wb, 'Export_Execution_Report.xlsx');\n" +
                "    });\n" +
                "    function exportTableToExcel(tableId, sheetName, workbook) {\n" +
                "        var table = document.getElementById(tableId);\n" +
                "        var wsData = [];\n" +
                "        // Get all rows of the table\n" +
                "        var rows = Array.from(table.querySelectorAll('tr'));\n" +
                "        // Iterate through each row\n" +
                "        rows.forEach(function (row) {\n" +
                "            var rowData = [];\n" +
                "            // Iterate through each cell (th or td) in the row\n" +
                "            row.querySelectorAll('th, td').forEach(function (cell) {\n" +
                "                // Modify here to get plain text content\n" +
                "                var cellText = cell.textContent.trim(); // Get text content and trim any leading/trailing whitespace\n" +
                "                rowData.push(cellText);\n" +
                "            });\n" +
                "            wsData.push(rowData);\n" +
                "        });\n" +
                "        // Convert data array to worksheet\n" +
                "        var ws = XLSX.utils.aoa_to_sheet(wsData);\n" +
                "        // Set styles for each cell to achieve left alignment and vertical center alignment\n" +
                "        Object.keys(ws).forEach(function (cellRef) {\n" +
                "            var cell = ws[cellRef];\n" +
                "            if (cell && cell.t === 's' && cell.v) {\n" +
                "                cell.s = {\n" +
                "                    alignment: {horizontal: 'left', vertical: 'center'},\n" +
                "                    font: {sz: 12, bold: false, underline: false, italic: false}\n" +
                "                };\n" +
                "            }\n" +
                "        });\n" +
                "        // Append the worksheet to the workbook with the specified sheet name\n" +
                "        XLSX.utils.book_append_sheet(workbook, ws, sheetName);\n" +
                "    }\n" +
                "    function exportTablesWithPrefixToExcel(tablePrefix, sheetNamePrefix, workbook) {\n" +
                "        var tables = document.querySelectorAll('[id^=\"' + tablePrefix + '\"]');\n" +
                "        tables.forEach(function (table, index) {\n" +
                "            var constVariable = getConstantVariable(table.id);\n" +
                "            var sheetName = sheetNamePrefix + \"_\" + constVariable;\n" +
                "            exportTableToExcel(table.id, sheetName, workbook); // Pass table id to exportTableToExcel\n" +
                "        });\n" +
                "    }\n" +
                "</script>\n" +
                "<script>\n" +
                "    const Constants = {\n" +
                "        createtestcasesconstantVariables\n" +
                "\n" +
                "        createlogssconstantVariables\n" +
                "    }\n" +
                "function getConstantVariable(tableName) {\n" +
                "        return Constants[tableName] || \"\"; // Returns empty string if tableName is not found in Constants\n" +
                "    }\n" +
                "</script>\n" +
                "</body>\n" +
                "</html>\n";
        String filePathAndName = System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + fileName;
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePathAndName))) {
            writer.write(htmlContent);
            //System.out.println("HTML file created successfully at: " + filePathAndName);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static String getBrowserNameForExportFile(String browserName) {
        BrowserType browserType;
        switch (browserName.toUpperCase().trim()) {
            case "CHROME":
                browserName="C";
                break;
            case "EDGE":
                browserName="E";
                break;
            case "CHROME_HEADLESS":
                browserName="CHL";
                break;
            case "EDGE_HEADLESS":
                browserName="EHL";
                break;
            case "FIREFOX":
                browserName="F";
                break;
            case "FIREFOX_HEADLESS":
                browserName="FHL";
                break;
        }
        return browserName;
    }



    //writeSummaryInfoIntoHtmlReport
    public static void writeSummaryInfoIntoHtmlReport(String excelFileName, String htmlFileName) throws IOException {
        FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + excelFileName);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet xs = wb.getSheet("Summary_Info");
        int passedTests = 0;
        int failedTests = 0;
        int rowNumber = xs.getLastRowNum();
        System.out.println(rowNumber);
        String valueForTable = "";
        String valueTCsConstVariable = "";
        String valueLogsConstVariable = "";
        for (int methodsCount = 1; methodsCount <= rowNumber; methodsCount++) {
            valueForTable = valueForTable + "<tr>";
            // Row row = xs.getRow(methodsCount);
            Row row = xs.getRow(methodsCount);
            for (int cell = 1; cell < row.getLastCellNum(); cell++) {
                Cell cellValue = row.getCell(cell);
                if (cellValue.toString().equals("PASSED")) {
                    passedTests++;
                } else if (cellValue.toString().equals("FAILED")) {
                    failedTests++;
                }
                if (cell == 1) {
                    valueForTable = valueForTable + "\n<td style=\"width: 40%; text-align: left; overflow-wrap: anywhere;\"class=\"hint\" data-hint=\""+row.getCell(cell+1).toString()+"\">\n<a href=\"#\" onclick=\"(function(){ toggleTable('" + "Table_For_TestCases_" + row.getCell(0) + "','" + "Table_For_Logs_" + row.getCell(0) + "'); countPassTextInColumns('Table_For_TestCases_" + row.getCell(0) + "'); return false; })();\">" + cellValue + "</a>\n</td>";
                    valueTCsConstVariable = valueTCsConstVariable + "Table_For_TestCases_" + row.getCell(0) + ":" + "\"" + cellValue;
                    valueLogsConstVariable = valueLogsConstVariable + "Table_For_Logs_" + row.getCell(0) + ":" + "\"" + cellValue;

                } else if (cell == 3) {
                    valueForTable = valueForTable + "\n<td style=\"width: 10%; text-align: center;\">\n<a href=\"#\" onclick=\"(function(){ toggleTable('" + "Table_For_TestCases_" + row.getCell(0) + "','" + "Table_For_Logs_" + row.getCell(0) + "'); countPassTextInColumns('Table_For_TestCases_" + row.getCell(0) + "'); return false; })();\">" + cellValue + "</a>\n</td>";
                    valueTCsConstVariable = valueTCsConstVariable + "_" + getBrowserNameForExportFile(String.valueOf(cellValue)) + "\",\n";
                    valueLogsConstVariable = valueLogsConstVariable + "_" + getBrowserNameForExportFile(String.valueOf(cellValue)) + "\",\n";

                } else if (cell == 4) {
                    valueForTable = valueForTable + "\n<td style=\"width: 15%; text-align: center;\">\n<a href=\"#\" onclick=\"(function(){ toggleTable('" + "Table_For_TestCases_" + row.getCell(0) + "','" + "Table_For_Logs_" + row.getCell(0) + "'); countPassTextInColumns('Table_For_TestCases_" + row.getCell(0) + "'); return false; })();\">" + cellValue + "</a>\n</td>";
                } else if (cell == 5) {
                    valueForTable = valueForTable + "\n<td style=\"width: 15%; text-align: center;\">\n<a href=\"#\" onclick=\"(function(){ toggleTable('" + "Table_For_TestCases_" + row.getCell(0) + "','" + "Table_For_Logs_" + row.getCell(0) + "'); countPassTextInColumns('Table_For_TestCases_" + row.getCell(0) + "'); return false; })();\">" + cellValue + "</a>\n</td>";
                } else if (cell == 6) {
                    valueForTable = valueForTable + "\n<td style=\"width: 10%; text-align: center;\">\n<a href=\"#\" onclick=\"(function(){ toggleTable('" + "Table_For_TestCases_" + row.getCell(0) + "','" + "Table_For_Logs_" + row.getCell(0) + "'); countPassTextInColumns('Table_For_TestCases_" + row.getCell(0) + "'); return false; })();\">" + cellValue + "</a>\n</td>";
                } else if (cell == 7) {
                    valueForTable = valueForTable + "\n<td style=\"width: 10%; text-align: center;\">\n<a href=\"#\" onclick=\"(function(){ toggleTable('" + "Table_For_TestCases_" + row.getCell(0) + "','" + "Table_For_Logs_" + row.getCell(0) + "'); countPassTextInColumns('Table_For_TestCases_" + row.getCell(0) + "'); return false; })();\">" + cellValue + "</a>\n</td>";
                }
            }
            valueForTable = valueForTable + "\n</tr>";
        }
        replaceTextInHTMLReport(htmlFileName, "summarytablesreplacetext", valueForTable + "summarytablesreplacetext");
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
        String current_Date = dateFormat.format(Calendar.getInstance().getTime());
        replaceTextInHTMLReport(htmlFileName, "Execution_Date", current_Date);
        replaceTextInHTMLReport(htmlFileName, "Executed_By", System.getProperty("user.name"));
        replaceTextInHTMLReport(htmlFileName, "Total_Summary_Executions", String.valueOf(passedTests + failedTests));
        replaceTextInHTMLReport(htmlFileName, "Passed_Summary_Executions", String.valueOf(passedTests));
        replaceTextInHTMLReport(htmlFileName, "Failed_Summary_Executions", String.valueOf(failedTests));
        replaceTextInHTMLReport(htmlFileName, "createtestcasesconstantVariables", valueTCsConstVariable + "createtestcasesconstantVariables");
        replaceTextInHTMLReport(htmlFileName, "createlogssconstantVariables", valueLogsConstVariable + "createlogssconstantVariables");
    }

    //writeTestCasesInfoIntoHtmlReport
    public static void writeTestCasesInfoIntoHtmlReport(String excelFileName, String htmlFileName) throws IOException {
        FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + excelFileName);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet xs = wb.getSheet("TestCases_Info");
        int passedTestCases = 0;
        int failedTestCases = 0;
        // Create a Set to store unique values
        Set<String> uniqueValues = new HashSet<>();
        for (int testCases = 1; testCases <= xs.getLastRowNum(); testCases++) {
            // Row row = xs.getRow(methodsCount);
            Row row = xs.getRow(testCases);
            Cell cellValue = row.getCell(0);
            uniqueValues.add(cellValue.toString());
        }
        String valueForTable = null;
        for (String value : uniqueValues) {
            valueForTable = "<table border=\"1\" id=\"Table_For_TestCases_" + value + "\" class=\"toggle-table\">\n<tr>\n<th><span style=\"text-transform: uppercase; font-size: 1.25em;\">T</span>EST <span style=\"text-transform: uppercase; font-size: 1.25em;\">C</span>ASES</th>\n<th><span style=\"text-transform: uppercase; font-size: 1.25em;\">I</span>MAGES</th>\n<th><span style=\"text-transform: uppercase; font-size: 1.25em;\">S</span>TATUS</th></tr>\n";
            for (int testCases = 1; testCases <= xs.getLastRowNum(); testCases++) {
                Row row = xs.getRow(testCases);
                if (value.equals(row.getCell(0).toString())) {
                    valueForTable = valueForTable + "<tr>\n";
                    for (int testCaseCol = 1; testCaseCol < row.getLastCellNum(); testCaseCol++) {

                        Cell cellValue = row.getCell(testCaseCol);
                        if (testCaseCol == 1) {
                            valueForTable = valueForTable + "<td style=\"width: 70%; text-align: left; overflow-wrap: anywhere;\">" + cellValue.toString() + "</td>\n";
                        } else if (testCaseCol == 2) {
                            if (cellValue.toString().isEmpty()) {
                                valueForTable = valueForTable + "<td>" + "" + "</td>\n";
                            } else {
                                // Read image file
                                File file = new File(cellValue.toString());
                                FileInputStream imageFile = new FileInputStream(file);
                                byte[] imageData = new byte[(int) file.length()];
                                imageFile.read(imageData);
                                imageFile.close();
                                // Convert image to Base64
                                String base64Image = Base64.getEncoder().encodeToString(imageData);
                                valueForTable = valueForTable + "<td><a href=\"#\" onclick=\"displayImage('data:image/png;base64," + base64Image + "')\">Image</a></td>";
                            }
                        } else {
                            valueForTable = valueForTable + "<td>" + cellValue.toString() + "</td>\n";
                        }
                    }
                    valueForTable = valueForTable + "</tr>\n";
                }
            }
            valueForTable = valueForTable + "</table>\n";
            replaceTextInHTMLReport(htmlFileName, "testcasetablereplace", valueForTable + "testcasetablereplace");
        }
    }

    //writeLogsInfoIntoHtmlReport
    public static void writeLogsInfoIntoHtmlReport(String excelFileName, String htmlFileName) throws IOException {
        FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + excelFileName);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet xs = wb.getSheet("Logs_Info");
        // Create a Set to store unique values
        Set<String> uniqueValues = new HashSet<>();
        for (int logs = 1; logs <= xs.getLastRowNum(); logs++) {
            // Row row = xs.getRow(methodsCount);
            Row row = xs.getRow(logs);
            Cell cellValue = row.getCell(0);
            uniqueValues.add(cellValue.toString());
        }
        String valueForTable = null;
        for (String value : uniqueValues) {
            valueForTable = "<table border=\"1\" id=\"Table_For_Logs_" + value + "\" class=\"toggle-table\"><tr>\n<th style=\"width: 15%; text-align: center;\"><span style=\"text-transform: uppercase; font-size: 1.25em;\">D</span>ATE</th>\n<th style=\"width: 10%; text-align: center;\"><span style=\"text-transform: uppercase; font-size: 1.25em;\">T</span>IME</th>\n<th style=\"width: 15%; text-align: center;\"><span style=\"text-transform: uppercase; font-size: 1.25em;\">L</span>OGTYPE</th>\n<th style=\"width: 60%; text-align: center;\"><span style=\"text-transform: uppercase; font-size: 1.25em;\">D</span>ESCRIPTION</th>\n</tr>\n";
            for (int logs = 1; logs <= xs.getLastRowNum(); logs++) {
                Row row = xs.getRow(logs);
                if (value.equals(row.getCell(0).toString())) {
                    valueForTable = valueForTable + "<tr>\n";
                    for (int logs_Cols = 1; logs_Cols < row.getLastCellNum(); logs_Cols++) {
                        Cell cellValue = row.getCell(logs_Cols);
                        if (logs_Cols == 4) {
                            valueForTable = valueForTable + "<td style=\"width: 60%; text-align: left; overflow-wrap: anywhere;\">" + cellValue.toString() + "</td>";
                        } else if (logs_Cols == 3) {
                            valueForTable = valueForTable + "<td style=\"width: 15%; text-align: center;\">" + cellValue.toString() + "</td>";
                        } else if (logs_Cols == 2) {
                            valueForTable = valueForTable + "<td style=\"width: 10%; text-align: center;\">" + cellValue.toString() + "</td>";
                        } else if (logs_Cols == 1) {
                            valueForTable = valueForTable + "<td style=\"width: 15%; text-align: center;\">" + cellValue.toString() + "</td>";
                        }
                    }
                    valueForTable = valueForTable + "\n</tr>\n";
                }
            }
            valueForTable = valueForTable + "</table>\n";
            replaceTextInHTMLReport(htmlFileName, "logtablereplace", valueForTable + "logtablereplace");
        }
    }

    //replaceTextInHTMLReport
    public static void replaceTextInHTMLReport(String htmlFileName, String sourceText, String destinationText) throws IOException {
        try {
            // Read the content of the HTML file
            BufferedReader reader = new BufferedReader(new FileReader(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + htmlFileName));
            StringBuilder stringBuilder = new StringBuilder();
            String line;
            while ((line = reader.readLine()) != null) {
                stringBuilder.append(line).append("\n");
            }
            String fileContent = stringBuilder.toString();
            reader.close();
            // Perform text replacement
            String modifiedContent = fileContent.replace(sourceText, destinationText);
            // Write the modified content back to the HTML file
            BufferedWriter writer = new BufferedWriter(new FileWriter(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + htmlFileName));
            writer.write(modifiedContent);
            writer.close();
            //System.out.println("Text replaced successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static synchronized void writeSummaryInfoIntoExcel(String logFileName, String sheetName, String[] logAndSummaryInfo) throws IOException, InterruptedException {
        ZipSecureFile.setMinInflateRatio(0);
        File file = new File(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + logFileName);
        Thread.sleep(2000);
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(sheetName);
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(++lastRowNum);
        for (int i = 0; i < logAndSummaryInfo.length; i++) {
            Cell cell_threadID = row.createCell(i);
            cell_threadID.setCellValue(logAndSummaryInfo[i]);
        }
        // Write the modified workbook content back to the same file
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
            //System.out.println("Data added to the existing Excel file.");
        }
    }

    public static synchronized void writeLogsInfoIntoExcel(String logFileName, String sheetName, String threadID, String logInfo, String logDescription) throws IOException {
        ZipSecureFile.setMinInflateRatio(0);
        File file = new File(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + logFileName);
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(sheetName);
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(++lastRowNum);
        String[] arrLogValues = {threadID, String.valueOf(LocalDate.now()), String.valueOf(LocalTime.now().format(DateTimeFormatter.ofPattern("HH:mm:ss"))), logInfo, logDescription};
        for (int i = 0; i < arrLogValues.length; i++) {
            Cell cell_threadID = row.createCell(i);
            cell_threadID.setCellValue(arrLogValues[i]);
        }
        // Write the modified workbook content back to the same file
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
            //System.out.println("Data added to the existing Excel file.");
        }
    }

    //writeTestCasesInfoIntoExcel
    public static synchronized void writeTestCasesInfoIntoExcel(String excelFileName, String sheetName, String threadID, String testCase, String imageString, String testCaseStatus) throws IOException {
        ZipSecureFile.setMinInflateRatio(0);
        File file = new File(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + excelFileName);
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(sheetName);
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(++lastRowNum);
        String[] arrLogValues = {threadID, testCase, imageString, testCaseStatus};
        for (int i = 0; i < arrLogValues.length; i++) {
            Cell cell_threadID = row.createCell(i);
            cell_threadID.setCellValue(arrLogValues[i]);
        }
        // Write the modified workbook content back to the same file
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
            //System.out.println("Data added to the existing Excel file.");
        }
    }

    public static void writeDataIntoExcelCell(String logFileName) throws IOException {
        // ZipSecureFile.setMinInflateRatio(0);
        File file = new File(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + logFileName);
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(++lastRowNum);
        Cell cell = row.createCell(0);
        cell.setCellValue("test");
        // Write the modified workbook content back to the same file
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
            //System.out.println("Data added to the existing Excel file.");
        }
    }

    public static void createExcelLogFile(String logFileName) throws IOException {
        try {
            File f = new File(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + logFileName);
            if (f.exists() && !f.isDirectory()) {
                System.out.println("Excel file already exists");
            } else {
                FileOutputStream fos = new FileOutputStream(System.getProperty("user.dir") + "\\src\\test\\java\\Reports\\" + logFileName);
                XSSFWorkbook workbook = new XSSFWorkbook();
                //Need to identify proper newly created excel sheet.
                XSSFSheet sheet = workbook.createSheet("Summary_Info");
                Row row = sheet.createRow(0);
                String arr[] = {"Thread_ID", "Workflow_Code","Workflow_Description", "Browser", "Start_Time", "End_Time", "Duration", "Status"};
                //HashMap<String, String> logData = new HashMap();
                sheet.setZoom(100);
                for (int cell = 0; cell < arr.length; cell++) {
                    Cell cellValue = row.createCell(cell);
                    cellValue.setCellValue(arr[cell]);
                    sheet.setColumnWidth(cell, 5000);
                }
                XSSFSheet sheetForTestCases = workbook.createSheet("TestCases_Info");
                Row rowValueForTestCases = sheetForTestCases.createRow(0);
                String[] arrForTestCasesHeaders = {"Thread_ID", "Test_Case", "Image", "Status"};
                sheet.setZoom(100);
                for (int cell = 0; cell < arrForTestCasesHeaders.length; cell++) {
                    Cell cellValue = rowValueForTestCases.createCell(cell);
                    cellValue.setCellValue(arrForTestCasesHeaders[cell]);
                    sheetForTestCases.setColumnWidth(cell, 5000);
                }
                XSSFSheet sheetForLogs = workbook.createSheet("Logs_Info");
                Row rowValueForLogs = sheetForLogs.createRow(0);
                String arrForLogHeaders[] = {"Thread_ID", "Date", "Time", "Log_Type", "Description"};
                sheet.setZoom(100);
                for (int cell = 0; cell < arrForLogHeaders.length; cell++) {
                    Cell cellValue = rowValueForLogs.createCell(cell);
                    cellValue.setCellValue(arrForLogHeaders[cell]);
                    sheetForLogs.setColumnWidth(cell, 5000);
                }
                workbook.write(fos);
                fos.flush();
                fos.close();
            }
        } catch (IOException e) {
            System.out.println(e);
            e.printStackTrace();
        }
    }

    static String globalPropertiesFileNameAndPath = System.getProperty("user.dir") + "\\src\\main\\java\\GlobalProperties\\global.properties";
    static String HTMLLogInfoPropertiesFileNameAndPath = System.getProperty("user.dir") + "\\src\\main\\java\\GlobalProperties\\HTML_Log_Info.properties";

    public static boolean isColumnAvaiInExcel(String dataSheetFile, String dataSheetName, int rowNumber, String columnHeader) throws IOException, InvalidFormatException {
        File file = new File(dataSheetFile);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet(dataSheetName);
        Row row = sheet.getRow(rowNumber);
        boolean bool = false;
        for (int column = 0; column < row.getLastCellNum(); column++) {
            Cell celll = row.getCell(column);
            String columnVal = celll.toString();
            if (columnVal.equalsIgnoreCase(columnHeader)) {
                bool = true;
                break;
            } else {
                bool = false;
            }
        }
        workbook.close();
        return bool;
    }

    public static int findColumnPositionInExcel(String dataSheetFile, String dataSheetName, int rowNumber, String columnHeader) throws IOException, InvalidFormatException {
        File file = new File(dataSheetFile);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet(dataSheetName);
        Row row = sheet.getRow(rowNumber);
        int columnPosition = -1;
        try {
            for (int column = 0; column < row.getLastCellNum(); column++) {
                Cell cell = row.getCell(column);
                String columnVal = cell.toString();
                if (columnVal.equalsIgnoreCase(columnHeader)) {
                    columnPosition = column;
                    break;
                }
            }
        } catch (Exception e) {
            System.out.println(e);
        }
        workbook.close();
        return columnPosition;
    }

    public static void writeDataIntoExcelCell(int colValue, int rowValue, String sheetName) throws IOException {
        // ZipSecureFile.setMinInflateRatio(0);
        FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "\\src\\main\\java\\Test_Data\\TestData.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet xs = wb.getSheet(sheetName);
        FileOutputStream fos = null;
        XSSFCell columnNumber = null;
        XSSFRow rowNumber = null;
        int colNum = colValue;
        rowNumber = xs.getRow(rowValue);
        if (rowNumber == null) {
            rowNumber = xs.createRow(rowValue);
        }
        columnNumber = rowNumber.getCell(colNum);
        if (columnNumber == null) {
            columnNumber = rowNumber.createCell(colNum);
        }
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy");
        String current_Date = dateFormat.format(Calendar.getInstance().getTime());
        SimpleDateFormat timeFormat = new SimpleDateFormat("HH:mm:ss");
        TimeZone etTimeZone = TimeZone.getTimeZone("Asia/Kolkata");
        timeFormat.setTimeZone(etTimeZone);
        String current_Time = timeFormat.format(Calendar.getInstance().getTime());
        columnNumber.setCellValue(current_Time);
        //System.out.println(current_Time);
        fos = new FileOutputStream(System.getProperty("user.dir") + "\\src\\main\\java\\Test_Data\\TestData.xlsx");
        wb.write(fos);
        wb.close();
        fos.close();
    }

    public static String[] readExecutionMethodsFromExcel(String filePath, String fileName, String sheetName) throws IOException {
        FileInputStream file = new FileInputStream(System.getProperty("user.dir") + "\\" + filePath + "\\" + fileName);
        XSSFWorkbook xwb = new XSSFWorkbook(file);
        //Need to change below sheetName to executedScriptName
        String[] arr_ExecutionMethods = new String[0];
        try {
            XSSFSheet xs = xwb.getSheet(sheetName);
            arr_ExecutionMethods = new String[xs.getLastRowNum() + 1];
            for (int i = 0; i <= xs.getLastRowNum(); i++) {
                Row row = xs.getRow(i);
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    arr_ExecutionMethods[i] = String.valueOf(row.getCell(j));
                }
            }
        } catch (Exception e) {
            System.out.println(e);
            writeLog("Error", "Execution sheet not found in Execution Data File");
            writeLog("Error", String.valueOf(e));
            arr_ExecutionMethods = null;
        }
        return arr_ExecutionMethods;
    }

    public static String executionTimeDiffCalculation(String StartTime, String EndTime) throws IOException, NoSuchMethodException, NoSuchFieldException, IllegalAccessException, ClassNotFoundException, InstantiationException, ParseException {
        // Parse start and end time strings into Date objects
        SimpleDateFormat format = new SimpleDateFormat("HH:mm:ss");
        Date startTime = format.parse(StartTime);
        Date endTime = format.parse(EndTime);
        // Calculate time difference in milliseconds
        long timeDifference = endTime.getTime() - startTime.getTime();
        // Format time difference as HH:mm:ss
        long hours = TimeUnit.MILLISECONDS.toHours(timeDifference);
        long minutes = TimeUnit.MILLISECONDS.toMinutes(timeDifference) % 60;
        long seconds = TimeUnit.MILLISECONDS.toSeconds(timeDifference) % 60;
        return String.format("%02d:%02d:%02d", hours, minutes, seconds);
    }

    public static String TimeDiffCalculation(Date StartTime, Date EndTime) throws IOException, NoSuchMethodException, NoSuchFieldException, IllegalAccessException, ClassNotFoundException, InstantiationException {
        long timeDifference = EndTime.getTime() - StartTime.getTime();
        //System.out.println(StartTime);
        //System.out.println(EndTime);

        String TimeTaken = String.format("%s:%s:%s", Long.toString(TimeUnit.MILLISECONDS.toHours(Math.round(timeDifference))), TimeUnit.MILLISECONDS.toMinutes(Math.round(timeDifference)), TimeUnit.MILLISECONDS.toSeconds(Math.round(timeDifference)));
        //System.out.println(String.format("Time taken %s", TimeTaken));
        return TimeTaken;
    }

    public static boolean isSheetAvailableInExcel(String fileName, String sheetName) throws IOException, InvalidFormatException {
        File file = new File(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        boolean sheetAvaliableStatus = false;
        if (workbook.getNumberOfSheets() != 0) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (workbook.getSheetName(i).equals(sheetName)) {
                    sheetAvaliableStatus = true;
                    break;
                } else {
                    sheetAvaliableStatus = false;
                }
            }
        }
        workbook.close();
        return sheetAvaliableStatus;
    }

    public static void copyFileUsingApacheCommonsIO(File source, File dest) throws IOException {
        FileUtils.copyFile(source, dest);
    }

    public static void createCustomDataAndTimeStampForFileNames(String dataTimeStamp) throws IOException {
        WriteDataIntoPropertiesFile(globalPropertiesFileNameAndPath, "customDataAndTimeStampForFileNames", String.valueOf(dataTimeStamp));

        WriteDataIntoPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath, "Scripts_Execution_Start_Time", getCustomTime());
    }

    public static void createReportFile() throws IOException {
        String logFileName = "Report_" + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("customDataAndTimeStampForFileNames");
        WriteDataIntoPropertiesFile(globalPropertiesFileNameAndPath, "executionReportFileName", logFileName);
    }

    public static void replaceTextInHTML(String text, String replacement) throws IOException {

        Path path = Paths.get(System.getProperty("user.dir") + "\\src\\main\\java\\Reports\\" + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("executionReportFileName") + ".html");
        // Get all the lines
        try (Stream<String> stream = Files.lines(path, StandardCharsets.UTF_8)) {
            // Do the replace operation
            List<String> list = stream.map(line -> line.replace(text, replacement)).collect(Collectors.toList());
            // Write the content back
            Files.write(path, list, StandardCharsets.UTF_8);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String customDateFormat() {
        DateFormat dateFormat = new SimpleDateFormat("dd_MMM_YYYY_HH_mm_ss");
        Date date = new Date();
        dateFormat.format(date);
        String dateandtime = dateFormat.format(date).toString();
        return dateandtime;
    }

    public static String getCustomTime() {
        SimpleDateFormat timeFormat = new SimpleDateFormat("HH:mm:ss");
        TimeZone etTimeZone = TimeZone.getTimeZone("Asia/Kolkata");
        timeFormat.setTimeZone(etTimeZone);
        String current_Time = timeFormat.format(Calendar.getInstance().getTime());
        return current_Time;
    }

    public static String getCustomDate() {
        DateFormat dateFormat = new SimpleDateFormat("dd-MMM-YYYY");
        Date date = new Date();
        dateFormat.format(date);
        String dateandtime = dateFormat.format(date).toString();
        return dateandtime;
    }

    public static Properties ReadDataFromPropertiesFile(String propertiesFileName) throws IOException {
        FileInputStream readPropertyFile = new FileInputStream(propertiesFileName);
        Properties pro = new Properties();
        pro.load(readPropertyFile);
        return pro;
    }


    //need to update below pageload method with default time
    public void pageLoadReadyState(WebDriver driver, String PageName) throws InterruptedException, IOException {
        String expectedPageTitle = "Test";
        Duration.ofSeconds(5);
        do {
            if (driver.getTitle().equalsIgnoreCase(PageName)) {
                break;
            }
        } while (!driver.getTitle().equalsIgnoreCase(expectedPageTitle));
    }

    public static void elementAvailabilityState(WebDriver driver, WebElement element, String elementName) throws InterruptedException {
        WebDriverWait wdWait = new WebDriverWait(driver, Duration.ofSeconds(5));
        wdWait.until(ExpectedConditions.presenceOfElementLocated(By.linkText(elementName)));
    }

    public void stableZoomSizeBrowserWindow() throws InterruptedException, AWTException {
        Thread.sleep(1000);
        Robot robot = new Robot();
        robot.keyPress(KeyEvent.VK_CONTROL);
        robot.keyPress(KeyEvent.VK_0);
        robot.keyRelease(KeyEvent.VK_CONTROL);
        robot.keyRelease(KeyEvent.VK_0);
    }

    public void zoomOutBrowserWindowWithRobotClass() throws InterruptedException, AWTException {
        Thread.sleep(1000);
        Robot robot = new Robot();
        for (int i = 0; i < 3; i++) {
            robot.keyPress(KeyEvent.VK_CONTROL);
            robot.keyPress(KeyEvent.VK_SUBTRACT);
            robot.keyRelease(KeyEvent.VK_SUBTRACT);
            robot.keyRelease(KeyEvent.VK_CONTROL);
        }
    }

    public static void WriteDataIntoPropertiesFile(String requiredPropertiesFileName, String propertiesKey, String propertiesValue) throws IOException {
        Properties properties = new Properties();
        FileInputStream fis = new FileInputStream(requiredPropertiesFileName);
        properties.load(fis);
        properties.setProperty(propertiesKey, propertiesValue);
        FileOutputStream fos = new FileOutputStream(requiredPropertiesFileName);
        properties.store(fos, "LogFileFormat ==> Options ==> TextFile,Excel" + "\n" + "LogDataAppendedToExistingFileorNot ==> Options ==> Yes, No" + "\n" + "If LogFileFormate is null then by default TextLogs will be created.");
    }

    public static void writeLog(String Category, String LogString) throws IOException {
        boolean defaultLog = true;
        String arrlogFormats[] = new String[]{};
        String logFormats = ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFileFormat");
        if (logFormats.contains(",")) {
            arrlogFormats = logFormats.split(",");
            for (int logFormatInit = 0; logFormatInit < arrlogFormats.length; logFormatInit++) {
                if (arrlogFormats[logFormatInit].contains("Excel")) {
                    writeLogInToExcelFile(Category, LogString);
                }
                if (arrlogFormats[logFormatInit].contains("TextFile")) {
                    writeLogInToTextFile(Category, LogString);
                }
            }
        } else if (logFormats.equalsIgnoreCase("Excel")) {
            writeLogInToExcelFile(Category, LogString);
        } else if (logFormats.equalsIgnoreCase("TextFile")) {
            writeLogInToTextFile(Category, LogString);
        } else if (logFormats.equalsIgnoreCase("") && defaultLog) {
            writeLogInToTextFile(Category, LogString);
        }
        String propertiesValue = "<table border=1 width=100%><tr><td>" + getCustomDate() + "</td><td>" + getCustomTime() + "</td></tr></table>";
        WriteDataIntoPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath, "ReportLogInfo", propertiesValue);
    }

    public static void createTextLogFile(String logFileName) throws IOException {
        File myObj = new File(System.getProperty("user.dir") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFilePath") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFileName") + ".txt");
        if (myObj.createNewFile()) {
            System.out.println("File created: " + myObj.getName());
            PrintWriter out = new PrintWriter(System.getProperty("user.dir") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFilePath") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFileName") + ".txt"); // Step 2
            out.println("Date" + "\t\t\t" + "Time" + "\t\t\t" + "Category" + "\t\t" + "Description");
            out.close();
        } else {
            System.out.println("File already exists.");
        }
    }

    public static void createTempReportTextFile() throws IOException {
        File myObj = new File(System.getProperty("user.dir") + "\\" + "createTempReportTextFile.txt");
        if (myObj.createNewFile()) {
            //System.out.println("File created: " + myObj.getName());
            PrintWriter out = new PrintWriter(System.getProperty("user.dir") + "\\" + "createTempReportTextFile.txt"); // Step 2
            out.println("");
            out.close();
        } else {
            //System.out.println("File already exists.");
        }
    }

    public static void createLogFile() throws IOException {
        boolean isWindows = System.getProperty("os.name").toLowerCase().startsWith("windows");
        Process process;
        if (isWindows) {
            process = Runtime.getRuntime().exec(String.format("taskkill /f /im excel.exe"));
            process = Runtime.getRuntime().exec(String.format("taskkill /f /im notepad.exe"));
        }
        String logFileName = "Logs_" + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("customDataAndTimeStampForFileNames");
        if (ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogDataAppendedToExistingFileorNot").equalsIgnoreCase("Yes")) {
            System.out.println("Data is appended to existing log file");
        } else {
            WriteDataIntoPropertiesFile(globalPropertiesFileNameAndPath, "LogFileName", logFileName);
        }
        String arrlogFormats[] = new String[]{};
        String logFormats = ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFileFormat");
        if (logFormats.contains(",")) {
            arrlogFormats = logFormats.split(",");
            for (int logFormatInit = 0; logFormatInit < arrlogFormats.length; logFormatInit++) {
                if (arrlogFormats[logFormatInit].contains("Excel")) {
                    createExcelLogFile(logFileName);
                }
                if (arrlogFormats[logFormatInit].contains("TextFile")) {
                    createTextLogFile(logFileName);
                }
            }
        } else if (logFormats.equalsIgnoreCase("Excel")) {
            createExcelLogFile(logFileName);
        } else if (logFormats.equalsIgnoreCase("TextFile")) {
            createTextLogFile(logFileName);
        }
    }

    public static void writeLogFileHeader(String logFileName) throws FileNotFoundException {
        PrintWriter out = new PrintWriter(System.getProperty("user.dir") + "\\src\\main\\java\\Logs\\" + logFileName + ".txt"); // Step 2
        out.println("Date" + "\t\t\t" + "Time" + "\t\t\t" + "Category" + "\t\t" + "Description");
        out.close();
    }

    public static void writeLogInToTextFile(String Category, String LogString) throws IOException {
        /*Categories - Info, Warning, Error*/
        FileWriter fw = new FileWriter(System.getProperty("user.dir") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFilePath") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFileName") + ".txt", true);
        PrintWriter out = new PrintWriter(fw);
        String customDate = getCustomDate();
        String customTime = getCustomTime();
        out.println(customDate + "\t\t" + customTime + "\t\t" + Category + "\t\t\t" + LogString);
        out.close();
        replaceTextInHTML("Scripts_Execution_Logs", " <TR><TD style=\"text-align: center\">" + customDate + "</TD><TD style=\"text-align: center\">" + customTime + "</TD><TD style=\"text-align: center\">" + Category + "</TD><TD>&nbsp;" + LogString + "</TD></TR>" + "Scripts_Execution_Logs");
    }

    public static void writeLogInToExcelFile(String Category, String LogString) throws IOException {
        // ZipSecureFile.setMinInflateRatio(0);
        FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFilePath") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFileName") + ".xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet xs = wb.getSheet("Logs");
        FileOutputStream fos = null;
        XSSFCell columnNumber = null;
        Row row = xs.createRow(xs.getLastRowNum() + 1);
        String arrLogCellValues[] = {getCustomDate(), getCustomTime(), Category, LogString};
        for (int writeLogCellValues = 0; writeLogCellValues <= 3; writeLogCellValues++) {
            Cell cellDate = row.createCell(writeLogCellValues);
            cellDate.setCellValue(arrLogCellValues[writeLogCellValues]);
        }
        fos = new FileOutputStream(System.getProperty("user.dir") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFilePath") + ReadDataFromPropertiesFile(globalPropertiesFileNameAndPath).getProperty("LogFileName") + ".xlsx");
        wb.write(fos);
        wb.close();
        fos.close();
    }

    public static void HTMLExecutionSummary() throws IOException {
        FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "\\src\\main\\java\\Test_Data\\TestData.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Executions");
        int rowNumber = sheet.getLastRowNum();
        String strPropFileValue = null;
        int passedCount = 0;
        int failedCount = 0;
        int notexecutedCount = 0;
        int totalExecutedCount = 0;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row1 = sheet.getRow(i);
            for (int j = 0; j <= row1.getLastCellNum(); j++) {
                strPropFileValue = ReadDataFromPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath).getProperty("HTMLReportExecutionMethodsInfo");
                Cell celll = row1.getCell(j);
                if (i == 0) {
                    WriteDataIntoPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath, "HTMLReportExecutionMethodsInfo", strPropFileValue + "<td style=\"text-align: center\"><b>" + String.valueOf(celll) + "<b></td>");
                } else if (j == 0 && i > 0) {
                    WriteDataIntoPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath, "HTMLReportExecutionMethodsInfo", strPropFileValue + "<td style=\"text-align: Left\">" + String.valueOf(celll) + "</td>");
                } else {
                    WriteDataIntoPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath, "HTMLReportExecutionMethodsInfo", strPropFileValue + "<td style=\"text-align: center\">" + String.valueOf(celll) + "</td>");
                }
            }
            WriteDataIntoPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath, "HTMLReportExecutionMethodsInfo", strPropFileValue + "</TR>");
        }
        WriteDataIntoPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath, "HTMLReportExecutionMethodsInfo", strPropFileValue + "</TABLE>");
        replaceTextInHTML("Current_Execution_Methods_List", ReadDataFromPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath).getProperty("HTMLReportExecutionMethodsInfo"));
        WriteDataIntoPropertiesFile(HTMLLogInfoPropertiesFileNameAndPath, "HTMLReportExecutionMethodsInfo", " <Table border=1 width=100% style=font-size:14px;><Tr>");
        //Adding Images to HTML Report
        replaceTextInHTML("Images_Base64_Code_For_Current_Execution", "");
        //Scripts_Execution_Logs
        replaceTextInHTML("Scripts_Execution_Logs", "");
    }

}