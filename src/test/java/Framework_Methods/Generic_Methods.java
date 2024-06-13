package Framework_Methods;

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
                "    <title>Toggle Table</title>\n" +
                "    <script src=\"https://cdn.jsdelivr.net/npm/chart.js\"></script>\n" +
                " <script src=\"https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js\"></script>" +
                "    <style>\n" +
                "        body, html {\n" +
                "            zoom: 97%;\n" +
                "            height: 99%;\n" +
                "            margin: 0;\n" +
                "            padding: 0;\n" +
                "        }\n" +
                "        /* Style for the scrollable panel */\n" +
                "        .scrollable-panel {\n" +
                "            width: 98%; /* Set your desired width */\n" +
                "            height: 475px; /* Set your desired height */\n" +
                "            overflow-y: scroll; /* Enable vertical scrolling */\n" +
                "            overflow-x: hidden; /* Hide horizontal scrollbar */\n" +
                "            border: 1px solid #ccc; /* Add a border for visualization */\n" +
                "            padding: 10px; /* Add padding to the content */\n" +
                "        }\n" +
                "        .container {\n" +
                "            display: flex;\n" +
                "            height: 99%;\n" +
                "        }\n" +
                "        .left-panel {\n" +
                "            flex: 75%;\n" +
                "            background-color: #f0f0f0; /* Adjust as needed */\n" +
                "        }\n" +
                "        .right-panel {\n" +
                "            flex: 25%;\n" +
                "            display: flex;\n" +
                "            flex-direction: column;\n" +
                "        }\n" +
                "        .right-top-panel {\n" +
                "            flex: 50%;\n" +
                "            background-color: #e0e0e0; /* Adjust as needed */\n" +
                "            display: flex;\n" +
                "            flex-direction: column;\n" +
                "            align-items: center; /* Center content horizontally */\n" +
                "            justify-content: center; /* Center content vertically */\n" +
                "        }\n" +
                "        .right-bottom-panel {\n" +
                "            flex: 50%;\n" +
                "            background-color: #e0e0e0; /* Adjust as needed */\n" +
                "            display: flex;\n" +
                "            flex-direction: column;\n" +
                "            align-items: center; /* Center content horizontally */\n" +
                "            justify-content: center; /* Center content vertically */\n" +
                "        }\n" +
                "        .left-top-panel {\n" +
                "            flex: 40%; /*background-color: #e0e0e0; /* Adjust as needed */\n" +
                "            display: flex;\n" +
                "            flex-direction: column; /*align-items: center; /* Center content horizontally */ /*justify-content: center; /* Center content vertically */\n" +
                "        }\n" +
                "        .left-bottom-panel {\n" +
                "            flex: 60%; /*background-color: #e0e0e0; /* Adjust as needed */\n" +
                "            display: flex;\n" +
                "            flex-direction: column;\n" +
                "            align-items: flex-start; /* Center content horizontally */\n" +
                "            justify-content: center; /* Center content vertically */\n" +
                "        }\n" +
                "        .tab {\n" +
                "            overflow: hidden;\n" +
                "            border: 1px solid #ccc;\n" +
                "            background-color: #f1f1f1;\n" +
                "        }\n" +
                "        /* Style the buttons inside the tab */\n" +
                "        .tab button {\n" +
                "            background-color: inherit;\n" +
                "            float: left;\n" +
                "            border: none;\n" +
                "            outline: none;\n" +
                "            cursor: pointer;\n" +
                "            padding: 14px 16px;\n" +
                "            transition: 0.3s;\n" +
                "        }\n" +
                "        /* Change background color of buttons on hover */\n" +
                "        .tab button:hover {\n" +
                "            background-color: #ddd;\n" +
                "        }\n" +
                "        /* Create an active/current tablink class */\n" +
                "        .tab button.active {\n" +
                "            background-color: #ccc;\n" +
                "        }\n" +
                "        /* Style the tab content */\n" +
                "        .tabcontent {\n" +
                "            display: none;\n" +
                "            width: 99%; /* Ensure the content occupies 100% of the available width */\n" +
                "            padding: 6px 12px;\n" +
                "            border: 1px solid #ccc;\n" +
                "            border-top: none;\n" +
                "        }\n" +
                "        table {\n" +
                "            width: 99%;\n" +
                "            border-collapse: collapse;\n" +
                "            border: 2px solid #ddd; /* Border around the table */\n" +
                "            font-family: 'Verdana', serif; /* Specify the font family */\n" +
                "            font-size: 13px; /* Set the font size */\n" +
                "            margin-bottom: 05px; /* Adjust the value as needed */\n" +
                "        }\n" +
                "        th {\n" +
                "            border: 1px solid #ddd; /* Border around header cells */\n" +
                "            padding: 8px;\n" +
                "            text-align: center;\n" +
                "            background-color: #00008B; /* Background color for header cells */\n" +
                "            color: white; /* Text color for header cells */" +
                "        }\n" +
                "        td {\n" +
                "            border: 1px solid #ddd; /* Border around data cells */\n" +
                "            padding: 8px;\n" +
                "            text-align: center;\n" +
                "        }\n" +
                "        a {\n" +
                "            text-decoration: none; /* Remove underline */\n" +
                "        }\n" +
                "        .toggle-table {\n" +
                "            display: none; /* Hide by default */\n" +
                "        }\n" +
                "        .pagination {\n" +
                "            display: inline-block;\n" +
                "            margin-top: 0px;\n" +
                "            font-family: Arial, sans-serif;\n" +
                "        }\n" +
                "        .pagination a {\n" +
                "            color: black;\n" +
                "            float: left;\n" +
                "            padding: 4px 10px;\n" +
                "            text-decoration: none;\n" +
                "            transition: background-color .3s;\n" +
                "            border: 1px solid #ddd;\n" +
                "            font-weight: bold;\n" +
                "        }\n" +
                "        .pagination a.active {\n" +
                "            background-color: #4CAF50;\n" +
                "            color: white;\n" +
                "            border: 1px solid #4CAF50;\n" +
                "        }\n" +
                "        .pagination a:hover:not(.active) {\n" +
                "            background-color: #ddd;\n" +
                "        }\n" +
                "    </style>\n" +
                "</head>\n" +
                "<body>\n" +
                "<button id=\"exportButton\">Export to Excel</button>" +
                "<div class=\"container\">\n" +
                "    <div class=\"left-panel\">\n" +
                "        <div class=\"left-top-panel\">\n" +
                "            <table>\n" +
                "                <tr>\n" +
                "                    <th>Executed By</th>\n" +
                "                    <th>Execution Date</th>\n" +
                "                    <th>Total Executions</th>\n" +
                "                    <th>Passed Executions</th>\n" +
                "                    <th>Failed Executions</th>\n" +
                "                </tr>\n" +
                "                <tr>\n" +
                "                    <td style=\"width: 340px; text-align: center;\">Executed_By</td>\n" +
                "                    <td>Execution_Date</td>\n" +
                "                    <td>Total_Summary_Executions</td>\n" +
                "                    <td>Passed_Summary_Executions</td>\n" +
                "                    <td>Failed_Summary_Executions</td>\n" +
                "                </tr>\n" +
                "            </table>\n" +
                "            <table id=\"SummaryTable\">\n" +
                "                <tr>\n" +
                "                    <th>Workflow</th>\n" +
                "                    <th>Browser</th>\n" +
                "                    <th>Start Time</th>\n" +
                "                    <th>End Time</th>\n" +
                "                    <th>Duration</th>\n" +
                "                    <th>Status</th>\n" +
                "                </tr>\n" +
                "                summarytablesreplacetext\n" +
                "            </table>\n" +
                "            <div class=\"pagination\" id=\"pagination\"></div>\n" +
                "        </div>\n" +
                "        <div class=\"left-bottom-panel\">\n" +
                "            <div class=\"scrollable-panel\">\n" +
                "                <div class=\"tab\">\n" +
                "                    <button class=\"tablinks\" onclick=\"openTab(event, 'testCases')\">Test Cases</button>\n" +
                "                    <button class=\"tablinks\" onclick=\"openTab(event, 'logs')\">Logs</button>\n" +
                "                </div>\n" +
                "                <div id=\"testCases\" class=\"tabcontent\">\n" +
                "                    testcasetablereplace\n" +
                "                </div>\n" +
                "                <div id=\"logs\" class=\"tabcontent\">\n" +
                "                    logtablereplace\n" +
                "                </div>\n" +
                "            </div>\n" +
                "        </div>\n" +
                "    </div>\n" +
                "    <div class=\"right-panel\">\n" +
                "        <div class=\"right-top-panel\">\n" +
                "            <h2>Execution Summary</h2>\n" +
                "            <canvas id=\"myPieChart\" width=\"300\" height=\"300\"></canvas>\n" +
                "            <script>\n" +
                "                var pieChartData = {\n" +
                "                    labels: ['Passed Scripts', 'Failed Scripts'],\n" +
                "                    datasets: [{\n" +
                "                        data: [Passed_Summary_Executions, Failed_Summary_Executions],\n" +
                "                        backgroundColor: [\n" +
                "                            'rgba(0, 255, 0, 0.5)',\n" +
                "                            'rgba(255, 0, 0, 0.5)'\n" +
                "                        ],\n" +
                "                        borderColor: [\n" +
                "                            'rgba(0, 255, 0, 0.5)',\n" +
                "                            'rgba(255, 0, 0, 0.5)'\n" +
                "                        ],\n" +
                "                        borderWidth: 1\n" +
                "                    }]\n" +
                "                };\n" +
                "                var ctx = document.getElementById('myPieChart').getContext('2d');\n" +
                "                var myPieChart = new Chart(ctx, {\n" +
                "                    type: 'pie',\n" +
                "                    data: pieChartData,\n" +
                "                    options: {\n" +
                "                        responsive: false, // Disable responsiveness\n" +
                "                        maintainAspectRatio: false // Disable aspect ratio maintenance\n" +
                "                    }\n" +
                "                });\n" +
                "            </script>\n" +
                "        </div>\n" +
                "        <div class=\"right-bottom-panel\">\n" +
                "            <h2>Test Case Summary</h2>\n" +
                "            <div id=\"passCountsTestCases\"></div>\n" +
                "           <div id=\"failCountsTestCases\"></div>\n" +
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
                "    var tableId = 'SummaryTable';\n" +
                "    var table = document.getElementById(tableId);\n" +
                "    var rowsPerPage = 5;\n" +
                "    var rows = table.rows.length - 1;\n" +
                "    var pageCount = Math.ceil(rows / rowsPerPage);\n" +
                "    var pagination = document.getElementById('pagination');\n" +
                "\n" +
                "    function createPaginationButtons() {\n" +
                "        var firstButton = createButton('First', function () {\n" +
                "            showPage(table, 1);\n" +
                "        });\n" +
                "        pagination.appendChild(firstButton);\n" +
                "        var prevButton = createButton('Previous', function () {\n" +
                "            var currentPage = getCurrentPage();\n" +
                "            if (currentPage > 1) {\n" +
                "                showPage(table, currentPage - 1);\n" +
                "            }\n" +
                "        });\n" +
                "        pagination.appendChild(prevButton);\n" +
                "        for (var i = 1; i <= pageCount; i++) {\n" +
                "            var link = createButton(i, function () {\n" +
                "                var pageNumber = parseInt(this.innerHTML);\n" +
                "                showPage(table, pageNumber);\n" +
                "            });\n" +
                "            pagination.appendChild(link);\n" +
                "        }\n" +
                "        var nextButton = createButton('Next', function () {\n" +
                "            var currentPage = getCurrentPage();\n" +
                "            if (currentPage < pageCount) {\n" +
                "                showPage(table, currentPage + 1);\n" +
                "            }\n" +
                "        });\n" +
                "        pagination.appendChild(nextButton);\n" +
                "        var lastButton = createButton('Last', function () {\n" +
                "            showPage(table, pageCount);\n" +
                "        });\n" +
                "        pagination.appendChild(lastButton);\n" +
                "    }\n" +
                "\n" +
                "    function createButton(label, onclick) {\n" +
                "        var button = document.createElement('a');\n" +
                "        button.href = '#';\n" +
                "        button.innerHTML = label;\n" +
                "        button.addEventListener('click', onclick);\n" +
                "        return button;\n" +
                "    }\n" +
                "\n" +
                "    function getCurrentPage() {\n" +
                "        var activeLink = pagination.querySelector('.active');\n" +
                "        return parseInt(activeLink.innerHTML);\n" +
                "    }\n" +
                "\n" +
                "    function showPage(table, pageNumber) {\n" +
                "        var start = (pageNumber - 1) * rowsPerPage + 1;\n" +
                "        var end = Math.min(start + rowsPerPage - 1, rows);\n" +
                "\n" +
                "        for (var i = 1; i <= rows; i++) {\n" +
                "            var row = table.rows[i];\n" +
                "            if (i >= start && i <= end) {\n" +
                "                row.style.display = 'table-row';\n" +
                "            } else {\n" +
                "                row.style.display = 'none';\n" +
                "            }\n" +
                "        }\n" +
                "        updateActiveLink(pageNumber);\n" +
                "    }\n" +
                "\n" +
                "    function updateActiveLink(pageNumber) {\n" +
                "        var links = pagination.getElementsByTagName('a');\n" +
                "        for (var i = 0; i < links.length; i++) {\n" +
                "            links[i].classList.remove('active');\n" +
                "            if (parseInt(links[i].innerHTML) === pageNumber) {\n" +
                "                links[i].classList.add('active');\n" +
                "            }\n" +
                "        }\n" +
                "    }\n" +
                "\n" +
                "    createPaginationButtons();\n" +
                "    showPage(table, 1);\n" +
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
                "            var heading= document.createElement(\"h2\");\n" +
                "\t    heading.style.textAlign = \"center\";\n" +
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
                "function displayImage(base64) {\n" +
                "            var width = 1000;\n" +
                "            var height = 500;\n" +
                "            var left = (window.innerWidth - width) / 2;\n" +
                "            var top = (window.innerHeight - height) / 2;\n" +
                "            var newWindow = window.open(\"\", \"_blank\", \"width=\" + width + \",height=\" + height + \",left=\" + left + \",top=\" + top + \",toolbar=no,location=no,menubar=no,status=no,titlebar=no\");\n" +
                "            newWindow.document.write(\"<style>body {margin: 0;}</style><img src='\" + base64 + \"' style='max-width: 100%; max-height: 100%;'>\");\n" +
                "        }\n" +
                "</script>\n" +
                "<script>\n" +
                "document.getElementById(\"exportButton\").addEventListener(\"click\", function() {\n" +
                "    var wb = XLSX.utils.book_new();\n" +
                "    exportTableToExcel(\"SummaryTable\", \"Summary\", wb);\n" +
                "    exportTablesWithPrefixToExcel(\"Table_For_Logs_\", \"Logs\", wb);\n" +
                "    exportTablesWithPrefixToExcel(\"Table_For_TestCases_\", \"Test Cases\", wb);\n" +
                "    XLSX.writeFile(wb, 'Export_Execution_Report.xlsx');\n" +
                "});\n" +
                "function exportTableToExcel(tableId, sheetName, workbook) {\n" +
                "    var table = document.getElementById(tableId);\n" +
                "    var wsData = [];\n" +
                "    var rows = Array.from(table.querySelectorAll('tr')); // Get all rows of the table\n" +
                "    rows.forEach(function(row) {\n" +
                "        var rowData = [];\n" +
                "        row.querySelectorAll('th, td').forEach(function(cell) {\n" +
                "            rowData.push(cell.textContent);\n" +
                "        });\n" +
                "        wsData.push(rowData);\n" +
                "    });\n" +
                "    var ws = XLSX.utils.aoa_to_sheet(wsData);\n" +
                "    XLSX.utils.book_append_sheet(workbook, ws, sheetName);\n" +
                "}\n" +
                "function exportTablesWithPrefixToExcel(tablePrefix, sheetNamePrefix, workbook) {\n" +
                "    var tables = document.querySelectorAll('[id^=\"' + tablePrefix + '\"]');\n" +
                "    tables.forEach(function(table, index) {\n" +
                "        var sheetName = sheetNamePrefix + \" \" + (index + 1);\n" +
                "        exportTableToExcel(table.id, sheetName, workbook); // Pass table id to exportTableToExcel\n" +
                "    });\n" +
                "}\n" +
                "</script>" +
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
                    valueForTable = valueForTable + "\n<td style=\"width: 475px; text-align: left;\">\n<a href=\"#\" onclick=\"(function(){ toggleTable('" + "Table_For_TestCases_" + row.getCell(0) + "','" + "Table_For_Logs_" + row.getCell(0) + "'); countPassTextInColumns('Table_For_TestCases_" + row.getCell(0) + "'); return false; })();\">" + cellValue + "</a>\n</td>";
                    //valueForTable = valueForTable + "<td style=\"width: 475px; text-align: left;\"><a href=\"#\" onclick=\"(function(){ toggleTable('Table_For_TestCases_" + row.getCell(0) + "','Table_For_Logs_" + row.getCell(0) + "'); countPassTextInColumns('Table_For_TestCases_" + row.getCell(0) + "'); return false; })();\">" + cellValue + "</a></td>";

                } else {
                    //valueForTable = valueForTable + "<td><a href=\"#\" onclick=\"toggleTable('" + "Table_For_TestCases_" + row.getCell(0) + "','" + "Table_For_Logs_" + row.getCell(0) + "')\">" + cellValue + "</a></td>\n";
                    valueForTable = valueForTable + "\n<td >\n<a href=\"#\" onclick=\"(function(){ toggleTable('" + "Table_For_TestCases_" + row.getCell(0) + "','" + "Table_For_Logs_" + row.getCell(0) + "'); countPassTextInColumns('Table_For_TestCases_" + row.getCell(0) + "'); return false; })();\">" + cellValue + "</a>\n</td>";
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
            valueForTable = "<table border=\"1\" id=\"Table_For_TestCases_" + value + "\" class=\"toggle-table\">\n<tr>\n<th>Test Case</th>\n<th>Image</th>\n<th>Status</th></tr>\n";
            for (int testCases = 1; testCases <= xs.getLastRowNum(); testCases++) {
                Row row = xs.getRow(testCases);
                if (value.equals(row.getCell(0).toString())) {
                    valueForTable = valueForTable + "<tr>\n";
                    for (int testCaseCol = 1; testCaseCol < row.getLastCellNum(); testCaseCol++) {

                        Cell cellValue = row.getCell(testCaseCol);
                        if (testCaseCol == 1) {
                            valueForTable = valueForTable + "<td style=\"width: 475px; text-align: left;\">" + cellValue.toString() + "</td>\n";
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
            valueForTable = "<table border=\"1\" id=\"Table_For_Logs_" + value + "\" class=\"toggle-table\"><tr>\n<th>Date</th>\n<th>Time</th>\n<th>Log Type</th>\n<th>Description</th>\n</tr>\n";
            for (int logs = 1; logs <= xs.getLastRowNum(); logs++) {
                Row row = xs.getRow(logs);
                if (value.equals(row.getCell(0).toString())) {
                    valueForTable = valueForTable + "<tr>\n";
                    for (int logs_Cols = 1; logs_Cols < row.getLastCellNum(); logs_Cols++) {
                        Cell cellValue = row.getCell(logs_Cols);
                        if (logs_Cols == 4) {
                            valueForTable = valueForTable + "<td style=\"width: 600px; text-align: left;\">" + cellValue.toString() + "</td>";
                        } else {
                            valueForTable = valueForTable + "<td>" + cellValue.toString() + "</td>";
                        }
                    }
                    valueForTable = valueForTable + "</tr>\n";
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
                String arr[] = {"Thread ID", "Workflow", "Browser", "Start Time", "End Time", "Duration", "Status"};
                //HashMap<String, String> logData = new HashMap();
                sheet.setZoom(100);
                for (int cell = 0; cell < arr.length; cell++) {
                    Cell cellValue = row.createCell(cell);
                    cellValue.setCellValue(arr[cell]);
                    sheet.setColumnWidth(cell, 5000);
                }
                XSSFSheet sheetForTestCases = workbook.createSheet("TestCases_Info");
                Row rowValueForTestCases = sheetForTestCases.createRow(0);
                String[] arrForTestCasesHeaders = {"Thread ID", "Test_Case", "Image", "Status"};
                sheet.setZoom(100);
                for (int cell = 0; cell < arrForTestCasesHeaders.length; cell++) {
                    Cell cellValue = rowValueForTestCases.createCell(cell);
                    cellValue.setCellValue(arrForTestCasesHeaders[cell]);
                    sheetForTestCases.setColumnWidth(cell, 5000);
                }
                XSSFSheet sheetForLogs = workbook.createSheet("Logs_Info");
                Row rowValueForLogs = sheetForLogs.createRow(0);
                String arrForLogHeaders[] = {"Thread ID", "Date", "Time", "Log Type", "Description"};
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