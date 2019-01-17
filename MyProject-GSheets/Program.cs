using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

//using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;

//// Syncfusion.XlsIO.WinForms
using Syncfusion.XlsIO;
//using System.IO;
using System.Reflection;


namespace SheetsQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        static void Main(string[] args)
        {

            ////** Syncfusion.XlsIO.WinForms
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {

                //Set the default application version as Excel 2016
                //excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2013;

                ///  ??  ///  Creating a Hello World sample //SRART//

                //New instance of ExcelEngine is created equivalent to launching Microsoft Excel with no workbooks open
                //Instantiate the spreadsheet creation engine
                //using (ExcelEngine excelEngine = new ExcelEngine())
                //{
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Assigns default application version
                application.DefaultVersion = ExcelVersion.Excel2013;

                //A new workbook is created equivalent to creating a new workbook in Excel
                //Create a workbook with 1 worksheet
                IWorkbook workbook = application.Workbooks.Create(1);

                //Access first worksheet from the workbook
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding text to a cell
                worksheet.Range["A1"].Text = "Hello World 3";

                //Saving the workbook to disk in XLSX format
                workbook.SaveAs("Sample2.xlsx");

                //Closing the workbook
                workbook.Close();

                //Dispose the Excel engine
                excelEngine.Dispose();
                //}
             
                ///  ??  ///  Creating a Hello World sample //END//


            }

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            //String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String spreadsheetId = "1Aip4X4PFZ58it31tcmswoRvfd65sLSAB37JSik7WbOs";
            //String range = "Class Data!A2:E";
            String range = "Financial Acc Statements!S2:X3";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("AccountsReceivable\tAccountsPayable\tWeeklyDeposits\tCashOnHand\tCreditLineBalance\tProductionHoursWorked");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", row[0], row[1], row[2], row[3], row[4], row[5]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.Read();
        }
    }
}


// Google Sheets -> API v4 https://developers.google.com/sheets/api/quickstart/dotnet
// Create Excel File Using C# Console Application https://www.csharp-console-examples.com/general/create-excel-file-using-c-console-application/
// write data to an existing Excel file https://www.syncfusion.com/kb/9107/how-to-write-data-to-an-existing-excel-file-in-c-vb-net // Syncfusion.XlsIO.WinForms //paid