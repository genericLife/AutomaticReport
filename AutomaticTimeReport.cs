using System;
using System.Data;
using System.Net;
using System.Net.Mail;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Collections.Generic;
using System.IO;


namespace AutomaticReport
{
    //===== Time Report API =======================================================================
    public class UserTimeReport
    {
        [JsonPropertyName("user_id")]
        public int UserId { get; set; }

        [JsonPropertyName("user_name")]
        public string UserName { get; set; }

        [JsonPropertyName("total_hours")]
        public float TotalHours { get; set; }

        [JsonPropertyName("billable_hours")]
        public float BillableHours { get; set; }
    }

    public class Report
    {
        public List<UserTimeReport> Results { get; set; }
    }

    //===== Users =================================================================================
    public class User
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("first_name")]
        public string FirstName { get; set; }

        [JsonPropertyName("last_name")]
        public string LastName { get; set; }

        public string Email { get; set; }

        [JsonPropertyName("is_active")]
        public bool active { get; set; }

        [JsonPropertyName("roles")]
        public string[] roles { get; set; }

    }

    public class ReturnedUsers
    {
        public List<User> Users { get; set; }
    }

    class AutomaticTimeReport
    {
        // API request variables
        static JsonSerializerOptions options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };

        static string apiDate;
        static string baseUrlTime = "https://api.harvestapp.com/v2/reports/time/team?";

        static string UrlUsers = "https://api.harvestapp.com/v2/users";

        static Report[] reports = new Report[7];

        // DataTable variables
        static List<string> UserNames = new List<string>();

        static DataTable reportTable;


        //===== Main ==============================================================================
        static async Task Main(string[] args)
        {
            DateTime today = DateTime.Now;

            await GetUsers();

            if (today.ToString("ddd") == "Mon")
            {
                await GetApiDataForWeek(today.AddDays(-7));
                CreateDataTableWeek(today.AddDays(-7));
            } else if (today.ToString("ddd") != "Sun" && today.ToString("ddd") != "Sat")
            {
                SetDateYesterday();
                await GetApiDataForDay(0);
                CreateDataTableDay(today.AddDays(-1));
            }
        }

        //===== Api request functions =============================================================

        /* Sends GET request to obtain team reports from harvest API and stores the data 
           in reports[index] */
        public static async Task GetApiDataForDay(int index)
        {

            using var client = new HttpClient();
            var url = baseUrlTime + apiDate;
            var authToken = "Insert Token";
            var account_id = "Insert ID";

            // Headers
            client.DefaultRequestHeaders.Add("Authorization", authToken);
            client.DefaultRequestHeaders.Add("Harvest-Account-Id", account_id);
            client.DefaultRequestHeaders.Add("User-Agent", "PostmanRuntime/7.26.10");
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

            var result = await client.GetAsync(url);
            var content = await result.Content.ReadAsStringAsync();

            reports[index] = JsonSerializer.Deserialize<Report>(content, options);
            //PrintJson(reports[index]);
        }

        public static async Task GetApiDataForWeek(DateTime startOfWeek)
        {
            DateTime currentDay = startOfWeek;
            
            for (int i = 0; i < 7; i++)
            {
                SetApiDate(DateToString(currentDay));
                await GetApiDataForDay(i);
                currentDay = currentDay.AddDays(1);
            }
        }

        static async Task GetUsers()
        {
            using var client = new HttpClient();
            var authToken = "Insert Token";
            var account_id = "Insert ID";

            // Headers
            client.DefaultRequestHeaders.Add("Authorization", authToken);
            client.DefaultRequestHeaders.Add("Harvest-Account-Id", account_id);
            client.DefaultRequestHeaders.Add("User-Agent", "PostmanRuntime/7.26.10");
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

            var result = await client.GetAsync(UrlUsers);
            var content = await result.Content.ReadAsStringAsync();

            ReturnedUsers users = JsonSerializer.Deserialize<ReturnedUsers>(content, options);

            foreach (User user in users.Users)
            {
                if (user.active && !NoTimeReport(user.roles)) 
                    UserNames.Add(user.FirstName + " " + user.LastName);
            }
        }

        //===== Date functions for GET request ====================================================
        public static void SetApiDate(string date)
        {
            apiDate = "from="+date+"&to="+date;
        }

        public static void SetDateYesterday()
        {
            DateTime date = DateTime.Now;
            date = date.AddDays(-1);
            string yday = date.ToString("yyyy/MM/dd");
            yday = yday.Replace("/", "");
            SetApiDate(yday);
        }

        public static string DateToString(DateTime date)
        {
            string sDate = date.ToString("yyyy/MM/dd");
            sDate = sDate.Replace("/", "");
            return sDate;
        }

        public static string GetTableDate(DateTime date)
        {
            string sDate = date.ToString("ddd, MMM dd");
            return sDate;
        }

        //===== DataTable functions ===============================================================
        public static void CreateDataTableDay(DateTime date)
        {
            Object[] rows = new Object[UserNames.Count];
            Object[] UnorderedRows = new Object[UserNames.Count];
            DataColumn[] tableCols = new DataColumn[2];

            reportTable = new DataTable("Time Report for the day");
            
            // Column setup
            tableCols[0] = new DataColumn("User", typeof(string));
            tableCols[1] = new DataColumn(GetTableDate(date), typeof(float));
            reportTable.Columns.AddRange(tableCols);
            reportTable.PrimaryKey = new DataColumn[] { reportTable.Columns["User"] };

            int NumCols = 2;

            // Create and populate rows
            for (var i = 0; i < rows.Length; i++)
            {
                UnorderedRows[i] = CreateRow(date, UserNames[i], NumCols);
            }

            SortRows(UnorderedRows, ref rows, NumCols);

            foreach (Object[] row in rows) {
                reportTable.Rows.Add(row);
            }

            //GenerateReport(reportTable, date, false, "txt");
            string HtmlText = ExportDatatableToHtml(reportTable, date, false);
        }

        public static void CreateDataTableWeek(DateTime startDate)
        {
            CreateWeekColumns(startDate);
            Object[] rows = new Object[UserNames.Count];
            Object[] UnorderedRows = new Object[UserNames.Count];
            var NumCols = 9;

            for (var i = 0; i < rows.Length; i++)
            {
                UnorderedRows[i] = CreateRow(startDate, UserNames[i], NumCols);
            }

            SortRows(UnorderedRows, ref rows, NumCols);


            foreach (Object[] row in rows) {
                reportTable.Rows.Add(row);
            }

            //GenerateReport(reportTable, startDate, true, "txt");
            ExportDatatableToHtml(reportTable, startDate, true);
        }

        static void SortRows(Object[] UnorderedRows, ref Object[] rows, int NumCols)
        {
            int c = 0;
            foreach (Object[] row in UnorderedRows) {
                object temp = (object) row[NumCols-1];
                if (Convert.ToInt32(temp) == 0) {
                    rows[c] = row;
                    c++;
                }
            }
            
            foreach (Object[] row in UnorderedRows) {
                object temp = (object) row[NumCols-1];
                if (Convert.ToInt32(temp) != 0) {
                    rows[c] = row;
                    c++;
                }
            }
        }

        static void CreateWeekColumns(DateTime date)
        {
            DataColumn[] tableCols = new DataColumn[9];
            reportTable = new DataTable("Time Report for the week");
            tableCols[0] = new DataColumn("User", typeof(string));

            for (var i = 1; i < tableCols.Length-1; i++)
            {
                tableCols[i] = new DataColumn(GetTableDate(date), typeof(float));
                date = date.AddDays(1);
            }

            tableCols[8] = new DataColumn("Total", typeof(float));

            reportTable.Columns.AddRange(tableCols);
            reportTable.PrimaryKey = new DataColumn[] { reportTable.Columns["User"] };
        }

        static Object[] CreateRow(DateTime startDate, string user, int rowLength)
        {
            Object[] row = new Object[rowLength];
            row[0] = user;
            var length = rowLength;
            var total = 0.0;

            if (length == 9) length -= 1;

            for (var i = 1; i < length; i++)
            {
                row[i] = 0;
                
                foreach (UserTimeReport report in reports[i-1].Results)
                {
                    if (report.UserName == user)
                    {
                        row[i] = report.TotalHours;
                        total += report.TotalHours;
                    }
                }
            }
            if (rowLength > 8) row[8] = total;

            return row;
        }

        static string ExportDatatableToHtml(DataTable dt, DateTime date, bool week)  
        {
            string fileName;

            if (week) {
                fileName = "Reports/WeekReport" + date.ToString("dd-MM-yyyy") + ".html";
            } else {
                fileName = "Reports/Report" + date.ToString("dd-MM-yyyy") + ".html";
            }

            StringBuilder strHTMLBuilder = new StringBuilder();  
            strHTMLBuilder.AppendLine("<html >");
            strHTMLBuilder.AppendLine("<head>"); 
            strHTMLBuilder.AppendLine("<style>");
            strHTMLBuilder.AppendLine("body{");
            strHTMLBuilder.AppendLine("font-family: 'Calibri', sans-serif;");
            strHTMLBuilder.AppendLine("size: 11;}");
            strHTMLBuilder.AppendLine("table, th, td {");
            strHTMLBuilder.AppendLine("border: 1px solid black;}");
            strHTMLBuilder.AppendLine("table{");
            strHTMLBuilder.AppendLine("border-collapse: collapse;}");
            strHTMLBuilder.AppendLine("</style>");
            strHTMLBuilder.AppendLine("</head>");  
            strHTMLBuilder.AppendLine("<body>");  
            strHTMLBuilder.AppendLine("<table cellpadding='10' cellspacing='1' bgcolor='white'>");  
            
            strHTMLBuilder.AppendLine("<tr >");  
            foreach (DataColumn myColumn in dt.Columns)  
            {  
                strHTMLBuilder.AppendLine("<td >");  
                strHTMLBuilder.AppendLine(myColumn.ColumnName);  
                strHTMLBuilder.AppendLine("</td>");  
            
            }  
            strHTMLBuilder.AppendLine("</tr>");  
            
            
            foreach (DataRow myRow in dt.Rows)  
            {
                strHTMLBuilder.AppendLine("<tr >");  
                foreach (DataColumn myColumn in dt.Columns)  
                {  
                    
                    strHTMLBuilder.AppendLine("<td >");  
                    strHTMLBuilder.AppendLine(myRow[myColumn.ColumnName].ToString());  
                    strHTMLBuilder.AppendLine("</td>");  
            
                }  
                strHTMLBuilder.AppendLine("</tr>");  
            }  
            
            strHTMLBuilder.AppendLine("</table>");  
            strHTMLBuilder.AppendLine("</body>");  
            strHTMLBuilder.AppendLine("</html>");  
            
            string Htmltext = strHTMLBuilder.ToString();

            //File.WriteAllText(fileName, Htmltext);
            MailAddress to = new MailAddress("cmorrison9924@gmail.com");
            MailAddress from = new MailAddress("notifications@metisware.com");
            SendEmail(from, to, Htmltext);

            return Htmltext;
        } 

        //===== Emails ============================================================================

        static void SendEmail(MailAddress from, MailAddress to, string HtmlBody)
        {
            MailMessage mail = new MailMessage(from, to);

            mail.Subject = "Time Report " + DateTime.Now.AddDays(-1).ToString("ddd, MMM dd");
            mail.IsBodyHtml = true;
            mail.Body = HtmlBody;

            var client = new SmtpClient("smtp.office365.com", 587);
            client.EnableSsl = true;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential("Insert Email", "Insert Password");
            
            try {
                client.Send(mail);

            } catch(Exception ex) {
                //Error, could not send the message
                Console.Write(ex.Message);
            }
        }

        //===== Utility ===========================================================================
        public static void PrintJsonReport(Report report)
        {
            var serializedReport = JsonSerializer.Serialize<Report>(report, options);
            Console.WriteLine(serializedReport);
        }
        
        public static void PrintJsonUsers(ReturnedUsers users)
        {
            var serializedReport = JsonSerializer.Serialize<ReturnedUsers>(users, options);
            Console.WriteLine(serializedReport);
        }

        public static bool NoTimeReport(string[] roles)
        {
            foreach (string role in roles)
            {
                if (role == "NoTimeReport") return true;
            }
            return false;
        }
        
        static void GenerateReport(DataTable table, DateTime date, bool week, string fileType) {
            
            string fileName;
            string format;

            if (week) {
                fileName = "Reports/WeekReport" + date.ToString("dd-MM-yyyy") + "." + fileType;
            } else {
                fileName = "Reports/Report" + date.ToString("dd-MM-yyyy") + "." + fileType;
            }

            if (fileType == "txt") format = "{0,-20}";
            else if (fileType == "csv") format = "{0:0.0},";
            else {
                Console.WriteLine("Expected file type of csv or txt");
                return;
            }

            using StreamWriter fileWriter = new(fileName);

            foreach (DataColumn col in table.Columns) {
                fileWriter.Write(format, col.ColumnName);
            }
            fileWriter.WriteLine();
            
            foreach (DataRow row in table.Rows) {
                foreach (DataColumn col in table.Columns) {
                    fileWriter.Write(format, row[col]);
                }
                fileWriter.WriteLine();
            }
            fileWriter.WriteLine();
        }
    } 

}


