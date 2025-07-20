using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.DirectoryServices;
using System.Globalization;

namespace ReportPolicy.Controllers
{
    public class GetadController : Controller
    {
        public IActionResult Index()
        {

            var username = "";
            var password = "*";

            string[] ouPaths = new string[]
                {
      
                "LDAP://"
               
                };

                var users = new List<Dictionary<string, string>>();
                string[] properties = new[]
                {
                    "sAMAccountName", "distinguishedName", "name", "accountExpires",
                    "displayName", "company", "title", "mail", "department",
                    "physicalDeliveryOfficeName", "userAccountControl", "pwdLastSet",
                    "msDS-UserPasswordExpiryTimeComputed", "logonCount", "lastLogonTimestamp", "manager"
                };

                foreach (var ouPath in ouPaths)
                {
                    DirectoryEntry entry = new DirectoryEntry(ouPath, username, password);
                    DirectorySearcher searcher = new DirectorySearcher(entry)
                    {
                        Filter = "(&(objectClass=user)(objectCategory=person))"
                    };


                    foreach (string prop in properties)
                        searcher.PropertiesToLoad.Add(prop);

                    SearchResultCollection results = searcher.FindAll();
                    foreach (SearchResult result in results)
                    {
                        var userDict = new Dictionary<string, string>();
                        foreach (string prop in properties)
                        {
                            if (result.Properties.Contains(prop))
                                userDict[prop] = result.Properties[prop][0]?.ToString() ?? "";
                            else
                                userDict[prop] = "";
                        }
                        users.Add(userDict);
                    }
                }
                string fileTimeString = "134280407230000000";
                CultureInfo userCulture = CultureInfo.CurrentCulture;

                DateTime baseDate = new DateTime(1601, 1, 1);
                string statusExpires = "Not Available";

                if (long.TryParse(fileTimeString, out long fileTime))
                {
                    statusExpires = (fileTime == 0 || fileTime == 9223372036854775807)
                        ? "No Expires"
                        : baseDate.AddTicks(fileTime).ToLocalTime().ToString("dd MMMM yyyy HH:mm", userCulture);
                }

                Console.WriteLine(statusExpires);
            using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("AD Users");

                    // Header
                    int col = 1;
                    foreach (var key in properties)
                    {
                        worksheet.Cell(1, col++).Value = key;
                    }

                    // Data
                    for (int i = 0; i < users.Count; i++)
                    {
                        int dataCol = 1;
                        foreach (var key in properties)
                        {
                            worksheet.Cell(i + 2, dataCol++).Value = users[i][key];
                        }
                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "AD_Users_Report.xlsx");
                    }
                }
            
        }

    }
}
