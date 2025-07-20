using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using ReportPolicy.Services.Report;
using System.DirectoryServices;
using System.Xml;
using Formatting = Newtonsoft.Json.Formatting;

namespace ReportPolicy.Controllers
{
    public class ReportController : Controller
    {
        private readonly ReportEmailService _reportEmailService;

        public ReportController(ReportEmailService reportEmailService)
        {
            _reportEmailService = reportEmailService;
        }


        public async Task<IActionResult> Index()
        {
            var allData = GetComputerData();
            return View(allData);
        }

        [HttpPost]
        public async Task<IActionResult> SendManualReport()
        {
            var allData = GetComputerData();
            bool success = await _reportEmailService.SendReportAsync(allData);
            TempData["Message"] = success ? "Email sent successfully." : "No data to send.";
            return RedirectToAction("Index");
        }

        private List<Dictionary<string, object>> GetComputerData()
        {
            var username = "";
            var password = "*";

            string[] ouPaths = new[]
            {
             "LDAP://"
            };

            var allData = new List<Dictionary<string, object>>();
            int no = 1;

            foreach (var ouPath in ouPaths)
            {
                using var entry = new DirectoryEntry(ouPath, username, password);
                using var searcher = new DirectorySearcher(entry)
                {
                    Filter = "(objectClass=computer)"
                };

                searcher.PropertiesToLoad.AddRange(new[] { "name", "description", "operatingSystem", "memberOf" });
                var results = searcher.FindAll();

                foreach (SearchResult result in results)
                {
                    if (!result.Properties.Contains("memberOf"))
                        continue;

                    bool isARALL = result.Properties["memberOf"]
                        .Cast<string>()
                        .Any(v => v.IndexOf("ARALL-Windows", StringComparison.OrdinalIgnoreCase) >= 0);

                    if (!isARALL)
                        continue;

                    var row = new Dictionary<string, object>
                    {
                        ["no"] = no++
                    };

                    string[] orderedProps = { "name", "description", "operatingSystem", "memberOf" };

                    foreach (var propName in orderedProps)
                    {
                        if (!result.Properties.Contains(propName))
                            continue;

                        var values = result.Properties[propName];

                        if (propName.Equals("memberOf", StringComparison.OrdinalIgnoreCase))
                        {
                            var cnValues = values
                                .Cast<string>()
                                .Select(v =>
                                {
                                    int startIndex = v.IndexOf("CN=", StringComparison.OrdinalIgnoreCase);
                                    int endIndex = v.IndexOf(",", startIndex + 3);
                                    if (startIndex >= 0)
                                    {
                                        return endIndex > 0
                                            ? v.Substring(startIndex + 3, endIndex - (startIndex + 3))
                                            : v.Substring(startIndex + 3);
                                    }
                                    return v;
                                }).ToList();

                            row[propName] = cnValues.Count == 1 ? (object)cnValues[0] : cnValues;
                        }
                        else
                        {
                            row[propName] = values.Count == 1
                                ? values[0]
                                : values.Cast<object>().ToList();
                        }
                    }

                    allData.Add(row);
                }
            }

            return allData;
        }

    }
}
