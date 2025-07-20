using ClosedXML.Excel;
using Newtonsoft.Json;
using System.DirectoryServices;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace ReportPolicy.Services.Report
{
    public class AdReportLogic
    {
        public async Task RunAsync()
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

                            row[propName] = cnValues.Count == 1 ? (object)cnValues[0] : cnValues.Cast<object>().ToList();
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

            // เช็คเวลา
            var now = DateTime.Now;
            bool isMonday = now.DayOfWeek == DayOfWeek.Monday;
            bool isMorning = now.Hour == 7 && now.Minute < 45;
            bool isNoon = now.Hour == 12 && now.Minute < 15;
            string lastSentFile = "lastSent.txt";
            DateTime? lastSentTime = null;

            if (File.Exists(lastSentFile))
            {
                var content = File.ReadAllText(lastSentFile);
                if (DateTime.TryParse(content, out var parsedTime))
                    lastSentTime = parsedTime;
            }

            bool isNewPeriod = lastSentTime == null || (now - lastSentTime.Value).TotalHours >= 1;

            if (allData.Count > 0 && isMonday && (isMorning || isNoon) && isNewPeriod)
            {
                string excelPath = Path.Combine(Path.GetTempPath(), "Report Policy.xlsx");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Report");
                    var headers = new[] { "no", "name", "description", "operatingSystem", "memberOf" };

                    for (int i = 0; i < headers.Length; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = headers[i];
                        worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                    }

                    int rowIdx = 2;
                    foreach (var row in allData)
                    {
                        for (int i = 0; i < headers.Length; i++)
                        {
                            if (row.TryGetValue(headers[i], out var value))
                            {
                                if (value is IEnumerable<string> stringList)
                                {
                                    worksheet.Cell(rowIdx, i + 1).Value = string.Join(", ", stringList);
                                }
                                else if (value is IEnumerable<object> objectList && !(value is string))
                                {
                                    worksheet.Cell(rowIdx, i + 1).Value = string.Join(", ", objectList);
                                }
                                else
                                {
                                    worksheet.Cell(rowIdx, i + 1).Value = value?.ToString();
                                }
                            }
                        }
                        rowIdx++;
                    }

                    workbook.SaveAs(excelPath);
                }

                string htmlTable = BuildHtmlTable(allData);

                var mail = new MailMessage
                {
                    From = new MailAddress("your@email", "Report Policy Production"),
                    Subject = "The list does not delete the production policy",
                    IsBodyHtml = true,
                    Body = $@"
                     <html><body style='font-family:Arial;'>
                        <p>Dear Team,</p>
                        <p><b>Report Time:</b> {now:yyyy-MM-dd HH:mm:ss}</p>
                        <p>This computer did not delete the policy:</p>
                        {htmlTable}
                        <p>Best regards,<br/>AD Report Bot</p>
                     </body></html>"
                };

                var recipients = new List<string>
                {
                    "your@email",
                  
                };

                 foreach (var email in recipients)
                  {
                  mail.To.Add(email);
                 }

                var ccRecipients = new List<string>
                {
                    "your@email",
                };

                foreach (var cc in ccRecipients)
                {
                    mail.CC.Add(cc);
                }


                var attachment = new Attachment(excelPath, MediaTypeNames.Application.Octet);
                attachment.ContentDisposition.FileName = "Report Policy Production.xlsx";
                mail.Attachments.Add(attachment);

                using var smtp = new SmtpClient("yourDomain")
                {
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = true
                };

                try
                {
                    smtp.Send(mail);
                    File.WriteAllText(lastSentFile, now.ToString("o"));
                }
                finally
                {
                    attachment.Dispose();
                    mail.Dispose();

                    try
                    {
                        if (File.Exists(excelPath))
                            File.Delete(excelPath);
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine("Could not delete temp file: " + ex.Message);
                    }
                }
            }
        }

        private string BuildHtmlTable(List<Dictionary<string, object>> data)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;font-size:13px;'>");

            if (data.Count > 0)
            {
                sb.AppendLine("<tr style='background-color:#f2f2f2;font-weight:bold;'>");
                foreach (var key in data[0].Keys)
                {
                    sb.AppendLine($"<th>{System.Net.WebUtility.HtmlEncode(key)}</th>");
                }
                sb.AppendLine("</tr>");
            }

            foreach (var row in data)
            {
                sb.AppendLine("<tr>");
                foreach (var value in row.Values)
                {
                    string display = value switch
                    {
                        null => "",
                        List<object> list => string.Join(", ", list),
                        _ => value.ToString()
                    };

                    sb.AppendLine($"<td>{System.Net.WebUtility.HtmlEncode(display)}</td>");
                }
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</table>");
            return sb.ToString();
        }

    }

}
