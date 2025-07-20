using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace ReportPolicy.Services.Report
{
    public class ReportEmailService
    {
        private readonly IWebHostEnvironment _env;

        public ReportEmailService(IWebHostEnvironment env)
        {
            _env = env;
        }

        public async Task<bool> SendReportAsync(List<Dictionary<string, object>> allData)
        {
            if (allData.Count == 0)
                return false;

            var now = DateTime.Now;
            string excelPath = Path.Combine(Path.GetTempPath(), "Report Policy Production.xlsx");

            // สร้าง Excel
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Report");
                var headers = new[] { "no", "name", "description", "operatingSystem", "memberOf" };

                // หัวตาราง
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(1, i + 1).Value = headers[i];
                    worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                }

                // ข้อมูล
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

            // สร้าง HTML table
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

            foreach (var email in recipients)
            {
                mail.To.Add(email);
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
            }
            finally
            {
                // ปล่อยการยึดไฟล์
                attachment.Dispose();
                mail.Dispose();

                try
                {
                    if (File.Exists(excelPath))
                        File.Delete(excelPath);
                }
                catch (IOException ex)
                {
                    // ลบไม่สำเร็จเพราะมีการใช้งาน
                    Console.WriteLine("Warning: Cannot delete file " + ex.Message);
                }
            }

            return true;
        }

        private string BuildHtmlTable(List<Dictionary<string, object>> data)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;font-size:13px;'>");

            if (data.Count > 0)
            {
                sb.AppendLine("<tr style='background-color:#f2f2f2;font-weight:bold;'>");
                foreach (var key in data[0].Keys)
                    sb.AppendLine($"<th>{System.Net.WebUtility.HtmlEncode(key)}</th>");
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
                        List<object> list => string.Join("<br>", list.Select(v => System.Net.WebUtility.HtmlEncode(v.ToString()))),
                        string str => string.Join("<br>", str.Split(',').Select(s => System.Net.WebUtility.HtmlEncode(s.Trim()))),
                        _ => System.Net.WebUtility.HtmlEncode(value.ToString())
                    };

                    sb.AppendLine($"<td>{display}</td>");
                }
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</table>");
            return sb.ToString();
        }


    }
}
