using Microsoft.Data.SqlClient;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Net.Mail;
using System.Net;
using System.Net.Mime;
using System.Reflection.PortableExecutable;
namespace ReportWorker;

public class Worker : IHostedService, IDisposable
{
    private Timer _timer = null;
    private DateTime currentInterval;
    private List<DateTime> dts;
    private Dictionary<DateTime, TimeSpan> intervals;
    private readonly ILogger<Worker> _logger;
    private readonly IConfiguration _configuration;
    private readonly IServiceScopeFactory _scopeFactory;
    private List<Catalog> catalog;
    private List<Layout> layout;
    private List<string> fileNames = new List<string>();
    private List<string> fileC = new List<string>();
    private string mailBody;
    public Worker(ILogger<Worker> logger, IConfiguration configuration, IServiceScopeFactory scopeFactory)
    {
        _logger = logger;
        _configuration = configuration;
        _scopeFactory = scopeFactory;
    }

    public void Dispose()
    {
        _timer?.Dispose();
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        //currentInterval = DateTime.Now;
        //GetReports(null);
        StartReportTask();
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        _timer?.Change(Timeout.Infinite, 0);
        return Task.CompletedTask;
    }

    private async void StartReportTask()
    {
        TimeZoneInfo timeZoneInfo;
        try
        {
            timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time (Mexico)");
        }
        catch (TimeZoneNotFoundException tex)
        {
            timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("America/Mexico_City");
        }
        var currentTImeMex = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, timeZoneInfo);

        var hrs = _configuration.GetSection("SendHrs").Get<List<string>>();
        var hrsDt = new List<TimeSpan>();
        dts = new List<DateTime>();
        hrs.ForEach(h =>
        {
            var ch = DateTime.Parse(h);
            Console.WriteLine(ch.TimeOfDay);
            hrsDt.Add(ch.TimeOfDay);
            dts.Add(new DateTime(DateOnly.FromDateTime(currentTImeMex), TimeOnly.FromDateTime(ch)));
        });

        var nextRunTime = currentTImeMex.Date;
        hrsDt.ForEach(h =>
        {
            if (nextRunTime < currentTImeMex)
            {
                nextRunTime = currentTImeMex.Date;
                nextRunTime = nextRunTime.Add(h);
                Console.WriteLine(nextRunTime);
            }

        });

        intervals = new Dictionary<DateTime, TimeSpan>();
        for (int i = 0; i < dts.Count; i++)
        {
            var nrDif = TimeSpan.Zero;
            if (i + 1 == dts.Count)
            {
                var ndt = dts[0].AddDays(1);
                nrDif = ndt.Subtract(dts[i]);
                Console.WriteLine(nrDif);
            }
            else
            {
                nrDif = dts[i + 1].Subtract(dts[i]);
                Console.WriteLine(nrDif);
            }
            intervals.Add(dts[i], nrDif);
        }
        var firstInterval = nextRunTime.Subtract(currentTImeMex);

        if (firstInterval < TimeSpan.FromMinutes(1))
        {
            firstInterval = TimeSpan.FromMinutes(1);
        }
        currentInterval = nextRunTime;

        _logger.LogInformation("Next run time: {time}", nextRunTime);
        _logger.LogInformation("Current time: {time}", currentTImeMex);
        _logger.LogInformation("First interval {time}", firstInterval);

        Action action = () =>
        {
            _timer = new Timer(
                GetReports,
                null,
                firstInterval,
                TimeSpan.Zero
                );
        };

        await Task.Run(action);
    }

    private async void GetReports(Object state)
    {
        try
        {
            var reportMonths = _configuration["ReportConfig:Months"];

            var ConnectionStrings = _configuration.GetSection("CS").Get<List<ConnectionString>>();

            ConnectionStrings.ForEach(cs =>
            {
                using (SqlConnection connection = new SqlConnection(cs.Cs))
                {
                    connection.Open();

                    string sql = string.Format(File.ReadAllText("./Scripts/catalog.sql"), reportMonths);

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            catalog = new List<Catalog>();
                            while (reader.Read())
                            {
                                var c = new Catalog
                                {
                                    RowNo = reader.GetInt64(3),
                                    GroupName = reader.GetString(0),
                                    Variable = reader.GetString(1),
                                    Alias = reader.GetString(2)
                                };
                                catalog.Add(c);
                            }
                        }
                    }

                    sql = string.Format(File.ReadAllText("./Scripts/layout.sql"), reportMonths);


                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            layout = new List<Layout>();
                            while (reader.Read())
                            {
                                var l = new Layout
                                {
                                    Pais = reader.GetString(0),
                                    Code = reader.GetString(1),
                                    CreationDate = reader.GetDateTime(2),
                                    Status = reader.GetString(3)
                                };
                                layout.Add(l);
                                Console.WriteLine("{0}", reader.GetString(0));
                            }
                        }
                    }
                    catalog = catalog.OrderBy(c => c.RowNo).ToList();

                    layout.ForEach(l =>
                    {
                        l.ExtraFlieds = new Dictionary<string, string>();
                        catalog.ForEach(c =>
                        {
                            l.ExtraFlieds.Add($"{c.RowNo}_{c.Alias}", "");
                        });

                    });

                    sql = string.Format(File.ReadAllText("./Scripts/selectdata.sql"), reportMonths);
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var rowToMod = layout.Where(x => x.Code == reader.GetString(3)).FirstOrDefault();
                                var Variable = reader.GetString(2);
                                var varObj = catalog.Where(v => v.Variable == Variable).FirstOrDefault();
                                if (varObj != null)
                                {
                                    rowToMod.ExtraFlieds[$"{varObj.RowNo}_{varObj.Alias}"] = reader.IsDBNull(0) ? "" : reader.GetString(0);
                                }

                            }
                        }
                    }
                }
                CreateExcelDoc(cs.Pais);
            });
            SendEmail();
        }
        catch (SqlException e)
        {
            Console.WriteLine(e.ToString());
        }
    }

    public void CreateExcelDoc(string pais)
    {
        var date = DateTime.Now.ToShortDateString();
        date = date.Replace('/', '-');
        var fileName = $"./Scripts/Reporte {pais} {date}.xlsx"; //"./Scripts/test.xlsx"
        fileNames.Add(fileName);
        MemoryStream ms = new MemoryStream();
        using (SpreadsheetDocument document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet" };

            sheets.Append(sheet);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            Row row = new Row();
            row.Append(CreateCell(nameof(Layout.Pais)));
            row.Append(CreateCell(nameof(Layout.Code)));
            row.Append(CreateCell(nameof(Layout.CreationDate)));
            row.Append(CreateCell(nameof(Layout.Status)));

            foreach (var ef in layout[0].ExtraFlieds)
            {
                row.Append(CreateCell(ef.Key));
            }

            sheetData.Append(row);

            layout.ForEach(l =>
            {
                row = new Row();
                row.Append(CreateCell(l.Pais));
                row.Append(CreateCell(l.Code));
                row.Append(CreateCell(l.CreationDate.ToString()));
                row.Append(CreateCell(l.Status));
                foreach (var ef in l.ExtraFlieds)
                {
                    row.Append(CreateCell(ef.Value));
                }
                sheetData.Append(row);
            });

            workbookPart.Workbook.Save();
            document.Save();
            using (MemoryStream auxms = new MemoryStream())
            {
                ms.Seek(0, SeekOrigin.Begin);
                ms.CopyTo(auxms);
                var docBytes = auxms.ToArray();
                var docB64 = Convert.ToBase64String(docBytes);
                fileC.Add(docB64);
            }

        }
    }

    private async void SendEmailSG()
    {
        using (var scope = _scopeFactory.CreateAsyncScope())
        {
            var sendGridService = scope.ServiceProvider.GetRequiredService<ISendGridService>();
            var senderData = sendGridService.GetSenderConfig(_configuration);
            var toS = _configuration.GetSection("SmtpConfig:ToS").Get<List<string>>();
            var ccS = _configuration.GetSection("SmtpConfig:CcS").Get<List<string>>();
            var res = await sendGridService.SendEmailTemplate(senderData, "Reporte LATAM", _configuration["SmtpConfig:Subject"], fileC, fileNames, toS, ccS);
            fileC = new List<string>();
            fileNames = new List<string>();
            Console.WriteLine("EmailSended " + res);
            RestartTimer();
        }
    }

    private void RestartTimer()
    {
        _logger.LogInformation("Next run in: {time}", intervals[currentInterval]);
        _timer.Change(intervals[currentInterval], TimeSpan.Zero);
        var index = (dts.FindIndex(x => x == currentInterval) + 1) == dts.Count ? 0 : dts.FindIndex(x => x == currentInterval) + 1;
        currentInterval = dts[index];
    }

    private void SendEmail()
    {
        var mailBody = currentInterval.Hour == 6 ? _configuration["bodyMorning"] : _configuration["bodyAfternoon"];
        var toS = _configuration.GetSection("SmtpConfig:ToS").Get<List<string>>();
        var cCS = _configuration.GetSection("SmtpConfig:CcS").Get<List<string>>();
        MailMessage mailMessage = new MailMessage
        {
            From = new MailAddress(_configuration["SmtpConfig:From"], _configuration["EmailSender"]),
            Body = mailBody,
            IsBodyHtml = true
        };
        toS.ForEach(t =>
        {
            mailMessage.To.Add(new MailAddress(t));
        });
        if (cCS != null && cCS.Count > 0)
        {
            cCS.ForEach(c =>
            {
                mailMessage.CC.Add(new MailAddress(c));
            });
        }
        mailMessage.Subject = _configuration["SmtpConfig:Subject"];
        var fileIndex = 0;
        fileC.ForEach(f =>
        {
            MemoryStream ms = new MemoryStream(Convert.FromBase64String(f));
            var na = new Attachment(ms, fileNames[fileIndex].Replace("./Scripts/", ""), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            //var attachment = new Attachment(f, MediaTypeNames.Application.Octet);
            mailMessage.Attachments.Add(na);
            fileIndex++;
        });
        fileIndex = 0;
        // var attachment = new Attachment("./Scripts/test.xlsx", MediaTypeNames.Application.Octet);
        // mailMessage.Attachments.Add(attachment);
        SmtpClient client = new SmtpClient(_configuration["SmtpConfig:Server"])
        {
            Port = int.Parse(_configuration["SmtpConfig:Port"]),
            EnableSsl = true,
            UseDefaultCredentials = false
        };
        NetworkCredential cred = new System.Net.NetworkCredential(_configuration["SmtpConfig:ServerAcc"], _configuration["SmtpConfig:Password"]);
        client.Credentials = cred;
        client.Send(mailMessage);
        fileC = new List<string>();
        fileNames = new List<string>();
        Console.WriteLine("EmailSended ");
        RestartTimer();
    }

    private Cell CreateCell(string text)
    {
        Cell cell = new Cell
        {
            StyleIndex = 1U,
            DataType = ResolveCellDataTypeOnValue(text),//CellValues.String,
            CellValue = new CellValue(text)
        };
        return cell;
    }

    private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
    {
        int intVal;
        double doubleVal;
        if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
        {
            return CellValues.Number;
        }
        else
        {
            return CellValues.String;
        }
    }
}
