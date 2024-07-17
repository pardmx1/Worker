using SendGrid;
using SendGrid.Helpers.Mail;

namespace ReportWorker;

public class SendGridService : ISendGridService
{
    private readonly IConfiguration _configuration;
    private readonly ISendGridClient _sendGridClient;
    public SendGridService(ISendGridClient sendGridClient, IConfiguration configuration)
    {
        _sendGridClient = sendGridClient;
        _configuration = configuration;
    }
    public EmailAddress GetSenderConfig(IConfiguration configuration)
    {
        try
        {
            string? fromName = _configuration["EmailSender"];
            string? fromAdd = _configuration["EmailUser"];

            return new EmailAddress(fromAdd, fromName);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            return null;
        }
    }

    public async Task<bool> SendEmailTemplate(EmailAddress senderData, string body, string subject, List<string> attFiles, List<string> fileNames, List<string> toS, List<string> ccS)
    {
        var tos = new List<EmailAddress>();

        toS.ForEach(t =>
        {
            tos.Add(new EmailAddress(t));
        });

        var emailMsg = MailHelper.CreateSingleEmailToMultipleRecipients(senderData, tos, subject, "", body, true);
        var ccs = new List<EmailAddress>();
        if (ccS != null && ccS.Count > 0)
        {
            ccS.ForEach(c =>
            {
                ccs.Add(new EmailAddress(c));
            });
            emailMsg.AddCcs(ccs);
        }

        var atts = new List<Attachment>();
        int i = 0;
        attFiles.ForEach(a =>
        {
            //var ba = File.ReadAllBytes(a);
            var na = new Attachment
            {
                Content = a,//Convert.ToBase64String(ba),
                Filename = fileNames[i].Replace("./Scripts/", ""),
                Type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                Disposition = "attachment"
            };
            atts.Add(na);
            i++;
        });
        emailMsg.AddAttachments(atts);
        var response = await _sendGridClient.SendEmailAsync(emailMsg);
        Console.WriteLine(await response.Body.ReadAsStringAsync());
        if (response.IsSuccessStatusCode)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
}
