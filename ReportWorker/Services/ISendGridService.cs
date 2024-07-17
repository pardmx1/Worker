using SendGrid.Helpers.Mail;

namespace ReportWorker;

public interface ISendGridService
{
    EmailAddress GetSenderConfig(IConfiguration configuration);
    Task<bool> SendEmailTemplate(EmailAddress senderData, string body, string subject, List<string> attFiles, List<string> fileNames, List<string> toS, List<string> ccS);
}
