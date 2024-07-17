using ReportWorker;
using SendGrid;
using SendGrid.Extensions.DependencyInjection;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddSingleton<ISendGridService, SendGridService>(serviceProvider => new SendGridService(serviceProvider.GetRequiredService<ISendGridClient>(), serviceProvider.GetRequiredService<IConfiguration>()));
builder.Services.AddHostedService<Worker>();
builder.Services.AddSendGrid(options =>
{
    options.ApiKey = builder.Configuration["SendGridEmailApi"];
});

var app = builder.Build();

app.MapGet("/", () => "Worker Running");

app.Run();
