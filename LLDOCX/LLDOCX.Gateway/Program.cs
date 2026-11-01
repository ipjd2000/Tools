using LLDOCX.Gateway;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices((context, services) =>
    {
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        var options = new GatewayOptions();
        context.Configuration.GetSection("Gateway").Bind(options);
        options.Validate();

        services.AddSingleton(options);
        services.AddSingleton<ITokenValidator>(new EntraTokenValidator(options, EntraTokenValidator.CreateConfigurationManager(options.TenantId)));
        services.AddSingleton(new UserRateLimiter(options.RequestsPerUserPerHour, () => DateTime.UtcNow));
        services.AddSingleton<IClaudeReviewer>(SdkClaudeReviewer.Create(options));
        services.AddSingleton<DeepCheckHandler>();
    })
    .Build();

host.Run();
