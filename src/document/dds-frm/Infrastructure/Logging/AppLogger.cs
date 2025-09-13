using Microsoft.Extensions.Logging;

namespace Server.Infrastructure.Logging;

public class AppLogger<T> : IAppLogger<T>
{
    private readonly ILogger<T> _logger;

    public AppLogger(ILogger<T> logger)
    {
        _logger = logger;
    }

    public void LogInformation(string message)
    {
        _logger.LogInformation(message);
    }

    public void LogInformation(string message, params object[] args)
    {
        _logger.LogInformation(message, args);
    }

    public void LogWarning(string message)
    {
        _logger.LogWarning(message);
    }

    public void LogWarning(string message, params object[] args)
    {
        _logger.LogWarning(message, args);
    }

    public void LogError(Exception ex, string message)
    {
        _logger.LogError(ex, message);
    }

    public void LogError(Exception ex, string message, params object[] args)
    {
        _logger.LogError(ex, message, args);
    }
} 