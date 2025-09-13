namespace Server.Infrastructure.Logging;

public interface IAppLogger<T>
{
    void LogInformation(string message);
    void LogInformation(string message, params object[] args);
    void LogWarning(string message);
    void LogWarning(string message, params object[] args);
    void LogError(Exception ex, string message);
    void LogError(Exception ex, string message, params object[] args);
} 