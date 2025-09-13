using System.Diagnostics;

namespace Server.Infrastructure.Logging
{
    public interface IAppLogger<T>
    {
        void LogInformation(string message, object? args = null);
        void LogWarning(string message, object? args = null);
        void LogError(string message, Exception exception, object? args = null);
    }

    public class AppLogger<T> : IAppLogger<T>
    {
        private readonly ILogger<T> _logger;

        public AppLogger(ILogger<T> logger)
        {
            _logger = logger;
        }

        public void LogInformation(string message, object? args = null)
        {
            var context = GetContext();
            _logger.LogInformation("{Message} | Controller/Service: {Context.Class} | Method: {Context.Method} | Args: {@Args}",
                message, context.Class, context.Method, args);
        }

        public void LogWarning(string message, object? args = null)
        {
            var context = GetContext();
            _logger.LogWarning("{Message} | Controller/Service: {Context.Class} | Method: {Context.Method} | Args: {@Args}",
                message, context.Class, context.Method, args);
        }

        public void LogError(string message, Exception exception, object? args = null)
        {
            var context = GetContext();
            _logger.LogError(exception, "{Message} | Controller/Service: {Context.Class} | Method: {Context.Method} | Args: {@Args}",
                message, context.Class, context.Method, args);
        }

        private (string Class, string Method) GetContext()
        {
            var stackTrace = new StackTrace();
            for (int i = 2; i < stackTrace.FrameCount; i++)
            {
                var method = stackTrace.GetFrame(i)?.GetMethod();
                if (method == null) continue;

                var declaringType = method.DeclaringType;
                if (declaringType != null && declaringType != typeof(AppLogger<T>))
                {
                    return (declaringType.Name, method.Name);
                }
            }
            return ("UnknownClass", "UnknownMethod");
        }
    }
}
