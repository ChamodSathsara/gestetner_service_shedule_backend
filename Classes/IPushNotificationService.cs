using ServvistaWebAppAPI.Models;

namespace ServvistaWebAppAPI.Classes
{
    public interface IPushNotificationService
    {
        Task<bool> SendJobAssignedNotificationAsync(
            BreakdownModel job,
            CancellationToken cancellationToken = default);
    }
}