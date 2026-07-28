namespace ServvistaWebAppAPI.Models
{
    // Matches the JSON shape produced by PushSubscription.toJSON() in the browser.
    public class PushSubscriptionKeys
    {
        public string P256dh { get; set; } = string.Empty;
        public string Auth { get; set; } = string.Empty;
    }

    public class PushSubscriptionRequest
    {
        public string TechCode { get; set; } = string.Empty;
        public string Endpoint { get; set; } = string.Empty;
        public PushSubscriptionKeys Keys { get; set; } = new();
        public string? UserAgent { get; set; }
    }

    public class PushUnsubscribeRequest
    {
        public string Endpoint { get; set; } = string.Empty;
    }

    public class StoredPushSubscription
    {
        public int ID { get; set; }
        public string TECH_CODE { get; set; } = string.Empty;
        public string ENDPOINT { get; set; } = string.Empty;
        public string P256DH { get; set; } = string.Empty;
        public string AUTH { get; set; } = string.Empty;
    }
}
