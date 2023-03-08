using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Kiota.Abstractions;

class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.TenantId, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }
    public static async Task<string> GetUserTokenAsync()
    {
        // Ensure credential isn't null
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Ensure scopes isn't null
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

        // Request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }
    public static Task<UserCollectionResponse?> GetUserAsync()
    {
        // Ensure client isn't null
        _=_userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

       return _userClient.Users.GetAsync( (u) =>
             {
                 // Only request specific properties
                 u.QueryParameters.Select = new string[] { "DisplayName", "Mail", "UserPrincipalName" };
             });
       
    }
    public static Task<MessageCollectionResponse?> GetInboxAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
       
        return _userClient.Me
            // Only messages from Inbox folder
            .MailFolders["Inbox"]
            .Messages.
            GetAsync((m) =>
            {
                // Only request specific properties
                m.QueryParameters.Select = new string[] { "From", "IsRead", "ReceivedDateTime", "Subject" };
                m.QueryParameters.Orderby = new string[] { "ReceivedDateTime desc" };
                m.QueryParameters.Top = 25;
              
            });
            
    }
    public static async Task SendMailAsync(string subject, string body, string recipient)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Create a new message
        var requestBody = new SendMailPostRequestBody
        {
        Message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                Content = body,
                ContentType = BodyType.Text
            },
            ToRecipients = new List<Recipient>
            {
            new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient
                }
            }
            }
        }
        };

        // Send the message
        await _userClient.Me
            .SendMail
            .PostAsync(requestBody);
    }
    // This function serves as a playground for testing Graph snippets
    // or other code
    public static Task<Task<EventCollectionResponse?>> MakeGraphCallAsync()
    {
        // INSERT YOUR CODE HERE
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return Task.FromResult(_userClient.Me.CalendarView.GetAsync((u) =>
        {
            u.QueryParameters.StartDateTime = "2020-01-01T19:00:00-08:00";
            u.QueryParameters.EndDateTime = "2020-01-02T19:00:00-08:00";
        }));
    }
}