using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

var spreadsheetId = "1XrH_-1JrSgZgXRuZhcy57aa1EVUXQ2ZV7mVUgAvib1g";
var range = "Words";

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

var credential = GoogleCredential.FromJson(@"{
    ""type"": ""service_account"",
    ""project_id"": ""i3rothers"",
    ""private_key_id"": ""1e9beac7b870ae06fe26c5189d9cdf8c813225c6"",
    ""private_key"": ""-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCo2oBG0+lt/SjZ\nyBGp+MV302DlvDKinDjtbr8A7rhJfJEaa4GBV5z/mngFADKB7qj6YIasgFfI/6BQ\nY4fFy2pLEbWtituwac8/uSC9OnndRgMqB3ml6SSewmudR9tzyw+Dq2/XtSDMuIGf\nqt/YWsrHezPeVGilQJrbGShLPAbTLNBDfDxE3GeJsjYEC3tum3GTERGqWn0MFLVV\nVPm8QCBcRiDiyNZfehzs1WPKS4riTBsZudPHCnqHXLvnq6TrgEElKC99NmqjDW40\nMtf/XwOQY42SIcIUkTFKMSp0qHdJzATZYv9DZJZQkKDc0mbyCUPk4jq0Ns5KNfOk\ntOdVlfQ/AgMBAAECggEAFrsXcrF7Zq2ibrLywc12E9v2Wub7cA2U0k3K077PO0FU\nLVjUnctkHLq4NoAgzIIK6G3PI3DBoJLdC4VilTt9iy6OpRag3X5ZAoyS/jZdim6u\n5V0NQMsCfYbRwuIWBmALLiAJleHR0Q1zkcGIkdEjJDiPMnptffWVAzK/TGV/y7Sw\nHJfv+OvTt94e8wPl32l3R+jScjnUu6O6Bm568/9jaZ2HvcdM8IxIXuSoruqdFOCX\nn//45h4fuPVj4U3QZzPdqCdBeakh4x/JeRHtOi0Ob1AvOf8SKY3axSKmVAKqFb/R\nWkxThPwxer01kinhvxh9hZB1LeLl7gliqSfXGKMabQKBgQDdQRF8E9j4UhYZbjFH\nWbzUUnf2vLFNtBTJCpDQALxFgnmVrY/M3MLT0c0OAswsAZJqe0qJohLkxw4bq7iP\nyBEKVvFCYfVBPgnAlNdYaeafwHHT3+jEYFv25THajTPuqn3pOwb1n8ycOx9zIe4V\nAwpbBx65DbFMOvHbwb0f+BypuwKBgQDDXs4Hj+sdSGpH39W89Xpp+AntJzFntab/\nGtSata/yyhpE66Kbr80uWNoQ0aFZaI/kKhbg7Vsd8FOXqyo8wqB5WtUFIn232/ae\nLNxAXjMfr0+9ByGkeVFoj3cyHSLoKY7eyBwAx6e/nfXVhd19Pd7uB5eGSH2XgT+/\n5Gd4qG7FTQKBgQCnKeE2+JvmScauogWTXeaAGGrQvZHMHnHRzyzIKrYUYlbQUpih\n9G0ysoGVw2FVIj7oOox/XjeeKBKtr1k7MLJHOJcBS5eMGn4txYbKIwD+09xscvCf\nZho1eMbo0+RXvvJwg4tniruBkl3Zk9oYf/qT+dYphIHfEW3oVgE5JTEqvwKBgQCA\na+YUNHcA9aPfAPRXVCkWVRP5ToT8Pfy6vaE43Or+NfkUiquFmQbPS1p0KcfcpI3J\nFh2Z1ovJXzsjfEC0Vd70Rk+2I1juLWmryaMxsHn8ftl0UKa9nX10tLFOQLa8Uuz1\n5iX6IUNUAnog0/Cmra/HWTgx7Z6Yoz4LXhDh0B2YFQKBgE7BGYBxGui0qAygrtBU\n5LBpDej+weCuit/CFDJoKHjSYF4gN872L1r5jD8lvDlBoBSvnFbqc0EJHwyNlIi0\naCQnbu5znIzPvV21iJlX7WpYWVQ8z8sXGC6UcnYX1A7ZQ0wwgaF/8lRL1bKTtN3e\n5Xkfe6R6QBI4I7ewIjGtGz9a\n-----END PRIVATE KEY-----\n"",
    ""client_email"": ""words-573@i3rothers.iam.gserviceaccount.com"",
    ""client_id"": ""110276966851059466845"",
    ""auth_uri"": ""https://accounts.google.com/o/oauth2/auth"",
    ""token_uri"": ""https://oauth2.googleapis.com/token"",
    ""auth_provider_x509_cert_url"": ""https://www.googleapis.com/oauth2/v1/certs"",
    ""client_x509_cert_url"": ""https://www.googleapis.com/robot/v1/metadata/x509/words-573%40i3rothers.iam.gserviceaccount.com"",
    ""universe_domain"": ""googleapis.com""
}")
.CreateScoped(SheetsService.Scope.Spreadsheets);

app.MapGet("/words", () =>
{
    var sheetsService = new SheetsService(new BaseClientService.Initializer() { 
        HttpClientInitializer = credential
    });

    var response = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range).Execute();
    
    var words = new List<Word>();

    foreach (var row in response.Values)
    {
        var word = new Word(
            row[0]?.ToString() ?? string.Empty,
            row[1]?.ToString() ?? string.Empty,
            row[2]?.ToString() ?? string.Empty,
            row[3]?.ToString() ?? string.Empty,
            int.TryParse(row[4].ToString(), out var point) ? point : 0
        );
        words.Add(word);
    }
    return words;
})
.WithOpenApi();

app.MapPost("/words/{row}", (int row, int point) =>
{
    var sheetsService = new SheetsService(new BaseClientService.Initializer() { 
        HttpClientInitializer = credential
    });
    var updateAction = sheetsService.Spreadsheets.Values.Update(new ValueRange
    {
        Values = new List<IList<object>> { new List<object> { point } }
    }, spreadsheetId, $"{range}!E{row}");

    updateAction.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
    updateAction.Execute();
    
    return Results.Ok();
})
.WithOpenApi();

app.Run();

record Word(string text, string type, string pronunciation, string meaning, int point);
