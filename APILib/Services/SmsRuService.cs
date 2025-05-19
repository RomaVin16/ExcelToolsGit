using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using APILib.Contracts;
using Microsoft.Extensions.Configuration;

namespace APILib.Services
{
    public class SmsRuService : ISmsService
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;

        public SmsRuService(IHttpClientFactory httpClientFactory, IConfiguration configuration)
        {
            _httpClient = httpClientFactory.CreateClient("SmsRu");
            _apiKey = configuration["SmsSettings:ApiKey"] ?? throw new ArgumentNullException("API key is missing");
        }

        public async Task SendAsync(string phoneNumber, string message, bool isTest = true)
        {
            if (isTest = true)
            {
                Console.WriteLine($"[TEST MODE] SMS to {phoneNumber}: {message}");
                return;
            }

            // Валидация номера и сообщения при необходимости
            var cleanPhone = new string(phoneNumber.Where(char.IsDigit).ToArray());
            if (!cleanPhone.StartsWith("7") || cleanPhone.Length != 11)
                throw new ArgumentException("Неверный формат номера. Нужен формат: 7XXXXXXXXXX");

            if (string.IsNullOrWhiteSpace(message) || message.Length > 160)
                throw new ArgumentException("Сообщение должно содержать 1–160 символов");

            // Подготовка параметров
            var parameters = new Dictionary<string, string>
            {
                { "sign", "SMS Aero" },         
                { "number", cleanPhone },
                { "text", message },
                { "channel", "DIRECT" },
                { "test", "1" }
                };

            var content = new FormUrlEncodedContent(parameters);


            const string email = "rvvinokurov@edu.hse.ru";  
            const string apiKey = "5E-S0UN4L4-SzmDywP-AIQ-ji1Ckf2_W";
            var authBytes = Encoding.ASCII.GetBytes($"{email}:{apiKey}");

            _httpClient.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(authBytes));

            var response = await _httpClient.PostAsync("https://gate.smsaero.ru/v2/sms/send", content);

            if (!response.IsSuccessStatusCode)
                throw new HttpRequestException($"HTTP error: {response.StatusCode}");

            var responseBody = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<SmsAeroResponse>(responseBody);

            if (result?.Success != true)
                throw new Exception($"Ошибка SMS Aero: {result?.Message ?? "Неизвестная ошибка"}");
        }

    }

    public class SmsAeroResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string? Message { get; set; }

        [JsonPropertyName("data")]
        public object? Data { get; set; } // или конкретная модель, если нужна
    }
}