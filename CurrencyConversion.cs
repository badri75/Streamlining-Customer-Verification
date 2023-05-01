using System;
using System.Net.Http;
using System.Text.Json;

public class CurrencyConversion
{
    static decimal GetExchangeRate(string currencyCode)
    {
        string url = $"https://api.exchangeratesapi.io/latest?base={currencyCode}&symbols=EUR";
        using (var client = new HttpClient())
        {
            var response = client.GetAsync(url).Result;
            var content = response.Content.ReadAsStringAsync().Result;
            var json = JsonDocument.Parse(content);
            decimal exchangeRate = json.RootElement.GetProperty("rates").GetProperty("EUR").GetDecimal();
            return exchangeRate;
        }
    }

    public static string ConvertToEuros(string input)
    {
        string[] parts = input.Split(' ');
        string currencyCode = parts[0];
        decimal amount = decimal.Parse(parts[1]);

        decimal exchangeRate = GetExchangeRate(currencyCode);
        decimal euroAmount = amount * exchangeRate;
        string finalAmt = "EUR " + euroAmount.ToString(); 

        return finalAmt;
    }
}
