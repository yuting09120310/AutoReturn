using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace AutoReturn
{
    public class Settings
    {
        [JsonPropertyName("connectionString")]
        public string ConnectionString { get; set; }

        [JsonPropertyName("apiUrl")]
        public string ApiUrl { get; set; }

        [JsonPropertyName("token")]
        public string Token { get; set; }
    }

}
