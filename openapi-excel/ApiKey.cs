using Microsoft.OpenApi.Models;

namespace openapi_excel
{
    public class ApiKey
    {
        public string Value { get; set; }
        public string Key { get; set; }
        public ParameterLocation In { get; set; }
    }
}