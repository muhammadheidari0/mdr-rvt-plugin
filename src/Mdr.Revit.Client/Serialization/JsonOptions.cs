using System.Text;
using System.Text.Json;

namespace Mdr.Revit.Client.Serialization
{
    public static class JsonOptions
    {
        public static readonly JsonSerializerOptions Default = new JsonSerializerOptions
        {
            PropertyNamingPolicy = SnakeCaseNamingPolicy.Instance,
            DictionaryKeyPolicy = SnakeCaseNamingPolicy.Instance,
            PropertyNameCaseInsensitive = true,
            WriteIndented = false,
        };

        private sealed class SnakeCaseNamingPolicy : JsonNamingPolicy
        {
            public static readonly SnakeCaseNamingPolicy Instance = new SnakeCaseNamingPolicy();

            public override string ConvertName(string name)
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    return string.Empty;
                }

                StringBuilder builder = new StringBuilder(name.Length + 8);
                for (int i = 0; i < name.Length; i++)
                {
                    char current = name[i];
                    if (char.IsUpper(current))
                    {
                        if (i > 0)
                        {
                            char previous = name[i - 1];
                            bool shouldInsert =
                                char.IsLower(previous) ||
                                char.IsDigit(previous) ||
                                (char.IsUpper(previous) && i + 1 < name.Length && char.IsLower(name[i + 1]));
                            if (shouldInsert)
                            {
                                builder.Append('_');
                            }
                        }

                        builder.Append(char.ToLowerInvariant(current));
                    }
                    else
                    {
                        builder.Append(current);
                    }
                }

                return builder.ToString();
            }
        }
    }
}
