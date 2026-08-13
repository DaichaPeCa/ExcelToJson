using System.Text.Json;

namespace ExcelToJson.Core.Infrastructure;

internal static class JsonOutput
{
    public static byte[] Serialize(JsonValue root)
    {
        using MemoryStream stream = new();
        using (Utf8JsonWriter writer = new(stream, new JsonWriterOptions { Indented = true }))
        {
            Write(writer, root);
        }

        return stream.ToArray();
    }

    private static void Write(Utf8JsonWriter writer, JsonValue value)
    {
        switch (value)
        {
            case JsonValue.Object objectValue:
                writer.WriteStartObject();
                foreach (KeyValuePair<string, JsonValue> property in objectValue.Properties)
                {
                    writer.WritePropertyName(property.Key);
                    Write(writer, property.Value);
                }

                writer.WriteEndObject();
                break;
            case JsonValue.Array arrayValue:
                writer.WriteStartArray();
                foreach (JsonValue item in arrayValue.Items)
                {
                    Write(writer, item);
                }

                writer.WriteEndArray();
                break;
            case JsonValue.String stringValue:
                writer.WriteStringValue(stringValue.Value);
                break;
            case JsonValue.Number numberValue:
                writer.WriteNumberValue(numberValue.Value);
                break;
            case JsonValue.Boolean booleanValue:
                writer.WriteBooleanValue(booleanValue.Value);
                break;
            case JsonValue.Null:
                writer.WriteNullValue();
                break;
            default:
                throw new InvalidOperationException($"未知のJSON値型です: {value.GetType().Name}");
        }
    }
}
