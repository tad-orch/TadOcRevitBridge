using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TadOcRevitBridge
{
    internal static class SupportedTools
    {
        public static readonly string[] All =
        {
            "revit_create_wall",
            "revit_session_status",
            "revit_open_cloud_model",
            "revit_activate_document",
            "revit_list_3d_views",
            "revit_export_nwc"
        };
    }

    internal sealed class BridgeRequest
    {
        [JsonProperty("jobId")]
        public string JobId { get; set; }

        [JsonProperty("tool")]
        public string Tool { get; set; }

        [JsonProperty("createdAt")]
        public string CreatedAt { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }

        [JsonProperty("payload")]
        public JObject Payload { get; set; }
    }

    internal static class BridgeJson
    {
        private static readonly JsonSerializer PayloadSerializer = JsonSerializer.Create(
            new JsonSerializerSettings
            {
                MissingMemberHandling = MissingMemberHandling.Ignore,
                NullValueHandling = NullValueHandling.Include
            });

        public static BridgeRequest DeserializeRequest(string json)
        {
            try
            {
                var request = JsonConvert.DeserializeObject<BridgeRequest>(json);
                if (request == null)
                {
                    throw new BridgeCommandException("invalid_request", "Request JSON did not contain a valid job envelope.");
                }

                return request;
            }
            catch (JsonException ex)
            {
                throw new BridgeCommandException("invalid_request", "Request file is not valid JSON.", null, ex);
            }
        }

        public static T ReadPayload<T>(BridgeRequest request) where T : class, new()
        {
            if (request == null || request.Payload == null)
            {
                return new T();
            }

            try
            {
                return request.Payload.ToObject<T>(PayloadSerializer) ?? new T();
            }
            catch (JsonException ex)
            {
                throw new BridgeCommandException("invalid_payload", "Request payload could not be parsed for this tool.", null, ex);
            }
        }

        public static string Serialize(JObject value)
        {
            return value.ToString(Formatting.Indented);
        }
    }

    internal static class BridgeResultFactory
    {
        public static JObject CreateSuccess(string jobId, string tool, string revitVersion)
        {
            return CreateEnvelope(true, "completed", jobId, tool, revitVersion);
        }

        public static JObject CreateError(string jobId, string tool, string revitVersion, BridgeCommandException exception)
        {
            var result = CreateEnvelope(false, "failed", jobId, tool, revitVersion);
            var error = new JObject
            {
                ["code"] = exception?.Code ?? "unknown_error",
                ["message"] = exception?.Message ?? "Unknown Revit add-in error."
            };

            if (exception?.Details != null && exception.Details.HasValues)
            {
                error["details"] = exception.Details;
            }

            result["error"] = error;
            return result;
        }

        private static JObject CreateEnvelope(bool ok, string status, string jobId, string tool, string revitVersion)
        {
            return new JObject
            {
                ["ok"] = ok,
                ["source"] = "revit-addin",
                ["status"] = status,
                ["jobId"] = jobId ?? string.Empty,
                ["tool"] = tool,
                ["revitVersion"] = revitVersion,
                ["timeUtc"] = DateTime.UtcNow.ToString("o")
            };
        }
    }

    internal sealed class BridgeCommandException : Exception
    {
        public BridgeCommandException(string code, string message, JObject details = null, Exception innerException = null)
            : base(message, innerException)
        {
            Code = code;
            Details = details;
        }

        public string Code { get; private set; }

        public JObject Details { get; private set; }
    }

    internal sealed class WallPayload
    {
        [JsonProperty("start")]
        public Point3 Start { get; set; }

        [JsonProperty("end")]
        public Point3 End { get; set; }

        [JsonProperty("height")]
        public double Height { get; set; }

        [JsonProperty("level")]
        public string Level { get; set; }

        [JsonProperty("wallType")]
        public string WallType { get; set; }
    }

    internal sealed class Point3
    {
        [JsonProperty("x")]
        public double X { get; set; }

        [JsonProperty("y")]
        public double Y { get; set; }

        [JsonProperty("z")]
        public double Z { get; set; }
    }

    internal sealed class OpenCloudModelPayload
    {
        [JsonProperty("region")]
        public string Region { get; set; }

        [JsonProperty("projectGuid")]
        public string ProjectGuid { get; set; }

        [JsonProperty("modelGuid")]
        public string ModelGuid { get; set; }

        [JsonProperty("openInUi")]
        public bool OpenInUi { get; set; }

        [JsonProperty("audit")]
        public bool Audit { get; set; }

        [JsonProperty("worksets")]
        public WorksetOpenRequest Worksets { get; set; }

        [JsonProperty("cloudOpenConflictPolicy")]
        public string CloudOpenConflictPolicy { get; set; }
    }

    internal sealed class WorksetOpenRequest
    {
        [JsonProperty("mode")]
        public string Mode { get; set; }
    }

    internal sealed class List3DViewsPayload
    {
        [JsonProperty("onlyExportable")]
        public bool OnlyExportable { get; set; }

        [JsonProperty("excludeTemplates")]
        public bool? ExcludeTemplates { get; set; }
    }

    internal sealed class ExportNwcPayload
    {
        [JsonProperty("viewNames")]
        public List<string> ViewNames { get; set; }

        [JsonProperty("outputPath")]
        public string OutputPath { get; set; }

        [JsonProperty("exportScope")]
        public string ExportScope { get; set; }
    }

    internal sealed class ActivateDocumentPayload
    {
        [JsonProperty("documentTitle")]
        public string DocumentTitle { get; set; }
    }
}
