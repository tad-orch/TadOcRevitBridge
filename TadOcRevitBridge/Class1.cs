using System;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Newtonsoft.Json.Linq;

namespace TadOcRevitBridge
{
    /// <summary>
    /// Executes one inbox job per Raise() inside a proper Revit ExternalEvent context.
    /// This avoids blocking the UI thread during long-running operations like
    /// OpenAndActivateDocument, which would freeze Revit when called from Idling.
    /// </summary>
    internal sealed class BridgeJobHandler : IExternalEventHandler
    {
        private const string InboxDir = @"D:\TAD\revit-bridge\inbox";
        private const string OutboxDir = @"D:\TAD\revit-bridge\outbox";
        private const string ArchiveDir = @"D:\TAD\revit-bridge\archive";

        private readonly RevitBridgeCommandProcessor _commandProcessor;
        private readonly Encoding _utf8NoBom;

        public string RevitVersion { get; set; } = "unknown";

        public BridgeJobHandler(RevitBridgeCommandProcessor commandProcessor, Encoding utf8NoBom)
        {
            _commandProcessor = commandProcessor;
            _utf8NoBom = utf8NoBom;
        }

        public void Execute(UIApplication uiApp)
        {
            var firstJob = Directory.GetFiles(InboxDir, "*.json")
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();

            if (string.IsNullOrWhiteSpace(firstJob))
            {
                return;
            }

            var jobId = Path.GetFileNameWithoutExtension(firstJob);
            string tool = null;
            JObject response;

            try
            {
                var json = File.ReadAllText(firstJob, Encoding.UTF8);
                var request = BridgeJson.DeserializeRequest(json);

                jobId = string.IsNullOrWhiteSpace(request.JobId) ? jobId : request.JobId;
                tool = request.Tool;

                response = _commandProcessor.Execute(uiApp, request);
            }
            catch (BridgeCommandException ex)
            {
                response = BridgeResultFactory.CreateError(jobId, tool, RevitVersion, ex);
            }
            catch (IOException ex) when (IsFileLocked(ex))
            {
                // File still being written by the bridge process — skip; will retry on next raise.
                return;
            }
            catch (Exception ex)
            {
                response = BridgeResultFactory.CreateError(
                    jobId,
                    tool,
                    RevitVersion,
                    new BridgeCommandException("unhandled_exception", ex.Message, null, ex));
            }

            var resultFile = Path.Combine(OutboxDir, MakeSafeResultName(jobId) + ".result.json");
            File.WriteAllText(resultFile, BridgeJson.Serialize(response), _utf8NoBom);
            ArchiveRequest(firstJob);
        }

        public string GetName() => "TadOcRevitBridge.BridgeJobHandler";

        private static bool IsFileLocked(IOException ex)
        {
            var errorCode = ex.HResult & 0xFFFF;
            return errorCode == 32 || errorCode == 33;
        }

        private static void ArchiveRequest(string requestPath)
        {
            var archivedPath = Path.Combine(ArchiveDir, Path.GetFileName(requestPath));
            if (File.Exists(archivedPath))
            {
                File.Delete(archivedPath);
            }
            File.Move(requestPath, archivedPath);
        }

        private static string MakeSafeResultName(string jobId)
        {
            const string fallback = "revit-job";
            if (string.IsNullOrWhiteSpace(jobId))
            {
                return fallback;
            }

            var invalidChars = Path.GetInvalidFileNameChars();
            var safe = new string(jobId
                .Trim()
                .Select(ch => invalidChars.Contains(ch) ? '_' : ch)
                .ToArray());

            return string.IsNullOrWhiteSpace(safe) ? fallback : safe;
        }
    }

    public class App : IExternalApplication
    {
        private const string InboxDir = @"D:\TAD\revit-bridge\inbox";
        private const string OutboxDir = @"D:\TAD\revit-bridge\outbox";
        private const string ArchiveDir = @"D:\TAD\revit-bridge\archive";
        private const string AliveFile = @"D:\TAD\revit-bridge\outbox\revit-addin-alive.json";

        private readonly RevitBridgeCommandProcessor _commandProcessor = new RevitBridgeCommandProcessor();
        private readonly Encoding _utf8NoBom = new UTF8Encoding(false);

        private BridgeJobHandler _jobHandler;
        private ExternalEvent _externalEvent;
        private string _revitVersion = "unknown";

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                _revitVersion = application.ControlledApplication.VersionNumber ?? "unknown";

                Directory.CreateDirectory(InboxDir);
                Directory.CreateDirectory(OutboxDir);
                Directory.CreateDirectory(ArchiveDir);

                WriteAliveFile();

                _jobHandler = new BridgeJobHandler(_commandProcessor, _utf8NoBom)
                {
                    RevitVersion = _revitVersion
                };
                _externalEvent = ExternalEvent.Create(_jobHandler);

                application.Idling += OnIdling;

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            application.Idling -= OnIdling;
            _externalEvent?.Dispose();
            return Result.Succeeded;
        }

        private void OnIdling(object sender, IdlingEventArgs e)
        {
            var uiApp = sender as UIApplication;
            if (uiApp == null)
            {
                return;
            }

            // Keep the handler's version string current.
            var version = uiApp.Application.VersionNumber;
            if (!string.IsNullOrWhiteSpace(version))
            {
                _revitVersion = version;
                _jobHandler.RevitVersion = version;
            }

            // Only raise if there is at least one pending job.
            // ExternalEvent.Raise() is safe to call repeatedly — if the event is already
            // pending it returns ExternalEventRequest.Pending without queuing a duplicate.
            if (Directory.GetFiles(InboxDir, "*.json").Any())
            {
                _externalEvent.Raise();
            }
        }

        private void WriteAliveFile()
        {
            var alive = new JObject
            {
                ["ok"] = true,
                ["source"] = "revit-addin",
                ["status"] = "alive",
                ["machine"] = Environment.MachineName,
                ["user"] = Environment.UserName,
                ["revitVersion"] = _revitVersion,
                ["timeUtc"] = DateTime.UtcNow.ToString("o"),
                ["supportedTools"] = new JArray(SupportedTools.All)
            };

            File.WriteAllText(AliveFile, BridgeJson.Serialize(alive), _utf8NoBom);
        }
    }
}
