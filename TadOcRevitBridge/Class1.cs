using System;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Newtonsoft.Json.Linq;

namespace TadOcRevitBridge
{
    public class App : IExternalApplication
    {
        private const string InboxDir = @"D:\TAD\revit-bridge\inbox";
        private const string OutboxDir = @"D:\TAD\revit-bridge\outbox";
        private const string ArchiveDir = @"D:\TAD\revit-bridge\archive";
        private const string AliveFile = @"D:\TAD\revit-bridge\outbox\revit-addin-alive.json";

        private RevitBridgeCommandProcessor _commandProcessor;
        private ActivateDocumentEventHandler _activateDocumentHandler;
        private ExternalEvent _activateDocumentEvent;
        private readonly Encoding _utf8NoBom = new UTF8Encoding(false);
        private bool _isProcessing;
        private string _revitVersion = "unknown";

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                _revitVersion = application.ControlledApplication.VersionNumber ?? "unknown";

                Directory.CreateDirectory(InboxDir);
                Directory.CreateDirectory(OutboxDir);
                Directory.CreateDirectory(ArchiveDir);

                _activateDocumentHandler = new ActivateDocumentEventHandler(OutboxDir, _utf8NoBom, () => _revitVersion);
                _activateDocumentEvent = ExternalEvent.Create(_activateDocumentHandler);
                _commandProcessor = new RevitBridgeCommandProcessor(_activateDocumentHandler, _activateDocumentEvent);

                WriteAliveFile();

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
            return Result.Succeeded;
        }

        private void OnIdling(object sender, IdlingEventArgs e)
        {
            if (_isProcessing)
            {
                return;
            }

            try
            {
                _isProcessing = true;

                var uiApp = sender as UIApplication;
                if (uiApp == null)
                {
                    return;
                }

                _revitVersion = uiApp.Application.VersionNumber ?? _revitVersion;
                ProcessInboxOnce(uiApp);
            }
            finally
            {
                _isProcessing = false;
            }
        }

        private void ProcessInboxOnce(UIApplication uiApp)
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
                response = BridgeResultFactory.CreateError(jobId, tool, _revitVersion, ex);
            }
            catch (Exception ex)
            {
                response = BridgeResultFactory.CreateError(
                    jobId,
                    tool,
                    _revitVersion,
                    new BridgeCommandException("unhandled_exception", ex.Message, null, ex));
            }

            // A null response means the handler is deferred (e.g. revit_activate_document uses
            // IExternalEventHandler and writes its own result). Archive the request now to
            // prevent re-processing on the next idle tick; the handler writes to outbox itself.
            if (response == null)
            {
                ArchiveRequest(firstJob);
                return;
            }

            var resultFile = Path.Combine(OutboxDir, BridgeResultFactory.MakeSafeResultName(jobId) + ".result.json");
            File.WriteAllText(resultFile, BridgeJson.Serialize(response), _utf8NoBom);

            ArchiveRequest(firstJob);
        }

        private void ArchiveRequest(string requestPath)
        {
            var archivedPath = Path.Combine(ArchiveDir, Path.GetFileName(requestPath));

            if (File.Exists(archivedPath))
            {
                File.Delete(archivedPath);
            }

            File.Move(requestPath, archivedPath);
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
