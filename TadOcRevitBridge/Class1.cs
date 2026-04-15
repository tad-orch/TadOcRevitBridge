using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
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

        private readonly RevitBridgeCommandProcessor _commandProcessor = new RevitBridgeCommandProcessor();
        private readonly Encoding _utf8NoBom = new UTF8Encoding(false);
        private bool _isProcessing;
        private string _revitVersion = "unknown";

        private ActivateDocumentEventHandler _activateDocumentHandler;
        private ExternalEvent _activateDocumentEvent;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                _revitVersion = application.ControlledApplication.VersionNumber ?? "unknown";

                Directory.CreateDirectory(InboxDir);
                Directory.CreateDirectory(OutboxDir);
                Directory.CreateDirectory(ArchiveDir);

                WriteAliveFile();

                _activateDocumentHandler = new ActivateDocumentEventHandler(OutboxDir);
                _activateDocumentEvent = ExternalEvent.Create(_activateDocumentHandler);

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
            JObject response = null;
            var asyncDispatched = false;

            try
            {
                var json = File.ReadAllText(firstJob, Encoding.UTF8);
                var request = BridgeJson.DeserializeRequest(json);

                jobId = string.IsNullOrWhiteSpace(request.JobId) ? jobId : request.JobId;
                tool = request.Tool;

                if (tool == "revit_activate_document")
                {
                    var payload = BridgeJson.ReadPayload<ActivateDocumentPayload>(request);
                    var documentTitle = (payload.DocumentTitle ?? string.Empty).Trim();

                    if (documentTitle.Length == 0)
                    {
                        throw new BridgeCommandException("invalid_payload", "Field 'documentTitle' is required for revit_activate_document.");
                    }

                    _activateDocumentHandler.SetJob(jobId, documentTitle, _revitVersion);
                    var raiseResult = _activateDocumentEvent.Raise();

                    if (raiseResult == ExternalEventRequest.Accepted)
                    {
                        asyncDispatched = true;
                    }
                    else
                    {
                        throw new BridgeCommandException(
                            "external_event_rejected",
                            string.Format(CultureInfo.InvariantCulture, "ExternalEvent.Raise() returned '{0}'. Revit may be shutting down.", raiseResult));
                    }
                }
                else
                {
                    response = _commandProcessor.Execute(uiApp, request);
                }
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

            if (!asyncDispatched)
            {
                var resultFile = Path.Combine(OutboxDir, MakeSafeResultName(jobId) + ".result.json");
                File.WriteAllText(resultFile, BridgeJson.Serialize(response), _utf8NoBom);
            }

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

        private static string MakeSafeResultName(string jobId)
        {
            var fallback = "revit-job";
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

    internal sealed class ActivateDocumentEventHandler : IExternalEventHandler
    {
        private readonly string _outboxDir;
        private readonly Encoding _utf8NoBom = new UTF8Encoding(false);

        // Written before Raise(), read in Execute() — both called on Revit's main thread via
        // Idling→Raise and then the subsequent ExternalEvent dispatch, so volatile is sufficient.
        private volatile string _jobId;
        private volatile string _documentTitle;
        private volatile string _revitVersion;

        public ActivateDocumentEventHandler(string outboxDir)
        {
            _outboxDir = outboxDir;
        }

        public void SetJob(string jobId, string documentTitle, string revitVersion)
        {
            _jobId = jobId;
            _documentTitle = documentTitle;
            _revitVersion = revitVersion;
        }

        public void Execute(UIApplication app)
        {
            var jobId = _jobId ?? string.Empty;
            var documentTitle = _documentTitle ?? string.Empty;
            var revitVersion = _revitVersion ?? "unknown";
            JObject response;

            try
            {
                response = ActivateDocument(app, jobId, documentTitle, revitVersion);
            }
            catch (BridgeCommandException ex)
            {
                response = BridgeResultFactory.CreateError(jobId, "revit_activate_document", revitVersion, ex);
            }
            catch (Exception ex)
            {
                response = BridgeResultFactory.CreateError(
                    jobId,
                    "revit_activate_document",
                    revitVersion,
                    new BridgeCommandException("unhandled_exception", ex.Message, null, ex));
            }

            var resultFile = Path.Combine(_outboxDir, MakeSafeResultName(jobId) + ".result.json");
            File.WriteAllText(resultFile, BridgeJson.Serialize(response), _utf8NoBom);
        }

        public string GetName()
        {
            return "ActivateDocumentEventHandler";
        }

        private static JObject ActivateDocument(UIApplication uiApp, string jobId, string documentTitle, string revitVersion)
        {
            // Search all open documents for the requested title.
            Document targetDoc = null;
            foreach (Document doc in uiApp.Application.Documents)
            {
                if (string.Equals(doc.Title, documentTitle, StringComparison.OrdinalIgnoreCase))
                {
                    targetDoc = doc;
                    break;
                }
            }

            if (targetDoc == null)
            {
                throw new BridgeCommandException(
                    "document_not_found",
                    string.Format(CultureInfo.InvariantCulture, "No open document with title '{0}' was found.", documentTitle),
                    new JObject { ["documentTitle"] = documentTitle });
            }

            var previousTitle = uiApp.ActiveUIDocument == null ? null : uiApp.ActiveUIDocument.Document == null ? null : uiApp.ActiveUIDocument.Document.Title;
            var alreadyActive = string.Equals(previousTitle, documentTitle, StringComparison.OrdinalIgnoreCase);

            if (!alreadyActive)
            {
                try
                {
                    if (targetDoc.IsModelInCloud)
                    {
                        uiApp.OpenAndActivateDocument(targetDoc.GetCloudModelPath());
                    }
                    else
                    {
                        uiApp.OpenAndActivateDocument(targetDoc.PathName);
                    }
                }
                catch (InvalidOperationException ex)
                {
                    throw new BridgeCommandException(
                        "activation_failed",
                        string.Format(CultureInfo.InvariantCulture, "Could not activate document '{0}' in the Revit UI.", documentTitle),
                        null,
                        ex);
                }
            }

            var result = BridgeResultFactory.CreateSuccess(jobId, "revit_activate_document", revitVersion);
            result["activated"] = true;
            result["alreadyActive"] = alreadyActive;
            result["documentTitle"] = documentTitle;
            result["previousActiveDocument"] = previousTitle;
            return result;
        }

        private static string MakeSafeResultName(string jobId)
        {
            var fallback = "revit-job";
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
}
