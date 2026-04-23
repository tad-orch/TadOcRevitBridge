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

        private readonly RevitBridgeCommandProcessor _commandProcessor = new RevitBridgeCommandProcessor();
        private readonly Encoding _utf8NoBom = new UTF8Encoding(false);
        private bool _isProcessing;
        private string _revitVersion = "unknown";

        // Pending document activation — set by HandleOpenCloudModel after
        // OpenDocumentFile succeeds. The next Idling tick will activate it
        // in the UI and write the final result file.
        internal static string PendingActivationPath = null;
        internal static string PendingActivationJobId = null;
        internal static string PendingActivationTool = null;
        internal static JObject PendingActivationResult = null;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                _revitVersion = application.ControlledApplication.VersionNumber ?? "unknown";

                Directory.CreateDirectory(InboxDir);
                Directory.CreateDirectory(OutboxDir);
                Directory.CreateDirectory(ArchiveDir);

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

                // Check if there is a document waiting to be activated from
                // a previous Idling tick (set by HandleOpenCloudModel).
                if (PendingActivationPath != null)
                {
                    FlushPendingActivation(uiApp);
                    return;
                }

                ProcessInboxOnce(uiApp);
            }
            finally
            {
                _isProcessing = false;
            }
        }

        private void FlushPendingActivation(UIApplication uiApp)
        {
            var path = PendingActivationPath;
            var jobId = PendingActivationJobId;
            var tool = PendingActivationTool;
            var result = PendingActivationResult;

            // Clear pending state before attempting activation so a crash
            // here does not leave the add-in stuck in a loop.
            PendingActivationPath = null;
            PendingActivationJobId = null;
            PendingActivationTool = null;
            PendingActivationResult = null;

            var activatedInUi = false;
            try
            {
                uiApp.OpenAndActivateDocument(path);
                activatedInUi = true;
            }
            catch
            {
                activatedInUi = false;
            }

            if (result != null)
            {
                result["openedInUi"] = activatedInUi;
                result["activeDocumentChanged"] = activatedInUi;

                // Update the isActive flag inside openedDocument if present.
                var openedDoc = result["openedDocument"] as JObject;
                if (openedDoc != null)
                {
                    openedDoc["isActive"] = activatedInUi;
                }

                var resultFile = Path.Combine(OutboxDir, MakeSafeResultName(jobId) + ".result.json");
                File.WriteAllText(resultFile, BridgeJson.Serialize(result), _utf8NoBom);
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
                // Retry up to 5 times with 200ms gap to handle the race where
                // the writer (bridge or PowerShell) still has the file open.
                string json = null;
                for (var attempt = 0; attempt < 5; attempt++)
                {
                    try
                    {
                        json = File.ReadAllText(firstJob, Encoding.UTF8);
                        break;
                    }
                    catch (IOException)
                    {
                        if (attempt == 4) throw;
                        System.Threading.Thread.Sleep(200);
                    }
                }

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

            // If HandleOpenCloudModel queued a pending activation, the result
            // file will be written by FlushPendingActivation on the next tick
            // instead of here — so we skip writing it now.
            if (PendingActivationPath != null)
            {
                // Store the partial result so FlushPendingActivation can
                // complete and write it after activation.
                PendingActivationResult = response;
                PendingActivationJobId = jobId;
                PendingActivationTool = tool;
            }
            else
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
}