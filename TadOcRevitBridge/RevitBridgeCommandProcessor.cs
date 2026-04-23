using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace TadOcRevitBridge
{
    internal sealed class RevitBridgeCommandProcessor
    {
        private const double MetersToFeet = 3.280839895;

        public JObject Execute(UIApplication uiApp, BridgeRequest request)
        {
            if (uiApp == null)
            {
                throw new BridgeCommandException("unsupported_context", "Revit UI application context is unavailable.");
            }

            if (request == null)
            {
                throw new BridgeCommandException("invalid_request", "Request payload could not be parsed.");
            }

            var tool = (request.Tool ?? string.Empty).Trim();
            if (tool.Length == 0)
            {
                throw new BridgeCommandException("invalid_request", "Request is missing the 'tool' name.");
            }

            switch (tool)
            {
                case "revit_create_wall":
                    return HandleCreateWall(uiApp, request);
                case "revit_session_status":
                    return HandleSessionStatus(uiApp, request);
                case "revit_open_cloud_model":
                    return HandleOpenCloudModel(uiApp, request);
                case "revit_list_3d_views":
                    return HandleList3DViews(uiApp, request);
                case "revit_export_nwc":
                    return HandleExportNwc(uiApp, request);
                default:
                    throw new BridgeCommandException(
                        "unsupported_tool",
                        string.Format(CultureInfo.InvariantCulture, "Unsupported tool '{0}'.", tool),
                        new JObject
                        {
                            ["supportedTools"] = new JArray(SupportedTools.All)
                        });
            }
        }

        private JObject HandleCreateWall(UIApplication uiApp, BridgeRequest request)
        {
            var doc = RequireActiveDocument(uiApp, request.Tool);
            if (doc.IsReadOnly)
            {
                throw new BridgeCommandException("document_read_only", "The active document is read-only and cannot be modified.");
            }

            var payload = BridgeJson.ReadPayload<WallPayload>(request);
            ValidateWallPayload(payload);

            var level = FindLevel(doc, payload.Level);
            var wallType = FindWallType(doc, payload.WallType);
            var line = Line.CreateBound(ToXyz(payload.Start), ToXyz(payload.End));

            using (var tx = new Transaction(doc, "TAD Create Wall"))
            {
                tx.Start();
                Wall.Create(doc, line, wallType.Id, level.Id, payload.Height * MetersToFeet, 0.0, false, false);
                tx.Commit();
            }

            var result = BridgeResultFactory.CreateSuccess(request.JobId, request.Tool, GetRevitVersion(uiApp));
            result["documentTitle"] = doc.Title;
            result["level"] = level.Name;
            result["wallType"] = wallType.Name;
            return result;
        }

        private JObject HandleSessionStatus(UIApplication uiApp, BridgeRequest request)
        {
            var result = BridgeResultFactory.CreateSuccess(request.JobId, request.Tool, GetRevitVersion(uiApp));
            result["revitRunning"] = true;
            result["activeDocument"] = BuildDocumentSummary(uiApp.ActiveUIDocument == null ? null : uiApp.ActiveUIDocument.Document, true);
            return result;
        }

        private JObject HandleOpenCloudModel(UIApplication uiApp, BridgeRequest request)
        {
            var payload = BridgeJson.ReadPayload<OpenCloudModelPayload>(request);
            var normalizedRegion = NormalizeCloudRegion(payload.Region);
            var projectGuid = ParseGuid(payload.ProjectGuid, "projectGuid");
            var modelGuid = ParseGuid(payload.ModelGuid, "modelGuid");

            if (payload.OpenInUi)
            {
                throw new BridgeCommandException(
                    "unsupported_context",
                    "openInUi=true is not supported by the current idling queue path. Use openInUi=false or activate the opened document manually in Revit after it is opened.",
                    new JObject
                    {
                        ["openInUi"] = true
                    });
            }

            var openOptions = BuildOpenOptions(payload);
            var callback = BuildOpenFromCloudCallback(payload.CloudOpenConflictPolicy);
            Document openedDocument;

                    var result = BridgeResultFactory.CreateSuccess(request.JobId, request.Tool, GetRevitVersion(uiApp));
                    result["openedInUi"] = false;
                    result["activeDocumentChanged"] = false;
                    result["openedDocument"] = BuildDocumentSummary(openedDocument, false);
                    return result;
                }
            }
            catch (BridgeCommandException)
            {
                throw;
            }
            catch (RevitServerUnauthenticatedUserException ex)
            {
                throw new BridgeCommandException("unauthenticated_user", "A signed-in Autodesk user is required in this Revit session before a cloud model can be opened.", null, ex);
            }
            catch (RevitServerUnauthorizedException ex)
            {
                throw new BridgeCommandException("unauthorized_access", "The signed-in Autodesk user does not have access to the requested cloud model.", null, ex);
            }
            catch (RevitServerCommunicationException ex)
            {
                throw new BridgeCommandException("communication_failure", "Revit could not reach Autodesk cloud services while resolving or opening the cloud model.", null, ex);
            }
            catch (CentralModelException ex)
            {
                throw new BridgeCommandException(
                    "invalid_identifiers",
                    "The region, project GUID, or model GUID could not be resolved to a valid Revit cloud model.",
                    new JObject
                    {
                        ["region"] = normalizedRegion,
                        ["projectGuid"] = projectGuid.ToString(),
                        ["modelGuid"] = modelGuid.ToString()
                    },
                    ex);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException ex)
            {
                throw new BridgeCommandException("open_cancelled", "Opening the cloud model was cancelled by Revit.", null, ex);
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                throw new BridgeCommandException("unsupported_context", "Revit rejected the cloud-model open request in the current application context.", null, ex);
            }

            if (openedDocument == null)
            {
                throw new BridgeCommandException("open_failed", "Revit did not return an opened document for the cloud-model request.");
            }

            var result = BridgeResultFactory.CreateSuccess(request.JobId, request.Tool, GetRevitVersion(uiApp));
            result["openedInUi"] = false;
            result["activeDocumentChanged"] = false;
            result["openedDocument"] = BuildDocumentSummary(openedDocument, false);
            return result;
        }

        private JObject HandleList3DViews(UIApplication uiApp, BridgeRequest request)
        {
            var doc = RequireActiveDocument(uiApp, request.Tool);
            var payload = BridgeJson.ReadPayload<List3DViewsPayload>(request);
            var exporterAvailable = OptionalFunctionalityUtils.IsNavisworksExporterAvailable();
            var excludeTemplates = payload.ExcludeTemplates ?? payload.OnlyExportable;

            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Select(view => new
                {
                    View = view,
                    CanExport = CanExportNavisworks(view, exporterAvailable)
                })
                .Where(item => !excludeTemplates || !item.View.IsTemplate)
                .Where(item => !payload.OnlyExportable || item.CanExport)
                .OrderBy(item => item.View.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var result = BridgeResultFactory.CreateSuccess(request.JobId, request.Tool, GetRevitVersion(uiApp));
            result["documentTitle"] = doc.Title;
            result["navisworksExporterAvailable"] = exporterAvailable;
            result["views"] = new JArray(
                views.Select(item => new JObject
                {
                    ["id"] = ElementIdToString(item.View.Id),
                    ["name"] = item.View.Name,
                    ["isTemplate"] = item.View.IsTemplate,
                    ["canExport"] = item.CanExport,
                    ["isPerspective"] = item.View.IsPerspective
                }));
            return result;
        }

        private JObject HandleExportNwc(UIApplication uiApp, BridgeRequest request)
        {
            var doc = RequireActiveDocument(uiApp, request.Tool);
            if (!OptionalFunctionalityUtils.IsNavisworksExporterAvailable())
            {
                throw new BridgeCommandException("navisworks_exporter_unavailable", "A compatible Navisworks exporter is not registered in the current Revit session.");
            }

            var payload = BridgeJson.ReadPayload<ExportNwcPayload>(request);
            var exportScope = NormalizeExportScope(payload.ExportScope);
            var baseOutputPath = NormalizeOutputPath(payload.OutputPath);

            var result = BridgeResultFactory.CreateSuccess(request.JobId, request.Tool, GetRevitVersion(uiApp));
            result["documentTitle"] = doc.Title;

            if (exportScope == "model")
            {
                ExportDocumentAsNwc(doc, baseOutputPath, null);
                result["exportScope"] = "model";
                result["outputPath"] = baseOutputPath;
                result["exportedViews"] = new JArray();
                return result;
            }

            var requestedViews = ResolveRequestedViews(doc, payload.ViewNames);
            var usedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var exportedViews = new JArray();

            for (var index = 0; index < requestedViews.Count; index++)
            {
                var view = requestedViews[index];
                var outputPath = requestedViews.Count == 1
                    ? baseOutputPath
                    : BuildPerViewOutputPath(baseOutputPath, view.Name, usedPaths);

                ExportDocumentAsNwc(doc, outputPath, view.Id);

                exportedViews.Add(new JObject
                {
                    ["id"] = ElementIdToString(view.Id),
                    ["name"] = view.Name,
                    ["outputPath"] = outputPath
                });
            }

            result["exportScope"] = "selected_views";
            result["outputPath"] = requestedViews.Count == 1 ? baseOutputPath : null;
            if (requestedViews.Count > 1)
            {
                result["outputPaths"] = new JArray(exportedViews.Select(item => item["outputPath"]));
            }

            result["exportedViews"] = exportedViews;
            return result;
        }

        private static JObject BuildDocumentSummary(Document doc, bool isActiveDocument)
        {
            var summary = new JObject
            {
                ["isOpen"] = doc != null,
                ["isActive"] = doc != null && isActiveDocument,
                ["title"] = doc == null ? null : doc.Title,
                ["isCloudModel"] = doc != null && doc.IsModelInCloud,
                ["projectGuid"] = null,
                ["modelGuid"] = null,
                ["region"] = null
            };

            if (doc == null || !doc.IsModelInCloud)
            {
                return summary;
            }

            try
            {
                var cloudPath = doc.GetCloudModelPath();
                summary["projectGuid"] = cloudPath.GetProjectGUID().ToString();
                summary["modelGuid"] = cloudPath.GetModelGUID().ToString();
                summary["region"] = cloudPath.Region;
            }
            catch
            {
                summary["projectGuid"] = null;
                summary["modelGuid"] = null;
                summary["region"] = null;
            }

            return summary;
        }

        private static Document RequireActiveDocument(UIApplication uiApp, string tool)
        {
            var document = uiApp.ActiveUIDocument == null ? null : uiApp.ActiveUIDocument.Document;
            if (document == null)
            {
                throw new BridgeCommandException(
                    "no_active_document",
                    string.Format(CultureInfo.InvariantCulture, "Tool '{0}' requires an active Revit document.", tool));
            }

            return document;
        }

        private static string GetRevitVersion(UIApplication uiApp)
        {
            return uiApp.Application.VersionNumber ?? "unknown";
        }

        private static XYZ ToXyz(Point3 point)
        {
            if (point == null)
            {
                throw new BridgeCommandException("invalid_payload", "Wall payload is missing a point definition.");
            }

            return new XYZ(point.X * MetersToFeet, point.Y * MetersToFeet, point.Z * MetersToFeet);
        }

        private static void ValidateWallPayload(WallPayload payload)
        {
            if (payload == null)
            {
                throw new BridgeCommandException("invalid_payload", "Wall payload is missing.");
            }

            if (payload.Start == null || payload.End == null)
            {
                throw new BridgeCommandException("invalid_payload", "Wall payload requires both 'start' and 'end' points.");
            }

            if (payload.Height <= 0.0)
            {
                throw new BridgeCommandException("invalid_payload", "Wall payload requires a positive 'height'.");
            }
        }

        private static Level FindLevel(Document doc, string requestedLevelName)
        {
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(level => level.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var level = string.IsNullOrWhiteSpace(requestedLevelName)
                ? levels.FirstOrDefault()
                : levels.FirstOrDefault(item => item.Name.Equals(requestedLevelName, StringComparison.OrdinalIgnoreCase)) ?? levels.FirstOrDefault();

            if (level == null)
            {
                throw new BridgeCommandException("level_not_found", "No level was found in the active document.");
            }

            return level;
        }

        private static WallType FindWallType(Document doc, string requestedWallTypeName)
        {
            var wallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .OrderBy(wallType => wallType.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (wallTypes.Count == 0)
            {
                throw new BridgeCommandException("wall_type_not_found", "No wall type was found in the active document.");
            }

            if (string.IsNullOrWhiteSpace(requestedWallTypeName))
            {
                return wallTypes[0];
            }

            return wallTypes.FirstOrDefault(item => item.Name.Equals(requestedWallTypeName, StringComparison.OrdinalIgnoreCase)) ?? wallTypes[0];
        }

        private static Guid ParseGuid(string value, string fieldName)
        {
            Guid guid;
            if (!Guid.TryParse(value, out guid))
            {
                throw new BridgeCommandException(
                    "invalid_identifiers",
                    string.Format(CultureInfo.InvariantCulture, "Field '{0}' must be a valid GUID.", fieldName),
                    new JObject
                    {
                        ["field"] = fieldName,
                        ["value"] = value
                    });
            }

            return guid;
        }

        private static string NormalizeCloudRegion(string region)
        {
            var value = (region ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                throw new BridgeCommandException("wrong_region", "Cloud model region is required.");
            }

            if (value.Equals("EU", StringComparison.OrdinalIgnoreCase))
            {
                value = "EMEA";
            }

            var knownRegions = GetKnownCloudRegions();
            if (knownRegions.Count > 0 && !knownRegions.Contains(value))
            {
                throw new BridgeCommandException(
                    "wrong_region",
                    string.Format(CultureInfo.InvariantCulture, "Unsupported cloud region '{0}'.", value),
                    new JObject
                    {
                        ["knownRegions"] = new JArray(knownRegions.OrderBy(item => item, StringComparer.OrdinalIgnoreCase))
                    });
            }

            return value;
        }

        private static HashSet<string> GetKnownCloudRegions()
        {
            var regions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var method = typeof(ModelPathUtils).GetMethod("GetAllCloudRegions", BindingFlags.Public | BindingFlags.Static, null, Type.EmptyTypes, null);
                if (method != null)
                {
                    var value = method.Invoke(null, null) as IEnumerable<string>;
                    if (value != null)
                    {
                        foreach (var region in value)
                        {
                            if (!string.IsNullOrWhiteSpace(region))
                            {
                                regions.Add(region.Trim());
                            }
                        }
                    }
                }
            }
            catch
            {
            }

            if (regions.Count == 0)
            {
                regions.Add("US");
                regions.Add("EMEA");
            }

            return regions;
        }

        private static OpenOptions BuildOpenOptions(OpenCloudModelPayload payload)
        {
            var options = new OpenOptions
            {
                Audit = payload.Audit
            };

            var mode = ((payload.Worksets == null ? null : payload.Worksets.Mode) ?? "default").Trim().ToLowerInvariant();
            switch (mode)
            {
                case "":
                case "default":
                case "open_all":
                    return options;
                case "close_all":
                    options.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));
                    return options;
                case "open_last_viewed":
                    options.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.OpenLastViewed));
                    return options;
                default:
                    throw new BridgeCommandException(
                        "invalid_workset_mode",
                        string.Format(CultureInfo.InvariantCulture, "Unsupported workset mode '{0}'.", mode));
            }
        }

        private static IOpenFromCloudCallback BuildOpenFromCloudCallback(string policy)
        {
            var normalized = (policy ?? "use_default").Trim().ToLowerInvariant();
            switch (normalized)
            {
                case "":
                case "use_default":
                case "discard_local_changes_and_open_latest_version":
                    return new CloudOpenConflictCallback(OpenConflictResult.DiscardLocalChangesAndOpenLatestVersion);
                case "keep_local_changes":
                    return new CloudOpenConflictCallback(OpenConflictResult.KeepLocalChanges);
                case "detach_from_central":
                    return new CloudOpenConflictCallback(OpenConflictResult.DetachFromCentral);
                case "cancel":
                    return new CloudOpenConflictCallback(OpenConflictResult.Cancel);
                default:
                    throw new BridgeCommandException(
                        "invalid_conflict_policy",
                        string.Format(CultureInfo.InvariantCulture, "Unsupported cloudOpenConflictPolicy '{0}'.", policy));
            }
        }

        private static bool CanExportNavisworks(View3D view, bool exporterAvailable)
        {
            return exporterAvailable && view != null && !view.IsTemplate;
        }

        private static string ElementIdToString(ElementId id)
        {
            if (id == null)
            {
                return null;
            }

#if REVIT2024
            return id.IntegerValue.ToString(CultureInfo.InvariantCulture);
#else
            return id.Value.ToString(CultureInfo.InvariantCulture);
#endif
        }

        private static string NormalizeExportScope(string exportScope)
        {
            var normalized = (exportScope ?? "selected_views").Trim().ToLowerInvariant();
            switch (normalized)
            {
                case "":
                case "selected_views":
                    return "selected_views";
                case "model":
                    return "model";
                default:
                    throw new BridgeCommandException(
                        "invalid_export_scope",
                        string.Format(CultureInfo.InvariantCulture, "Unsupported exportScope '{0}'.", exportScope));
            }
        }

        private static string NormalizeOutputPath(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new BridgeCommandException("invalid_output_path", "An absolute outputPath is required for Navisworks export.");
            }

            var trimmed = outputPath.Trim();
            if (!Path.IsPathRooted(trimmed))
            {
                throw new BridgeCommandException("invalid_output_path", "outputPath must be an absolute path.");
            }

            if (string.IsNullOrWhiteSpace(Path.GetExtension(trimmed)))
            {
                trimmed += ".nwc";
            }
            else if (!Path.GetExtension(trimmed).Equals(".nwc", StringComparison.OrdinalIgnoreCase))
            {
                throw new BridgeCommandException("invalid_output_path", "outputPath must end with the .nwc extension.");
            }

            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(trimmed);
            }
            catch (Exception ex)
            {
                throw new BridgeCommandException("invalid_output_path", "outputPath is not a valid filesystem path.", null, ex);
            }

            var directory = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new BridgeCommandException("invalid_output_path", "outputPath must include a target directory.");
            }

            try
            {
                Directory.CreateDirectory(directory);
            }
            catch (Exception ex)
            {
                throw new BridgeCommandException("invalid_output_path", "Revit add-in could not create or access the output directory.", null, ex);
            }

            return fullPath;
        }

        private static List<View3D> ResolveRequestedViews(Document doc, IList<string> viewNames)
        {
            if (viewNames == null || viewNames.Count == 0)
            {
                throw new BridgeCommandException("requested_view_not_found", "At least one 3D view name is required for exportScope='selected_views'.");
            }

            var allViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(view => !view.IsTemplate)
                .ToList();

            var resolved = new List<View3D>();
            foreach (var requestedViewName in viewNames)
            {
                var viewName = (requestedViewName ?? string.Empty).Trim();
                if (viewName.Length == 0)
                {
                    throw new BridgeCommandException("requested_view_not_found", "View names for export cannot be blank.");
                }

                var match = allViews.FirstOrDefault(view => view.Name.Equals(viewName, StringComparison.OrdinalIgnoreCase));
                if (match == null)
                {
                    throw new BridgeCommandException(
                        "requested_view_not_found",
                        string.Format(CultureInfo.InvariantCulture, "Requested 3D view '{0}' was not found in the active document.", viewName),
                        new JObject
                        {
                            ["viewName"] = viewName
                        });
                }

                resolved.Add(match);
            }

            return resolved;
        }

        private static string BuildPerViewOutputPath(string baseOutputPath, string viewName, ISet<string> usedPaths)
        {
            var directory = Path.GetDirectoryName(baseOutputPath);
            var baseName = Path.GetFileNameWithoutExtension(baseOutputPath);
            var extension = Path.GetExtension(baseOutputPath);
            var suffix = SanitizePathSegment(viewName);
            var candidate = Path.Combine(directory, baseName + "-" + suffix + extension);
            var counter = 2;

            while (usedPaths.Contains(candidate))
            {
                candidate = Path.Combine(directory, baseName + "-" + suffix + "-" + counter.ToString(CultureInfo.InvariantCulture) + extension);
                counter++;
            }

            usedPaths.Add(candidate);
            return candidate;
        }

        private static string SanitizePathSegment(string value)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitized = new string((value ?? string.Empty)
                .Trim()
                .Select(ch => invalidChars.Contains(ch) ? '_' : ch)
                .ToArray());

            return string.IsNullOrWhiteSpace(sanitized) ? "view" : sanitized;
        }

        private static void ExportDocumentAsNwc(Document doc, string outputPath, ElementId viewId)
        {
            var outputDirectory = Path.GetDirectoryName(outputPath);
            var outputName = Path.GetFileNameWithoutExtension(outputPath);
            var exportStartedUtc = DateTime.UtcNow;
            var options = new NavisworksExportOptions();

            if (viewId != null)
            {
                options.ExportScope = NavisworksExportScope.View;
                options.ViewId = viewId;
            }

            try
            {
                doc.Export(outputDirectory, outputName, options);
            }
            catch (Exception ex)
            {
                throw new BridgeCommandException(
                    "export_failed",
                    string.Format(CultureInfo.InvariantCulture, "Navisworks export failed for '{0}'.", outputPath),
                    null,
                    ex);
            }

            if (!File.Exists(outputPath))
            {
                throw new BridgeCommandException(
                    "export_failed",
                    string.Format(CultureInfo.InvariantCulture, "Navisworks export did not produce '{0}'.", outputPath));
            }

            var writeTimeUtc = File.GetLastWriteTimeUtc(outputPath);
            if (writeTimeUtc < exportStartedUtc.AddSeconds(-1))
            {
                throw new BridgeCommandException(
                    "export_failed",
                    string.Format(CultureInfo.InvariantCulture, "Navisworks export did not update '{0}'.", outputPath));
            }
        }
    }

    internal sealed class CloudOpenConflictCallback : IOpenFromCloudCallback
    {
        private readonly OpenConflictResult _result;

        public CloudOpenConflictCallback(OpenConflictResult result)
        {
            _result = result;
        }

        public OpenConflictResult OnOpenConflict(OpenConflictScenario scenario)
        {
            return _result;
        }
    }
}