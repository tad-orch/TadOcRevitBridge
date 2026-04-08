using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Threading.Tasks;

namespace TadOcRevitBridge
{
    public class App : IExternalApplication
    {
        
        private const string InboxDir = @"D:\TAD\revit-bridge\inbox";
        private const string OutboxDir = @"D:\TAD\revit-bridge\outbox";
        private const string ArchiveDir = @"D:\TAD\revit-bridge\archive";
        private const string AliveFile = @"D:\TAD\revit-bridge\outbox\revit-addin-alive.json";

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                Directory.CreateDirectory(InboxDir);
                Directory.CreateDirectory(OutboxDir);
                Directory.CreateDirectory(ArchiveDir);

                var json = BuildAliveJson();
                File.WriteAllText(AliveFile, json, Encoding.UTF8);

                application.Idling += OnIdling;

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }

        private void ProcessInboxOnce(Document doc)
        {
            var firstJob = Directory.GetFiles(InboxDir, "*.json")
                .OrderBy(f => f)
                .FirstOrDefault();

            if (string.IsNullOrWhiteSpace(firstJob))
                return;

            var json = File.ReadAllText(firstJob, Encoding.UTF8);
            var serializer = new JavaScriptSerializer();
            var job = serializer.Deserialize<WallJob>(json);

            if (job == null || job.tool != "revit_create_wall" || job.payload == null)
                return;

            var resultFile = Path.Combine(OutboxDir, $"{job.jobId}.result.json");

            try
            {
                CreateWall(doc, job.payload);

                var okJson =
        $@"{{
  ""ok"": true,
  ""source"": ""revit-addin"",
  ""status"": ""completed"",
  ""jobId"": ""{Escape(job.jobId)}"",
  ""tool"": ""revit_create_wall"",
  ""timeUtc"": ""{DateTime.UtcNow:o}""
}}";

                File.WriteAllText(resultFile, okJson, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                var errorJson =
        $@"{{
  ""ok"": false,
  ""source"": ""revit-addin"",
  ""status"": ""failed"",
  ""jobId"": ""{Escape(job.jobId)}"",
  ""tool"": ""revit_create_wall"",
  ""error"": ""{Escape(ex.Message)}"",
  ""timeUtc"": ""{DateTime.UtcNow:o}""
}}";

                File.WriteAllText(resultFile, errorJson, Encoding.UTF8);
            }

            var archived = Path.Combine(ArchiveDir, Path.GetFileName(firstJob));
            if (File.Exists(archived))
                File.Delete(archived);

            File.Move(firstJob, archived);
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            application.Idling -= OnIdling;
            return Result.Succeeded;
        }

        private string BuildAliveJson()
        {
            var machine = Environment.MachineName;
            var user = Environment.UserName;
            var timestamp = DateTime.UtcNow.ToString("o");

            return
$@"{{
  ""ok"": true,
  ""source"": ""revit-addin"",
  ""status"": ""alive"",
  ""machine"": ""{Escape(machine)}"",
  ""user"": ""{Escape(user)}"",
  ""revitVersion"": ""2024"",
  ""timeUtc"": ""{timestamp}""
}}";
        }

        private string Escape(string value)
        {
            return (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"");
        }

        private bool _isProcessing = false;

        private void OnIdling(object sender, IdlingEventArgs e)
        {
            if (_isProcessing) return;

            try
            {
                _isProcessing = true;

                var uiApp = sender as UIApplication;
                if (uiApp == null) return;

                var doc = uiApp.ActiveUIDocument?.Document;
                if (doc == null || doc.IsReadOnly) return;

                ProcessInboxOnce(doc);
            }
            finally
            {
                _isProcessing = false;
            }
        }

        private void CreateWall(Document doc, WallPayload payload)
        {
            var level = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .FirstOrDefault(l => l.Name.Equals(payload.level, StringComparison.OrdinalIgnoreCase))
                ?? new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .FirstOrDefault();

            if (level == null)
                throw new Exception("No level found.");

            var wallType = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault();

            if (wallType == null)
                throw new Exception("No wall type found.");

            double metersToFeet = 3.280839895;

            var start = new XYZ(
                payload.start.x * metersToFeet,
                payload.start.y * metersToFeet,
                payload.start.z * metersToFeet);

            var end = new XYZ(
                payload.end.x * metersToFeet,
                payload.end.y * metersToFeet,
                payload.end.z * metersToFeet);

            var line = Line.CreateBound(start, end);
            var heightFeet = payload.height * metersToFeet;

            using (var tx = new Transaction(doc, "TAD Create Wall"))
            {
                tx.Start();

                Wall.Create(doc, line, wallType.Id, level.Id, heightFeet, 0, false, false);

                tx.Commit();
            }
        }
    }
}

public class WallJob
{
    public string jobId { get; set; }
    public string tool { get; set; }
    public string createdAt { get; set; }
    public string status { get; set; }
    public WallPayload payload { get; set; }
}

public class WallPayload
{
    public Point3 start { get; set; }
    public Point3 end { get; set; }
    public double height { get; set; }
    public string level { get; set; }
    public string wallType { get; set; }
}

public class Point3
{
    public double x { get; set; }
    public double y { get; set; }
    public double z { get; set; }
}
