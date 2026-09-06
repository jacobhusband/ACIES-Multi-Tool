using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Reflection;
using System.Text;

// Experimental harness only: not included in the installed command bundle.
public static class HeadlessEmbeddingProbe
{
    [CommandMethod("ACIESOLEPROBE")]
    public static void Probe()
    {
        var report = new StringBuilder();
        string directory = Environment.GetEnvironmentVariable("ACIES_EMBED_PROBE_DIR");
        using (var ole = new Ole2Frame())
        {
            foreach (string property in new[] { "OleObject", "Type", "IsLinked", "UserType" })
            {
                try
                {
                    object value = typeof(Ole2Frame).GetProperty(property).GetValue(ole, null);
                    report.AppendLine(property + ": " + (value == null ? "<null>" : value.GetType().FullName + " = " + value));
                    if (property == "OleObject" && value != null)
                    {
                        foreach (PropertyInfo info in value.GetType().GetProperties())
                            report.AppendLine("  " + info.Name + ": " + info.PropertyType + " writable=" + info.CanWrite);
                    }
                }
                catch (System.Exception ex) { report.AppendLine(property + ": " + ex); }
            }
        }
        File.WriteAllText(Path.Combine(directory, "ole-probe.txt"), report.ToString());
    }
}
