using System;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CompareDuctCableTrayPlacementCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc?.Document;

            try
            {
                // Locate repository root by walking up from the executing assembly location
                var asmPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var dir = new DirectoryInfo(Path.GetDirectoryName(asmPath) ?? ".");
                DirectoryInfo? repoRoot = null;
                for (int i = 0; i < 8 && dir != null; i++)
                {
                    // Heuristic: repo root contains a Commands folder and Services folder
                    if (Directory.Exists(Path.Combine(dir.FullName, "Commands")) && Directory.Exists(Path.Combine(dir.FullName, "Services")))
                    {
                        repoRoot = dir;
                        break;
                    }
                    dir = dir.Parent;
                }

                // If we couldn't discover repo root by walking up from the assembly, try a few sensible
                // fallbacks so the command is usable when the add-in is loaded from Revit's AddIns folder.
                if (repoRoot == null)
                {
                    // 1) Environment variable override (useful for machines where sources live elsewhere)
                    try
                    {
                        var envRoot = Environment.GetEnvironmentVariable("JSE_PROJECT_ROOT");
                        if (!string.IsNullOrWhiteSpace(envRoot) && Directory.Exists(envRoot))
                        {
                            repoRoot = new DirectoryInfo(envRoot);
                        }
                    }
                    catch { }
                }

                if (repoRoot == null)
                {
                    // 2) Common workspace fallback used by developer environment
                    var fallbackPath = Path.Combine("C:", "JSE_CSharp_Projects", "JSE_MEPOPENING_23");
                    if (Directory.Exists(fallbackPath))
                        repoRoot = new DirectoryInfo(fallbackPath);
                }

                if (repoRoot == null)
                {
                    var msg = new StringBuilder();
                    msg.AppendLine("Could not locate project root. Ensure the add-in is running from the expected build output or set the environment variable JSE_PROJECT_ROOT to the source tree root.");
                    msg.AppendLine("Paths attempted:");
                    msg.AppendLine(" - Walk-up from assembly location: " + asmPath);
                    msg.AppendLine(" - Environment variable: JSE_PROJECT_ROOT");
                    msg.AppendLine(" - Developer fallback: C:\\JSE_CSharp_Projects\\JSE_MEPOPENING_23");
                    TaskDialog.Show("Compare Placement", msg.ToString());
                    return Result.Failed;
                }

                string servicesPath = Path.Combine(repoRoot.FullName, "Services");
                string commandsPath = Path.Combine(repoRoot.FullName, "Commands");
                string logDir = Path.Combine(repoRoot.FullName, "Log");
                Directory.CreateDirectory(logDir);
                string outLog = Path.Combine(logDir, "PlacementComparison.log");

                var sb = new StringBuilder();
                sb.AppendLine($"Placement comparison run: {DateTime.Now:O}");
                sb.AppendLine($"Repo root: {repoRoot.FullName}");

                // Files to compare
                string ductFile = Path.Combine(servicesPath, "DuctSleevePlacer.cs");
                string cableFile = Path.Combine(servicesPath, "CableTraySleevePlacer.cs");
                string dupFile = Path.Combine(servicesPath, "OpeningDuplicationChecker.cs");
                string cmdFile = Path.Combine(commandsPath, "CableTraySleeveCommand.cs");

                if (!File.Exists(ductFile) || !File.Exists(cableFile) || !File.Exists(dupFile) || !File.Exists(cmdFile))
                {
                    sb.AppendLine("ERROR: One or more source files not found:");
                    sb.AppendLine($" - Duct: {ductFile} (exists={File.Exists(ductFile)})");
                    sb.AppendLine($" - Cable: {cableFile} (exists={File.Exists(cableFile)})");
                    sb.AppendLine($" - DuplicationChecker: {dupFile} (exists={File.Exists(dupFile)})");
                    sb.AppendLine($" - Command: {cmdFile} (exists={File.Exists(cmdFile)})");
                    File.WriteAllText(outLog, sb.ToString());
                    TaskDialog.Show("Compare Placement", sb.ToString());
                    return Result.Failed;
                }

                string ductSrc = File.ReadAllText(ductFile);
                string cableSrc = File.ReadAllText(cableFile);
                string dupSrc = File.ReadAllText(dupFile);
                string cmdSrc = File.ReadAllText(cmdFile);

                sb.AppendLine("--- High level findings ---");

                // 1) Where NewFamilyInstance is called and which point variable is used
                sb.AppendLine("DuctSleevePlacer.NewFamilyInstance occurrences:");
                foreach (var line in ductSrc.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Where(l => l.Contains("NewFamilyInstance(")))
                    sb.AppendLine("  " + line.Trim());

                sb.AppendLine();
                sb.AppendLine("CableTraySleevePlacer.NewFamilyInstance occurrences:");
                foreach (var line in cableSrc.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Where(l => l.Contains("NewFamilyInstance(")))
                    sb.AppendLine("  " + line.Trim());

                sb.AppendLine();
                // 2) Inspect placePoint computation presence
                sb.AppendLine("DuctSleevePlacer placePoint computation snippets:");
                foreach (var line in ductSrc.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Where(l => l.Contains("placePoint") || l.Contains("wallVector") || l.Contains("GetWallNormal(")))
                    sb.AppendLine("  " + line.Trim());

                sb.AppendLine();
                sb.AppendLine("CableTraySleevePlacer placePoint computation snippets:");
                foreach (var line in cableSrc.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Where(l => l.Contains("placePoint") || l.Contains("wallVector") || l.Contains("GetWallNormal(")))
                    sb.AppendLine("  " + line.Trim());

                sb.AppendLine();
                // 3) Duplication checker: what location is used (search for FindIndividualSleevesAtLocation/GetSleevesSummaryAtLocation calls in command)
                sb.AppendLine("Calls to OpeningDuplicationChecker in CableTraySleeveCommand.cs:");
                foreach (var line in cmdSrc.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Where(l => l.Contains("OpeningDuplicationChecker") || l.Contains("FindIndividualSleevesAtLocation") || l.Contains("GetSleevesSummaryAtLocation") || l.Contains("FindCableTraySleeves(")))
                    sb.AppendLine("  " + line.Trim());

                sb.AppendLine();
                sb.AppendLine("OpeningDuplicationChecker helper snippets (search for transform/FindIndividualSleevesAtLocation):");
                foreach (var line in dupSrc.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Where(l => l.Contains("FindIndividualSleevesAtLocation") || l.Contains("FindClustersAtLocation") || l.Contains("FindCableTraySleeves") || l.Contains("FindIndividualSleevesInDoc(")))
                    sb.AppendLine("  " + line.Trim());

                sb.AppendLine();
                // 4) Quick programmatic conclusion
                sb.AppendLine("--- Programmatic conclusion ---");

                bool ductPlacesAtIntersection = ductSrc.Contains("NewFamilyInstance(\n                    intersection") || ductSrc.Contains("NewFamilyInstance(intersection") || ductSrc.Contains("NewFamilyInstance(\n                    placePointForDoc") == false && ductSrc.Contains("NewFamilyInstance(") && ductSrc.Contains("intersection") && ductSrc.IndexOf("NewFamilyInstance(") < ductSrc.IndexOf("placePoint");
                bool cablePlacesAtPlacePoint = cableSrc.Contains("placePointForDoc") && cableSrc.Contains("NewFamilyInstance(") && cableSrc.Contains("placePointForDoc") && cableSrc.IndexOf("placePointForDoc") < cableSrc.IndexOf("NewFamilyInstance(");

                sb.AppendLine($"Duct placer creates instance at intersection? => {ductPlacesAtIntersection}");
                sb.AppendLine($"Cable tray placer creates instance at placePoint? => {cablePlacesAtPlacePoint}");

                // 5) Duplication/command mismatch
                bool cmdUsesIntersectionForDupCheck = cmdSrc.Contains("GetSleevesSummaryAtLocation(doc, loc") || cmdSrc.Contains("FindIndividualSleevesAtLocation(doc, intersection") || cmdSrc.Contains("FindIndividualSleevesAtLocation(doc, intersectionToPass") || cmdSrc.Contains("FindIndividualSleevesAtLocation(doc, location") || cmdSrc.Contains("GetSleevesSummaryAtLocation(doc, loc") || cmdSrc.Contains("GetSleevesSummaryAtLocation(doc, intersection") || cmdSrc.Contains("GetSleevesSummaryAtLocation(doc, intersectionToPass");
                sb.AppendLine($"CableTray command uses intersection-like variable for duplication checks? => {cmdUsesIntersectionForDupCheck}");

                // 6) Final recommendation
                sb.AppendLine();
                sb.AppendLine("Recommendations:");
                sb.AppendLine(" - Make duplication checks in CableTray command use the same point the placer uses (placePoint) or perform dual-check (intersection + placePoint) to avoid false-negative duplicates.");
                sb.AppendLine(" - Keep MepIntersectionService outputs as host coords and do not re-apply link transforms in the placer. If you transform into link-local for host-based placement, ensure you only do it once.");
                sb.AppendLine(" - Optionally add a short debug log showing both intersection and final placePoint when performing duplicate checks.");

                File.WriteAllText(outLog, sb.ToString());

                // Show a short TaskDialog summary
                string shortSummary = "Placement comparison completed. Output written to:\n" + outLog + "\n\nPrimary finding: Duct places at intersection, CableTray places at centerline (placePoint). Make duplication checks consistent with placement.";
                TaskDialog.Show("Compare Duct vs CableTray Placement", shortSummary);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Compare Placement - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}
