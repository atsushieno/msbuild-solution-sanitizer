using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Build.Construction;

namespace MSBuildSolutionSanitizer {

	public class MainDriver 
	{
		public static int Main (string [] args)
		{
			var proc = new ExistingSolutionProcessor ();
			string baseDir = null, input = null, output = null;
			foreach (var arg_ in args) {
				var arg = (arg_.StartsWith ("--")) ? arg_.Substring (1) : arg_;
				if (arg == "-help")
					return ShowHelp ();
				if (arg.StartsWith ("-basedir:"))
					baseDir = arg.Substring ("-basedir:".Length);
				else if (input == null)
					input = arg;
				else
					output = arg;
			}
			if (input == null)
				return ShowHelp ();
			if (input.EndsWith (".sln", StringComparison.OrdinalIgnoreCase))
				proc.LoadSolution (input);
			else
				proc.LoadXml (input, baseDir);

			proc.ReorderProjects ();

			proc.Dump ();

			if (output != null) {
				if (output.EndsWith (".sln", StringComparison.OrdinalIgnoreCase))
					proc.SaveSolution (output);
				else
					proc.SaveXml (output);
			}
			return 0;
		}

		static int ShowHelp ()
		{
			Console.Error.WriteLine (@"Usage: MSBuildSolutionSanitizer [options] input [output]
-help			show this help
-basedir:[path]		Path to resolve paths to project
			(default: the solution directory)
input			Solution (.sln) or definition (.xml) to read
output			Sanitized solution (.sln) or definition (.xml) to write
");
			return 1;
		}
	}

	public class ExistingSolutionProcessor
	{
		public void ReorderProjects ()
		{
			projects = projects.OrderBy (p => p.Name, StringComparer.OrdinalIgnoreCase).ToList ();
			solution_configuration_platforms = solution_configuration_platforms.OrderBy (s => s, StringComparer.OrdinalIgnoreCase).ToList ();
			nested_projects = NestedProjects.OrderBy (g => g.Key, StringComparer.OrdinalIgnoreCase).SelectMany (g => g.OrderBy (n => n.Item, StringComparer.OrdinalIgnoreCase)).ToList ();
			project_configuration_platforms = project_configuration_platforms.OrderBy (c => projects.IndexOf (projects.First (p => p.ProjectGuid == c.ProjectGuid)))
											 .ThenBy (c => c.SolutionConfigurationName, StringComparer.OrdinalIgnoreCase)
											 .ThenBy (c => c.Property, StringComparer.OrdinalIgnoreCase)
											 .ToList ();
		}

		public void SaveSolution (string outputSolutionFile)
		{
			var output = File.CreateText (outputSolutionFile);
			output.WriteLine ("Microsoft Visual Studio Solution File, Format Version 12.00");
			output.WriteLine ("# Visual Studio 2012");
			foreach (var project in projects) {
				output.WriteLine ($"Project(\"{{{project.ProjectTypeGuids}}}\") = \"{project.Name}\", \"{project.Path}\", \"{{{project.ProjectGuid}}}\"");
				output.WriteLine ("EndProject");
			}
			output.WriteLine ("Global");

			output.WriteLine ("\tGlobalSection(SolutionConfigurationPlatforms) = preSolution");
			foreach (var scp in solution_configuration_platforms)
				output.WriteLine ($"\t\t{scp} = {scp}");
			output.WriteLine ("\tEndGlobalSection");

			output.WriteLine ("\tGlobalSection(ProjectConfigurationPlatforms) = postSolution");
			foreach (var pcp in project_configuration_platforms)
				output.WriteLine ($"\t\t{{{pcp.ProjectGuid}}}.{pcp.SolutionConfigurationName}.{pcp.Property} = {pcp.ConfigurationValue}");
			output.WriteLine ("\tEndGlobalSection");

			output.WriteLine ("\tGlobalSection(NestedProjects) = preSolution");
			foreach (var npg in NestedProjects)
				foreach (var np in npg)
					output.WriteLine ($"\t\t{{{np.Item}}} = {{{np.Parent}}}");
			output.WriteLine ("\tEndGlobalSection");

			foreach (var other in other_global_sections) {
				output.WriteLine ($"\tGlobalSection({other.Name}) = {other.Type}");
				output.WriteLine (other.Value);
				output.WriteLine ("\tEndGlobalSection");
			}

			output.WriteLine ("EndGlobal");

			output.Close ();
		}

		public void LoadXml (string inputXmlFile, string baseDirectory = null)
		{
			base_directory = baseDirectory ?? Path.GetDirectoryName (inputXmlFile);
			var doc = XDocument.Load (inputXmlFile, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo | LoadOptions.SetBaseUri);
			var rootElement = doc.Element ("solution");

			var projectsElement = rootElement.Element ("projects");
			this.projects = new List<ProjectInSolution> ();
			foreach (var projectElement in projectsElement.Elements ("project")) {
				var projectFile = projectElement.Attribute ("file")?.Value;
				if (projectFile == null) {
					Console.Error.WriteLine ($"WARNING: <project> element '{projectElement.Attribute ("name")}' has no 'file' attribute. Ignoring.");
					continue;
				}
				var projectFileFullPath = Path.Combine (base_directory, projectFile.Replace ('\\', Path.DirectorySeparatorChar));
				if (!File.Exists (projectFileFullPath)) {
					Console.Error.WriteLine ($"WARNING: <project> file '{projectFileFullPath}' does not exist. Ignoring.");
					continue;
				}
				var project = ProjectRootElement.Open (projectFileFullPath);
				Func<string, string> getProp = name => project.PropertyGroups.SelectMany (g => g.Properties).FirstOrDefault (p => p.Name.Equals (name))?.Value;
				projects.Add (new ProjectInSolution {
					Path = projectFile,
					Name = projectElement.Attribute ("name")?.Value ?? Path.GetFileNameWithoutExtension (projectFileFullPath),
					ProjectGuid = getProp ("ProjectGuid") ?? "{" + Guid.NewGuid () + "}",
					ProjectTypeGuids = getProp ("ProjectTypeGuids"),
				});
			}

			var solutionConfigsElement = rootElement.Element ("config-names-for-solution");
			this.solution_configuration_platforms = new List<string> ();
			foreach (var configElement in solutionConfigsElement.Elements ("config"))
				solution_configuration_platforms.Add (configElement.Attribute ("name").Value);

			var folders = rootElement.Element ("solution-folders");
			this.nested_projects = new List<Nest> ();
			foreach (var folder in folders.Elements ("folder")) {
				string folderName = folder.Attribute ("name").Value;
				// if the solution folder is not part of project list, then add it.
				if (!projects.Any (p => p.Name.Equals (folderName, StringComparison.OrdinalIgnoreCase)))
					projects.Add (new ProjectInSolution {
						Name = folderName,
						Path = folderName,
						ProjectTypeGuids = ProjectTypeGuids.SolutionFolder,
						ProjectGuid = "{" + Guid.NewGuid () + "}" });

				foreach (var item in folder.Elements ("project")) {
					var itemName = item.Attribute ("name").Value;
					nested_projects.Add (new Nest { Item = ProjectNameToGuid (itemName), Parent = ProjectNameToGuid (folderName) });
				}
			}

			string [] default_properties = { "ActiveCfg", "Build.0" };
			var projectConfigsElement = rootElement.Element ("project-config-properties");
			this.project_configuration_platforms = new List<ProjectConfigurationPlatform> ();
			foreach (var projectElement in projectConfigsElement.Elements ("project")) {
				var projectName = projectElement.Attribute ("name").Value;
				var projectGuid = ProjectNameToGuid (projectName);
				var pattern = projectElement.Attribute ("pattern")?.Value;
				switch (pattern) {
				case "all":
					foreach (var config in solution_configuration_platforms)
						foreach (var prop in default_properties)
							project_configuration_platforms.Add (new ProjectConfigurationPlatform {
								ProjectGuid = projectGuid,
								SolutionConfigurationName = config,
								Property = prop,
								ConfigurationValue = config.Replace ("AnyCPU", "Any CPU")
						});
					break;
				case null:
					foreach (var configPlatformElement in projectElement.Elements ("config-platform")) {
						var config = configPlatformElement.Attribute ("config").Value;
						var prop = configPlatformElement.Attribute ("property").Value;
						var value = configPlatformElement.Attribute ("value").Value;
						project_configuration_platforms.Add (new ProjectConfigurationPlatform {
							ProjectGuid = projectGuid,
							SolutionConfigurationName = config,
							Property = prop,
							ConfigurationValue = value });
					}
					break;
				default:
					throw new ArgumentException ("Unexpected pattern value: " + pattern);
				}
			}

			var otherElement = rootElement.Element ("other-global-sections");
			this.other_global_sections = new List<GlobalSection> ();
			foreach (var sectionElement in otherElement.Elements ("section"))
				other_global_sections.Add (new GlobalSection {
					Name = sectionElement.Attribute ("name")?.Value,
					Type = sectionElement.Attribute ("type")?.Value,
					Value = sectionElement.Value,
				});
		}

		string ProjectNameToGuid (string name)
		{
			return projects.FirstOrDefault (p => p.Name == name)?.ProjectGuid;
		}
		
		public void SaveXml (string outputXmlFile)
		{
			var output = new XDocument ();
			output.Add (new XElement ("solution"));

			var projectsElement = new XElement ("projects");
			output.Root.Add (new XComment ("Project name is omitted if it is identical to the file name (without extension)."));
			output.Root.Add (projectsElement);
			foreach (var proj in projects) {
				var file = proj.Path.Replace ('\\', Path.DirectorySeparatorChar);
				if (File.Exists (Path.Combine (base_directory, file)))
					projectsElement.Add (new XElement ("project",
						Path.GetFileNameWithoutExtension (file) != proj.Name ? new XAttribute ("name", proj.Name) : null,
						new XAttribute ("file", file)
						));
			}

			var configsElement = new XElement ("config-names-for-solution");
			output.Root.Add (configsElement);
			foreach (var plat in solution_configuration_platforms)
				configsElement.Add (new XElement ("config", new XAttribute ("name", plat)));

			var foldersElement = new XElement ("solution-folders");
			output.Root.Add (foldersElement);

			foreach (var nest in NestedProjects) {
				var folder = ProjectGuidToName (nest.Key);
				var folderElement = new XElement ("folder", new XAttribute ("name", folder));
				foldersElement.Add (folderElement);
				foreach (var n in nest) {
					var item = ProjectGuidToName (n.Item);
					folderElement.Add (new XElement ("project", new XAttribute ("name", item)));
				}
			}

			var pcpElement = new XElement ("project-config-properties");
			output.Root.Add (new XComment ("Value 'all' indicates that the project contains `ActiveCfg` and `Build.0` properties for all configs."));
			output.Root.Add (pcpElement);
			foreach (var cfgGroup in ProjectConfigurationPlatforms) {
				var deftpl = ProjectConfigurationPlatformTemplate.CreateTemplateFor ("all", solution_configuration_platforms);
				var pattern = deftpl.Matches (cfgGroup) ? "all" : null;
				var p = new XElement ("project", new XAttribute ("name", ProjectGuidToName (cfgGroup.Key)));
				if (pattern != null)
					p.Add (new XAttribute ("pattern", pattern));
				else {
					foreach (var cfg in cfgGroup) {
						p.Add (new XElement ("config-platform",
						                     new XAttribute ("config", cfg.SolutionConfigurationName),
						                     new XAttribute ("property", cfg.Property),
						                     new XAttribute ("value", cfg.ConfigurationValue)
						                    ));
					}
				}
				pcpElement.Add (p);
			}

			var otherElement = new XElement ("other-global-sections");
			output.Root.Add (otherElement);
			foreach (var other in other_global_sections)
				otherElement.Add (new XElement ("section", new XAttribute ("name", other.Name), new XAttribute ("type", other.Type), new XText (other.Value)));

			output.Save (outputXmlFile);
		}

		public void Dump ()
		{
			Console.WriteLine ("Projects");

			foreach (var proj in projects)
				Console.WriteLine ($"\t{ProjectTypeGuids.GetNameIfAvailable (proj.ProjectTypeGuids)} {proj.Path}");

			Console.WriteLine ();
			Console.WriteLine ("SolutionConfigurationPlatform");
			foreach (var plat in solution_configuration_platforms)
				Console.WriteLine ('\t' + plat);

			Console.WriteLine ();
			Console.WriteLine ("Solution Folders:");
			foreach (var nest in NestedProjects) {
				var folder = ProjectGuidToName (nest.Key);
				Console.WriteLine ("\t" + folder);
				foreach (var n in nest) {
					var item = ProjectGuidToName (n.Item);
					Console.WriteLine ("\t\t" + item);
				}
			}

			Console.WriteLine ();
			Console.WriteLine ("ProjectConfigurationPlatforms");
			foreach (var cfgGroup in ProjectConfigurationPlatforms) {
				var deftpl = ProjectConfigurationPlatformTemplate.CreateTemplateFor ("all", solution_configuration_platforms);
				var matches = deftpl.Matches (cfgGroup);
				var pattern = deftpl.Matches (cfgGroup) ? "all" : null;
				Console.WriteLine ("\t" + GetProjectName (cfgGroup.First ()) + " : " + pattern);
				if (pattern != null)
					continue;
				foreach (var cfg in cfgGroup)
					Console.WriteLine ($"\t\tCFG#{SolutionConfigurationIndexOf (cfg)} / {cfg.Property} / {ProjectConfigurationPlatform.GetValueIfAvailable (cfg.ConfigurationValue)}");
			}

			Console.WriteLine ();

			foreach (var sec in other_global_sections) {
				Console.WriteLine ($"OTHER SECTION: {sec.Name} ({sec.Type})");
				Console.WriteLine (sec.Value);
			}
		}

		public string ProjectGuidToName (string guid)
		{
			return projects.FirstOrDefault (p => p.ProjectGuid == guid)?.Name;
		}

		public string GetProjectName (ProjectConfigurationPlatform item)
		{
			return projects.First (p => p.ProjectGuid == item.ProjectGuid).Name;
		}

		public int SolutionConfigurationIndexOf (ProjectConfigurationPlatform item)
		{
			return solution_configuration_platforms.IndexOf (item.SolutionConfigurationName);
		}

		public void LoadSolution (string solutionFile)
		{
			base_directory = Path.GetDirectoryName (solutionFile);
			string sln = File.ReadAllText (solutionFile);

			projects = GetProjects (sln).ToList ();
			solution_configuration_platforms = GetSolutionConfigurationPlatforms (sln).ToList ();
			nested_projects = GetNestedProjects (sln).ToList ();
			project_configuration_platforms = GetProjectConfigurationPlatforms (sln).ToList ();
			global_section_names = GetGlobalSectionNames (sln).ToList ();
			other_global_sections = global_section_names.Where (n => !IsWellKnownSection (n))
								    .Select (n => GetGlobalSection (sln, n))
			                                            .ToList ();
		}

		string base_directory;

		public bool IsWellKnownSection (string name)
		{
			switch (name) {
			case "SolutionConfigurationPlatforms":
			case "ProjectConfigurationPlatforms":
			case "NestedProjects":
				return true;
			}
			return false;
		}

		GlobalSection GetGlobalSection (string sln, string globalSectionName)
		{
			int sidx = sln.IndexOf ($"GlobalSection({globalSectionName})", StringComparison.Ordinal);
			if (sidx < 0)
				return null;
			int eidx = sln.IndexOf ("EndGlobalSection", sidx, StringComparison.Ordinal);
			int lidx = sln.IndexOf ('\n', sidx);
			int eqidx = sln.IndexOf ('=', sidx);
			var type = eqidx < 0 ? null : sln.Substring (eqidx + 1, lidx - eqidx).Trim ();
			return new GlobalSection { Name = globalSectionName, Type = type, Value = sln.Substring (lidx, eidx - lidx) };
		}

		List<ProjectInSolution> projects;

		IEnumerable<ProjectInSolution> GetProjects (string sln)
		{
			var rex = new Regex (@"Project\(""\{(?<project_type_guid>.+)\}""\) = ""(?<name>.+)"", ""(?<path>.+)"", ""\{(?<project_guid>.+)\}");
			foreach (Match m in rex.Matches (sln))
				yield return new ProjectInSolution {
					ProjectTypeGuids = m.Groups [1].Value,
					Name = m.Groups [2].Value,
					Path = m.Groups [3].Value,
					ProjectGuid = m.Groups [4].Value,
				};
		}

		List<string> global_section_names;
		public IList<string> GlobalSectionNames => global_section_names;

		IEnumerable<string> GetGlobalSectionNames (string sln)
		{
			var rex = new Regex (@"GlobalSection\((?<name>.+)\)");
			foreach (Match section in rex.Matches (sln))
				yield return section.Groups [1].Value;
		}

		List<string> solution_configuration_platforms;
		public IList<string> SolutionConfigurationPlatforms => solution_configuration_platforms;

		IEnumerable<string> GetSolutionConfigurationPlatforms (string sln)
		{
			var section = GetGlobalSection (sln, "SolutionConfigurationPlatforms");
			var lines = section.Value.Split ('\n').Select (s => s.Trim ());
			foreach (var line in lines)
				if (line.Length > 0)
					yield return line.Substring (0, line.IndexOf ('=')).Trim ();
		}

		List<Nest> nested_projects;
		public IList<IGrouping<string, Nest>> NestedProjects => nested_projects.GroupBy (p => p.Parent).ToList ();

		IEnumerable<Nest> GetNestedProjects (string sln)
		{
			var section = GetGlobalSection (sln, "NestedProjects");
			if (section == null)
				yield break;
			var lines = section.Value.Split ('\n').Select (s => s.Trim ());
			int i;
			Func<string, string> trim = s => s.Trim ().TrimStart ('{').TrimEnd ('}');
			foreach (var line in lines)
				if ((i = line.IndexOf ('=')) > 0)
					yield return new Nest {
						Item = trim (line.Substring (0, i)),
						Parent = trim (line.Substring (i + 1))
					};
		}

		List<ProjectConfigurationPlatform> project_configuration_platforms;
		public IEnumerable<IGrouping<string, ProjectConfigurationPlatform>> ProjectConfigurationPlatforms => project_configuration_platforms.GroupBy (cfg => cfg.ProjectGuid);

		IEnumerable<ProjectConfigurationPlatform> GetProjectConfigurationPlatforms (string sln)
		{
			var section = GetGlobalSection (sln, "ProjectConfigurationPlatforms");
			if (section == null)
				yield break;
			var lines = section.Value.Split ('\n').Select (s => s.Trim ());
			foreach (var line in lines) {
				int i = line.IndexOf ('=');
				if (i < 0)
					continue;
				var key = line.Substring (0, i).Trim ();
				var value = line.Substring (i + 1).Trim ();
				var tokens = key.Split ('.');
				var guid = tokens [0];
				int i2 = line.IndexOf ('}');
				yield return new ProjectConfigurationPlatform {
					ProjectGuid = guid.Substring (1, guid.Length - 2),
					SolutionConfigurationName = tokens [1],
					Property = string.Join (".", tokens.Skip (2)),
					ConfigurationValue = value
				};
			}
		}

		List<GlobalSection> other_global_sections;
		public IList<GlobalSection> OtherGlobalSections => other_global_sections;
	}

	public class GlobalSection
	{
		public string Name { get; set; }
		public string Type { get; set;  }
		public string Value { get; set; }
	}

	public class ProjectInSolution {
		public string ProjectTypeGuids { get; set; }
		public string Name { get; set; }
		public string Path { get; set; }
		public string ProjectGuid { get; set; }
	}

	public class ProjectTypeGuids {
		public const string GenericProject = "9344BDBB-3E7F-41FC-A0DD-8665D75EE146";
		public const string CSharpDesktop = "FAE04EC0-301F-11D3-BF4B-00C04F79EFBC";
		public const string SolutionFolder = "2150E333-8FDC-42A3-9474-1A3956D46DE8";
		public const string SharedProject = "D954291E-2A0B-460D-934E-DC6B0785DB48";

		public static string GetNameIfAvailable (string projectTypeGuid)
		{
			return typeof (ProjectTypeGuids).GetFields ().FirstOrDefault (f => f.GetValue (null) as string == projectTypeGuid)?.Name ?? projectTypeGuid;
		}

		public static string GetProjectTypeGuidIfAvailable (string name)
		{
			return typeof (ProjectTypeGuids).GetFields ().FirstOrDefault (f => f.Name == name)?.GetValue (null) as string ?? name;
		}
	}

	public class ProjectConfigurationPlatformTemplate {
		public string Name { get; set; }
		public ProjectConfigurationPlatform [] TemplatedItems { get; set; }

		public bool Matches (IEnumerable<ProjectConfigurationPlatform> target)
		{
			if (target.Count () != TemplatedItems.Length)
				return false;
			bool [] occurred = new bool [TemplatedItems.Length];
			foreach (var entry in target) {
				int idx = Array.FindIndex (TemplatedItems, i =>
					i.Property == entry.Property &&
					i.SolutionConfigurationName == entry.SolutionConfigurationName &&
					i.ConfigurationValue == entry.ConfigurationValue);
				if (idx < 0 || idx >= occurred.Length || occurred [idx])
					return false;
				occurred [idx] = true;
			}
			return occurred.All (b => b);
		}

		static IEnumerable<ProjectConfigurationPlatform> createItems (IEnumerable<string> configs)
		{
			foreach (var scn in configs) {
				int sep = scn.LastIndexOf ('|');
				var conf = sep < 0 ? "" : scn.Substring (0, sep);
				var arch = scn.Substring (sep + 1);
				var confFixed = conf.EndsWith ("Debug") ? "Debug" : conf.EndsWith ("Release") ? "Release" : conf;
				var archFixed = arch == "AnyCPU" ? "Any CPU" : arch;
				foreach (var prop in new string [] { "ActiveCfg", "Build.0" })
					yield return new ProjectConfigurationPlatform {
						SolutionConfigurationName = scn,
						Property = prop,
						ConfigurationValue = sep < 0 ? scn : confFixed + '|' + archFixed
					};
			}
		}

		public static ProjectConfigurationPlatformTemplate CreateTemplateFor (string name, IEnumerable<string> configs)
		{
			return new ProjectConfigurationPlatformTemplate {
				Name = name,
				TemplatedItems = createItems (configs).ToArray ()
			};
		}
	}

	public class ProjectConfigurationPlatform {
		public string ProjectGuid { get; set; }
		public string SolutionConfigurationName { get; set; }
		public string Property { get; set; }
		public string ConfigurationValue { get; set; }

		public static string GetValueIfAvailable (string configurationValue)
		{
			return typeof (ProjectConfigurationPlatform).GetFields ().FirstOrDefault (f => f.GetValue (null) as string == configurationValue)?.Name ?? configurationValue;
		}
	}

	public class Nest {
		public string Item { get; set; }
		public string Parent { get; set; }
	}
}
