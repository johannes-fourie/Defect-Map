using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using static System.Console;

namespace TFSBugQuery
{
    public class QueryBugs
    {
        private static string _log_file = null;

        public static void RunQuery(string tfs_url, string project_name, string query_folder, string query_name, string log_file, int take_first = 0)
        {
            _log_file = log_file;

            WriteLine("Initialize TFS Server object");
            TfsConfigurationServer configuration_server = TfsConfigurationServerFactory.GetConfigurationServer(new Uri(tfs_url));

            WriteLine(@"Get the catalog of team project collections");
            CatalogNode catalog_node = configuration_server.CatalogNode;

            WriteLine(@"Get all CatalogNodes which are ProjectCollection");
            ReadOnlyCollection<CatalogNode> tpc_nodes = catalog_node.QueryChildren(new Guid[] { CatalogResourceTypes.ProjectCollection }, false, CatalogQueryOptions.None);

            WriteLine(@"Get InstanceId of a ProjectCollection");
            Guid tpc_id = Guid.Empty;
            foreach (CatalogNode tpc_node in tpc_nodes)
            {
                tpc_id = new Guid(tpc_node.Resource.Properties["InstanceId"]);
                break;
            }

            WriteLine(@"Fill list of projects in a local variable");
            TfsTeamProjectCollection project_collection = configuration_server.GetTeamProjectCollection(tpc_id);
            project_collection.Authenticate();

            WriteLine(@"Get WorkItem Tracking client for workitem collection for selected ProjectCollection");
            WorkItemStore work_item_store = project_collection.GetService<WorkItemStore>();

            WriteLine(@"Get Project from Tracking client");

            Project project = work_item_store.Projects[project_name];

            WriteLine(@"Run Query");
            QueryFolder team_query_folder = project.QueryHierarchy[query_folder] as QueryFolder;
            QueryItem query_item = team_query_folder[query_name];
            QueryDefinition queryDefinition = work_item_store.GetQueryDefinition(query_item.Id);

            Dictionary<string, string> variables = new Dictionary<string, string> { { "project", query_item.Project.Name } };

            WorkItemCollection work_item_collection = work_item_store.Query(queryDefinition.QueryText, variables);

            WriteLine(@"Get Source Control/Version Control repository for selected project collection");
            VersionControlServer version_control_server = project_collection.GetService<VersionControlServer>();
            WriteLine(@"Get Details of Version Control using artifact provider");
            VersionControlArtifactProvider artifact_provider = version_control_server.ArtifactProvider;

            Write_to_excel(new string[] { "WorkItemID", "WorkItemTitle", "ChangesetId", "CreationDate", "ChangeType", "File" });

            WriteLine(@"Iterate through each item to get its details");

            IEnumerable<WorkItem> work_items = take_first > 0 ? work_item_collection.OfType<WorkItem>().Take(take_first) : work_item_collection.OfType<WorkItem>();

            foreach (WorkItem work_item in work_items)
            {
                WriteLine($"    ->{work_item.Id}");
                IEnumerable<Changeset> changesets = work_item.Links.OfType<ExternalLink>().Select(link =>
                {
                    Changeset set;
                    try
                    {
                        set = artifact_provider.GetChangeset(new Uri(link.LinkedArtifactUri));
                    }
                    catch (Exception ex)
                    {
                        BackgroundColor = ConsoleColor.Red;
                        WriteLine(ex.Message);
                        BackgroundColor = ConsoleColor.Black;
                        set = null;
                    }
                    return set;
                })
                .Where(s => s != null);

                foreach (Changeset changeset in changesets)
                {
                    WriteLine($"        {changeset.ChangesetId}");
                    foreach (Change change in changeset.Changes)
                    {
                        WriteLine($"           {change.Item.ServerItem}");

                        var data = new string[] {
                            work_item.Id.ToString(),
                            work_item.Title,
                            changeset.ChangesetId.ToString(),
                            changeset.CreationDate.ToString(),
                            change.ChangeType.ToString(),
                            change.Item.ServerItem
                        };
                        Write_to_excel(data);
                    }
                }
            }
        }

        private static void Write_to_excel(string[] data_row)
        {
            StreamWriter streamWriter = new StreamWriter(_log_file, true);

            for (int j = 0; j < data_row.Length; j++)
                streamWriter.Write($"{data_row[j]};");

            streamWriter.WriteLine();
            streamWriter.Close();
        }
    }
}