using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.IO;

namespace CRMLastLoginReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Extract the command line args
            CommandLineArgs options = new CommandLineArgs();
            if (!CommandLine.Parser.Default.ParseArguments(args, options))
            {
                Environment.Exit(1);
            }

            // Log into CRM
            Console.WriteLine("Logging in");
            CrmServiceClient service = new CrmServiceClient(options.ConnectionString);
            if (!string.IsNullOrEmpty(service.LastCrmError))
            {
                Console.WriteLine($"Unable to connect to crm {service.LastCrmError}");
                Environment.Exit(2);
            }

            // Get all of the system users
            QueryExpression systemUserQuery = new QueryExpression("systemuser");
            systemUserQuery.ColumnSet = new ColumnSet("fullname","isdisabled");
            if (options.IgnoreDeactivated)
            {
                systemUserQuery.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);
            }
            var systemUsers = service.RetrieveMultiple(systemUserQuery).Entities;

            // Build a filter for the api access that we care about
            FilterExpression filter = new FilterExpression(LogicalOperator.Or);
            filter.AddCondition("action", ConditionOperator.Equal, 64);
            filter.AddCondition("action", ConditionOperator.Equal, 65);

            // Create the header for the CSV
            List<string> dump = new List<string>() { @"""User"",""Date"",""Status""" };

            foreach (Entity e in systemUsers)
            {
                // Get the last time a user accessed the system
                QueryExpression query = new QueryExpression("audit");
                query.ColumnSet = new ColumnSet("createdon");
                query.Criteria.AddCondition("objectid", ConditionOperator.Equal, e.Id);
                query.Criteria.AddFilter(filter);
                query.AddOrder("createdon", OrderType.Descending);
                query.TopCount = 1;
                string lastAccess = "";
                var result = service.RetrieveMultiple(query).Entities;
                if (result.Count != 0)
                {
                    lastAccess = ((DateTime)result[0]["createdon"]).ToShortDateString();
                }
                Console.WriteLine($"{e["fullname"]}: {lastAccess}");
                string status = (bool)e["isdisabled"] ? "Disabled" : "Enabled";

                // Add that to the CSV we are building
                dump.Add($@"""{(string)e["fullname"]}"", ""{lastAccess}"",""{status}""");
            }

            // Attempt to write the resulting file
            try
            {
                File.WriteAllLines("result.csv", dump);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error writing output file");
                Console.WriteLine(ex.Message);
                Environment.Exit(3);
            }
            Console.WriteLine("All done");
            Console.ReadLine();
        }
    }
}
