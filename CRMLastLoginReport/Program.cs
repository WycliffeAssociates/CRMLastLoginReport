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
            Console.WriteLine("Logging in");
            CrmServiceClient service = new CrmServiceClient(args[0]);
            QueryExpression systemUserQuery = new QueryExpression("systemuser");
            systemUserQuery.ColumnSet = new ColumnSet("fullname","isdisabled");
            var systemUsers = service.RetrieveMultiple(systemUserQuery).Entities;
            FilterExpression filter = new FilterExpression(LogicalOperator.Or);
            filter.AddCondition("action", ConditionOperator.Equal, 64);
            filter.AddCondition("action", ConditionOperator.Equal, 65);
            List<string> dump = new List<string>() { @"""User"",""Date"",""Status""" };
            foreach (Entity e in systemUsers)
            {
                QueryExpression query = new QueryExpression("audit");
                query.ColumnSet = new ColumnSet("createdon");
                query.Criteria.AddCondition("objectid", ConditionOperator.Equal, e.Id);
                query.Criteria.AddFilter(filter);
                query.AddOrder("createdon", OrderType.Descending);
                query.TopCount = 1;
                string lastAccess = "None";
                var result = service.RetrieveMultiple(query).Entities;
                if (result.Count != 0)
                {
                    lastAccess = ((DateTime)result[0]["createdon"]).ToShortDateString();
                }
                Console.WriteLine($"{e["fullname"]}: {lastAccess}");
                string status = (bool)e["isdisabled"] ? "Disabled" : "Enabled";
                dump.Add($@"""{(string)e["fullname"]}"", ""{lastAccess}"",""{status}""");
            }
            File.WriteAllLines("result.csv", dump);
            Console.WriteLine("All done");
            Console.ReadLine();
        }
    }
}
