using {{SafeProjectName}}.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;

namespace {{SafeProjectName}}
{
    public static class Repair
    {
        [FunctionName("repair")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            // Log that the HTTP trigger function received a request.
            log.LogInformation("C# HTTP trigger function processed a request.");

            // Get the query parameters from the request.
            string assignedTo = req.Query["assignedTo"];

            // Create the repair records.
            var repairRecords = new RepairModel[]
            {
                new RepairModel {
                    Id = 1,
                    Title = "Oil change",
                    Description = "Need to drain the old engine oil and replace it with fresh oil to keep the engine lubricated and running smoothly.",
                    AssignedTo = "Karin Blair",
                    Date = "2023-05-23",
                    Image = "https://www.howmuchisit.org/wp-content/uploads/2011/01/oil-change.jpg"
                }
            };

            // Filter the repair records by the assignedTo query parameter.
            var repairs = repairRecords.Where(r =>
            {
                // Split assignedTo into firstName and lastName
                var parts = r.AssignedTo.Split(' ');

                // Check if assignedTo matches firstName or lastName
                return r.assignedTo.Equals(assignedTo?.Trim(), StringComparison.InvariantCultureIgnoreCase) ||
                       parts[0].Equals(assignedTo?.Trim(), StringComparison.InvariantCultureIgnoreCase) ||
                       parts[1].Equals(assignedTo?.Trim(), StringComparison.InvariantCultureIgnoreCase);
            });
            
            // Return filtered repair records, or an empty array if no records were found.
            var results = new { results = repairs };
            return new OkObjectResult(results);
        }
    }
}