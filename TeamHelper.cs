using Microsoft.Graph;

namespace graph_console
{
    public class TeamHelper
    {
        private static GraphServiceClient graphClient;

        public static void Initialize(GraphServiceClient client)
        {
            graphClient = client;
        }

        public static async Task<Team> CreateTeamAsync(string teamName, string description)
        {
            // กำหนดข้อมูลของ Team ที่จะสร้างขึ้นมาใหม่
            var newTeam = new Team
            {
                DisplayName = teamName,
                Description = description,

                // กำหนดใช้ template แบบมาตรฐาน
                AdditionalData = new Dictionary<string, object>()
                {
                    {"template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"}
                },
            };

            try
            {
                // สร้างเพิ่มข้อมูล Team ใหม่เข้าไป
                return await graphClient.Teams.Request().AddAsync(newTeam);
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating team: {ex.Message}");
                return null;
            }
        }

    }
}