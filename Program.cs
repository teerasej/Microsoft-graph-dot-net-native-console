using Microsoft.Extensions.Configuration;
using graph_console;
using Microsoft.Graph;

var appConfig = new ConfigurationBuilder()
        .AddUserSecrets<Program>()
        .Build();

if (string.IsNullOrEmpty(appConfig["appId"]) ||
    string.IsNullOrEmpty(appConfig["scopes"]))
{
    Console.WriteLine("Missing or invalid appsettings.json...exiting");
    return;
}

var appId = appConfig["appId"];
var scopesString = appConfig["scopes"];
var scopes = scopesString.Split(';');

GraphHelper.Initialize(appId, scopes, (code, cancellation) =>
{
    Console.WriteLine(code.Message);
    return Task.FromResult(0);
});

var accessToken = GraphHelper.GetAccessTokenAsync(scopes).Result;

EmailHelper.Initialize(GraphHelper.graphClient);
OneDriveHelper.Initialize(GraphHelper.graphClient);
CalendarHelper.Initialize(GraphHelper.graphClient);
TeamHelper.Initialize(GraphHelper.graphClient);

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. Hello, Me");

    Console.WriteLine("---- Email ----");
    Console.WriteLine("3. Show Emails");
    Console.WriteLine("4. Send Mail");

    Console.WriteLine("---- OneDrive ----");
    Console.WriteLine("5. List all files in My Files");
    Console.WriteLine("6. Create new folder");
    Console.WriteLine("7. Download File");
    Console.WriteLine("8. Upload File");

    Console.WriteLine("---- Mail to Drive ----");
    Console.WriteLine("9. Download attachment to OneDrive");

    Console.WriteLine("---- Team ----");
    Console.WriteLine("10. Create Team");

    Console.WriteLine("---- Calendar ----");
    Console.WriteLine("11. Create Calendar");
    Console.WriteLine("12. Create Event");

    try
    {
        choice = int.Parse(Console.ReadLine());
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            Console.WriteLine($"Access token: {accessToken}\n");
            break;

        case 2:
            // Permission User.Read
            var user = await GraphHelper.GetMeAsync();
            Console.WriteLine($"Hi, {user.DisplayName}");
            break;

        case 3:
            var emails = await EmailHelper.GetEmailsAsync();
            for (int i = 0; i < emails.Count; i++)
            {
                Message message = emails[i];

                Console.WriteLine($"{i+1}: message {message.Id}");
                Console.WriteLine($"    Sender: {message.Sender.EmailAddress.Name} {message.Sender.EmailAddress.Address} ");
                Console.WriteLine($"    Subject: {message.Subject}");

                if((bool)message.HasAttachments) 
                {
                    Console.WriteLine($"    has {message.Attachments.Count} attachment files");

                    foreach (var attachment in message.Attachments)
                    {
                        if(attachment is FileAttachment) 
                        {
                            var fileAttachment = attachment as FileAttachment;
                            System.IO.File.WriteAllBytes(fileAttachment.Name, fileAttachment.ContentBytes);
                        }
                    }
                }

                Console.WriteLine("\n");
            }
            break;

        case 4:
            Console.WriteLine("Recipient Email:");
            var recipientEmailAddress = Console.ReadLine();
            
            Console.WriteLine("Subject:");
            var subject = Console.ReadLine();

            Console.WriteLine("Content:");
            var content = Console.ReadLine();

            Console.WriteLine("Sending...");
            await EmailHelper.SendSimpleEmailAsync(recipientEmailAddress, subject, content);
            Console.WriteLine("Sent!\n");
            break;

        case 5: 
            var onedriveItems = await OneDriveHelper.GetUserDriveItemsAsync();
            Console.WriteLine($"Items in drive: {onedriveItems.Count}");

            for (int i = 0; i < onedriveItems.Count; i++)
            {
                DriveItem driveItem = onedriveItems[i];

                var itemType = "File";
                if(driveItem.Folder != null) 
                {
                    itemType = "Folder";
                }

                Console.WriteLine($"{i+1}: ({itemType}) {driveItem.Name}");
                Console.WriteLine($"       id: {driveItem.Id}");
            }
            Console.WriteLine();
            break;

        case 6: 
            Console.WriteLine("Folder Name:");
            var folderName = Console.ReadLine();

            Console.WriteLine("Creating...");
            await OneDriveHelper.CreateNewFolderAsync(folderName);
            Console.WriteLine("Done!\n");
            break;
        
        case 7: 
            Console.WriteLine("Item Id:");
            var itemId = Console.ReadLine();

            Console.WriteLine("Downloading...");
            await OneDriveHelper.DownloadFileAsync(itemId);
            Console.WriteLine("Done!\n");
            break;

        case 8:
            Console.WriteLine("file name to upload (put file in project's root only):");
            var fileName = Console.ReadLine();

            Console.WriteLine("Uploading...");
            await OneDriveHelper.UploadFileAsync(fileName);
            Console.WriteLine("Done!\n");
            break;

        case 9:
            // อ่่าน message Id จาก console
            Console.WriteLine("Target Message Id:");
            var targetMessageId = Console.ReadLine();

            // ขอข้อมูล Message พร้อม Attachment
            var targetMessage = await EmailHelper.GetMessageWithAttachmentAsync(targetMessageId);
            

            if ((bool)targetMessage.HasAttachments)
            {
                Console.WriteLine($"   Attachments: {targetMessage.Attachments.Count}");

                foreach (var attachment in targetMessage.Attachments.CurrentPage)
                {
                    if(attachment is FileAttachment) 
                    {
                        // ดาวน์โหลด attachments มาเก็บไว้ในเครื่อง
                        var fileAttachment = attachment as FileAttachment;
                        System.IO.File.WriteAllBytes(fileAttachment.Name, fileAttachment.ContentBytes);
                        Console.WriteLine($"   Downloaded: {fileAttachment.Name}");

                        // อัพโหลดไปที่ OneDrive
                        Console.WriteLine($"        Uploading to OneDrive: {fileAttachment.Name}");
                        await OneDriveHelper.UploadFileAsync(fileAttachment.Name);
                        Console.WriteLine($"        Done.");
                    }
                }
            }
            else 
            {
                Console.WriteLine("Sorry, this email has no attachment.\n");
            }

            break;

        case 11:

            Console.WriteLine("Calendar name:");
            var calendarName = Console.ReadLine();

            Console.WriteLine("     Creating Calendar...");
            var createdCalendar = await CalendarHelper.CreateCalendar(calendarName);
            Console.WriteLine("     Done.");

            break;

        case 12:
            Console.WriteLine("Event Name:");
            var eventName = Console.ReadLine();

            Console.WriteLine("Event Description:");
            var eventDescription = Console.ReadLine();

            Console.WriteLine("Location name:");
            var locationName = Console.ReadLine();

            Console.WriteLine("     Creating Event...");
            var createdEvent = await CalendarHelper.CreateEvent(eventName, eventDescription, locationName);
            Console.WriteLine("     Done.");
            break;

        case 10:
            // รับชื่อ Team
            Console.WriteLine("Team Name:");
            var teamName = Console.ReadLine();

            // รับรายละเอียดของ Team
            Console.WriteLine("Team Description:");
            var teamDescription = Console.ReadLine();


            // สั่งสร้าง Team
            Console.WriteLine("     Creating team...");
            var createdTeam = await TeamHelper.CreateTeamAsync(teamName, teamDescription);
            Console.WriteLine("     Done.");
            break;

        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}