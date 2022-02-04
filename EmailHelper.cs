using Microsoft.Graph;

namespace graph_console
{
    public class EmailHelper
    {
        private static GraphServiceClient graphClient;

        public static void Initialize(GraphServiceClient client)
        {
            graphClient = client;
        }

        public static async Task<IUserMessagesCollectionPage> GetEmailsAsync()
        {
            try
            {
                return await graphClient.Me.Messages
                .Request()
                .Top(5)
                .Expand("attachments")
                .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting user's message: {ex.Message}");
                return null;
            }
        }

        public static async Task SendSimpleEmailAsync(string recipientAddress, string subject, string content)
        {
            try
            {
                // Message class ตัวแทนของข้อความที่ต้องการส่ง
                var message = new Message
                {
                    // กำหนดหัว Email
                    Subject = subject,

                    // กำหนดตัวเนื้อหาของ Email (Body)
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = content
                    },

                    // กำหนดผู้รับ สามารถกำหนดได้หลายคน
                    ToRecipients = new List<Recipient>()
                    {
                        // 1 Recpient
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = recipientAddress
                            }
                        }
                    }
                };


                var saveToSentItems = true;

                await graphClient.Me
                    .SendMail(message, saveToSentItems)
                    .Request()
                    .PostAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error sending message: {ex.Message}");
            }
        }

        public static async Task<Message> GetMessageWithAttachmentAsync(string messageId)
        {
            try
            {
                // Request ข้อมูลของ Message แบบระบุ Id
                return await graphClient.Me.Messages[messageId]
                .Request()
                // กำหนดขอข้อมูล Attachments ด้วย
                .Expand("attachments")
                .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting message's attachment: {ex.Message}");
                return null;
            }
        }

    }
}