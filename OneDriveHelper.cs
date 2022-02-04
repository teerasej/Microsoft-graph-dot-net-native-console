using Microsoft.Graph;

namespace graph_console
{
    public class OneDriveHelper
    {
        private static GraphServiceClient graphClient;

        public static void Initialize(GraphServiceClient client)
        {
            graphClient = client;
        }

        public static async Task<IDriveItemChildrenCollectionPage> GetUserDriveItemsAsync()
        {
            try
            {
                return await graphClient.Me.Drive.Root.Children
                .Request()
                .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error on get files: {ex.Message}");
                return null;
            }
        }

        public static async Task<DriveItem> CreateNewFolderAsync(string name = "temp")
        {
            var driveItem = new DriveItem
            {
                Name = name,
                Folder = new Folder { },
                AdditionalData = new Dictionary<string, object>()
                {
                    {"@microsoft.graph.conflictBehavior", "rename"}
                }
            };

            try
            {
                return await graphClient.Me
                .Drive
                .Root
                .Children
                .Request()
                .AddAsync(driveItem);
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error create folder: {ex.Message}");
                return null;
            }
        }

        public static async Task DownloadFileAsync(string driveItemId)
        {
            try
            {
                // เข้าถึงตัว DriveItem ที่ต้องการ
                var driveItem = await graphClient.Me.Drive.Items[driveItemId].Request().GetAsync();

                // เข้าถึง Content ของ DriveItem เพื่อใช้ stream
                var stream = await graphClient.Me.Drive.Items[driveItemId].Content.Request().GetAsync();

                // บันทึก Stream เป็นไฟล์
                using (var fileStream = System.IO.File.Create(driveItem.Name))
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                }
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error downloading a file: {ex.Message}");
            }
        }

        public static async Task UploadFileAsync(string fileName)
        {
            using (var fileStream = System.IO.File.OpenRead(fileName))
            {
                // กำหนดเงื่อนไขว่าถ้าเจอไฟล์ชื่อซ้ำกัน ให้แทนที่ไฟล์เดิมไปเลย
                var uploadProps = new DriveItemUploadableProperties
                {
                    ODataType = null,
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", "replace" }
                    }
                };

                // สร้าง upload session
                var uploadSession = await graphClient.Me.Drive.Root
                    .ItemWithPath(fileName)
                    .CreateUploadSession(uploadProps)
                    .Request()
                    .PostAsync();

                // แบ่งขนาดไฟล์
                int maxSliceSize = 320 * 1024;
                var fileUploadTask =
                    new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize);

                // สร้าง function ที่จะทำงานขณะที่การอัพโหลดดำเดินไป
                IProgress<long> progress = new Progress<long>(prog =>
                {
                    Console.WriteLine($"        Uploaded {prog} bytes of {fileStream.Length} bytes");
                });

                try
                {
                    // อัพโหลดไฟล์
                    var uploadResult = await fileUploadTask.UploadAsync(progress);

                    // เช็คว่าการอัพโหลดเสร็จสมบูรณ์ไหม
                    if (uploadResult.UploadSucceeded)
                    {
                        // ถ้าอัพโหลดสมบูรณ์ ก็สามารถดึง id ของ item มาแสดงได้
                        Console.WriteLine($"        Upload complete, item ID: {uploadResult.ItemResponse.Id}");
                    }
                    else
                    {
                        Console.WriteLine("         Upload failed");
                    }
                }
                catch (ServiceException ex)
                {
                    Console.WriteLine($"        Error uploading: {ex.ToString()}");
                }
            }

        }
    }
}