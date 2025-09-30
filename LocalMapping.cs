using System.Text;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System.Globalization;
using Microsoft.AspNetCore.WebUtilities;
using Quartz;
using static DriveAgent.Job.LocalMapping.LocalFileManager;

namespace DriveAgent.Job
{
    [DisallowConcurrentExecution]
    public class LocalMapping : IJob
    {
        public class UserInfo
        {
            public string EmpNo { get; set; }
            public string EmpName { get; set; }
            public string AD { get; set; }
            public string DepName { get; set; }
        }

        public class Files
        {
            public string nodeID { get; set; } // 節點id
            public string parentID { get; set; } // 父節點id
            public string fileID { get; set; } // 檔案id
            public string fileName { get; set; } // 檔案名稱
            public string fullPath { get; set; } // 檔案完整路徑
            public string comparePath { get; set; } // 比對路徑
            public DateTime createTime { get; set; }
            public DateTime lastModiDate { get; set; }
            public bool isNewFile { get; set; } // 是否為新檔案
            public string oldFileName { get; set; } // 舊檔案名稱
            public string newFileName { get; set; } // 新檔案名稱
        }

        public class APIResponse
        {
            public int status { get; set; }
            public string message { get; set; }

        }

        public class DownloadAPIResponse
        {
            public int status { get; set; }
            public byte[] fileBytes { get; set; }

        }

        public enum SyncAction
        {
            LocalAdd, // 地端新增檔案
            LocalDelete, // 地端刪除檔案
            LocalRename, // 地端Rename
            CloudAdd, // 雲端新增等案
            CloudDelete, // 雲端刪除檔案
            CloudRename, // 雲端Rename
        }

        public class SyncFileList
        {
            public DateTime lastSyncTime { get; set; }
            public SyncAction action { get; set; }
            public int count { get; set; }
            public List<Files> files { get; set; }
        }

        public async Task Execute(IJobExecutionContext context)
        {
            // 指定路徑 判斷user本機是否存在
            //string localFolderPath = "C:\\MyDrive"; // 之後改成彈性設定參數
            string localFolderPath = InxSetting.ClsCommon.GetIni("Config.ini", "Folder", "MyDriver");
            if (localFolderPath == string.Empty)
            {
                return;
            }
            // 進行本機路徑檢驗，不存在則建立資料夾
            if (!Directory.Exists(localFolderPath))
            {
                // 建立資料夾
                Directory.CreateDirectory(localFolderPath);
                Console.WriteLine("已建立指定路徑的資料夾：" + localFolderPath);
            }

            // 創建 SyncManager 實例並進行後續操作
            SyncManager syncManager = new SyncManager(localFolderPath);

            // 取得當前時間作為同步開始時間
            DateTime syncStartTime = DateTime.Now;

            // 呼叫 雲端地端檔案同步 程式碼
            syncManager.SyncFiles(syncStartTime);

            
            Console.ReadLine();
        }

        /// <summary>
        /// 讀取appsettings取得ApiUrl
        /// </summary>
        /// <returns></returns>
        public static string GetAppUrl()
        {
            IConfigurationRoot configuration = new ConfigurationBuilder()
                    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                    .AddJsonFile("appsettings.json")
                    .Build();

            string appUrl = configuration["ApiUrl"];
            return appUrl;
        }

        /// <summary>
        /// 顯示提示訊息
        /// </summary>
        /// <param name="errorMessage"></param>
        public void ShowMessageInRightBottomCorner(string Msg ,string Title)
        {
            string message = Msg;
            string title = Title;

            // Get the screen size
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;

            // Calculate the position of the message box
            int x = screenWidth - 350;
            int y = screenHeight - 150;

            // Show the message box
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly, true);

            // Move the message box to the right bottom corner
            //MessageBox.Show.Location = new Point(x, y);
        }

        /// <summary>
        /// 地端檔案管理
        /// </summary>
        public class LocalFileManager
        {
            private string localFolderPath;

            // 宣告本機資料夾路徑
            public LocalFileManager(string folderPath)
            {
                localFolderPath = folderPath;
            }

            /// <summary>
            /// 取得本地端資料夾中的所有檔案列表
            /// </summary>
            /// <returns></returns>
            public List<Files> GetLocalFiles()
            {
                // 創建 Files List 物件
                List<Files> fileList = new List<Files>();

                // 呼叫 ProcessFolder 方法，處理地端資料夾
                ProcessFolder(localFolderPath, fileList);

                return fileList;
            }

            /// <summary>
            /// 遞迴處理資料夾
            /// </summary>
            /// <param name="folderPath"></param>
            /// <param name="fileList"></param>
            private void ProcessFolder(string folderPath, List<Files> fileList)
            {
                // 建立指定資料夾的實例
                DirectoryInfo directoryInfo = new DirectoryInfo(folderPath);

                if (directoryInfo.Exists)
                {
                    // 獲取資料夾中的檔案列表
                    FileInfo[] files = directoryInfo.GetFiles();

                    // 使用FileInfo來獲取每個檔案的資訊
                    foreach (FileInfo file in files)
                    {
                        // 過濾真實檔案
                        if (IsRealFile(file))
                        {
                            // 獲取檔案的相關資訊
                            string fileName = file.Name;
                            string fullPath = System.IO.Path.Combine(file.DirectoryName, fileName);
                            DateTime createTime = file.CreationTime;
                            DateTime lastModifyTime = file.LastWriteTime;

                            // 創建 LocalFile 物件，並將檔案資訊添加到 fileList 列表中
                            Files localFile = new Files
                            {
                                fileName = fileName,
                                fullPath = fullPath,
                                createTime = createTime,
                                lastModiDate = lastModifyTime
                            };

                            fileList.Add(localFile);

                            Console.WriteLine($"File Name: {fileName}");
                            Console.WriteLine($"Create Time: {createTime}");
                            Console.WriteLine($"Last Modify Time: {lastModifyTime}");
                            Console.WriteLine();
                        }
                    }

                    // 獲取資料夾中的子資料夾列表
                    DirectoryInfo[] subDirectories = directoryInfo.GetDirectories();

                    foreach (DirectoryInfo subDirectory in subDirectories)
                    {
                        // 遞迴處理子資料夾
                        ProcessFolder(subDirectory.FullName, fileList);
                    }
                }
                else
                {
                    Console.WriteLine("Folder does not exist.");
                }
            }

            /// <summary>
            /// 判斷檔案是否為真實檔案
            /// </summary>
            /// <param name="file"></param>
            /// <returns></returns>
            private bool IsRealFile(FileInfo file)
            {
                string extension = file.Extension;

                // 根據需要自訂判斷暫存檔的條件
                if (extension == ".tmp" || extension == ".bak" || file.Name.StartsWith("~") || file.Name.StartsWith("."))
                {
                    return false;
                }

                return true;
            }

            /// <summary>
            /// 新增地端檔案(雲端下載到地端)
            /// </summary>
            /// <param name="file"></param>
            // 將二進制數據寫入檔案
            public void SaveFile(string filePath, byte[] data)
            {
                try
                {
                    // 獲取 filePath 中的目錄部分
                    string directoryPath = System.IO.Path.GetDirectoryName(filePath);

                    // 檢查路徑是否存在，若不存在則新增資料夾
                    if (!Directory.Exists(directoryPath))
                    {
                        Directory.CreateDirectory(directoryPath);
                    }

                    using (FileStream fs = new FileStream(filePath, FileMode.Create))
                    {
                        fs.Write(data, 0, data.Length); // 將data儲存到FileStream中
                    }
                    //ClsCommon.showMsg($"檔案已存入本機:{filePath}");
                    Console.WriteLine("檔案已存入本機");
                }
                catch (Exception ex)
                {
                    //ClsCommon.showMsg($"儲存檔案時發生錯誤:{ex.Message}");
                    Console.WriteLine("儲存檔案時發生錯誤：" + ex.Message);
                }
            }

            /// <summary>
            /// 刪除地端檔案
            /// </summary>
            /// <param name="filePath"></param>
            public void DeleteLocalFile(string filePath)
            {   
                try
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                        //ClsCommon.showMsg($"檔案已成功刪除:{filePath}");
                        //MessageBox.Show($"檔案已成功刪除:{filePath}", "檔案列表");
                        // 遞迴刪除資料夾
                        //DeleteEmptyParentDirectory(filePath);   
                    }
                    else if (Directory.Exists(filePath))
                    {
                        Directory.Delete(filePath, true);
                        //ClsCommon.showMsg($"資料夾已成功刪除:{filePath}");
                        //MessageBox.Show("資料夾已成功刪除。", "檔案列表");
                        // 遞迴刪除資料夾
                        //DeleteEmptyParentDirectory(filePath);
                    }
                    else
                    {
                        //ClsCommon.showMsg($"檔案已成功刪除:{filePath}");
                        //MessageBox.Show("檔案或資料夾不存在。", "檔案列表");
                    }
                }
                catch (IOException ex)
                {
                    //ClsCommon.showMsg($"刪除檔案時發生錯誤:{ex.Message}");
                    //MessageBox.Show($"刪除檔案時發生錯誤：{ex.Message}", "檔案列表");
                    //Console.WriteLine($"刪除檔案時發生錯誤：{ex.Message}");
                }
                catch (UnauthorizedAccessException ex)
                {
                    //ClsCommon.showMsg($"無權限刪除檔案:{ex.Message}");
                    //MessageBox.Show($"無權限刪除檔案：{ex.Message}", "檔案列表");
                    //Console.WriteLine($"無權限刪除檔案：{ex.Message}");
                }
                catch (Exception ex)
                {
                    //ClsCommon.showMsg($"刪除檔案時發生錯誤:{ex.Message}");
                    //MessageBox.Show($"刪除檔案時發生錯誤：{ex.Message}", "檔案列表");
                    //Console.WriteLine($"刪除檔案時發生錯誤：{ex.Message}");
                }
            }

            /// <summary>
            /// 遞迴處理刪除空資料夾
            /// </summary>
            /// <param name="filePath"></param>
            private void DeleteEmptyParentDirectory(string filePath)
            {
                // 取得父資料夾
                string parentDirectory = Directory.GetParent(filePath).FullName;
                if (parentDirectory.Equals(localFolderPath, StringComparison.OrdinalIgnoreCase))
                {
                    // 如果上層資料夾是最上層的 "C:\MyDrive" 資料夾，則不進行刪除
                    return;
                }

                // 資料夾中存在 檔案、資料夾 
                if (Directory.GetFiles(parentDirectory).Length == 0 &&
                    Directory.GetDirectories(parentDirectory).Length == 0)
                {
                    Directory.Delete(parentDirectory);
                    MessageBox.Show($"上層資料夾已成功刪除：{parentDirectory}", "檔案列表");
                    DeleteEmptyParentDirectory(parentDirectory);
                }
            }

            public enum EnumlocalItems
            {
                File,
                Directory
            }

            /// <summary>
            /// 修改地端資料夾/檔案名稱
            /// </summary>
            /// <param name="actTarget">EnumlocalItems.File or EnumlocalItems.Directory</param>
            /// <param name="oldFilePath"></param>
            /// <param name="newFilePath"></param>
            public void RenameLocalItem(EnumlocalItems actTarget, string oldFilePath, string newFilePath)
            {
                try
                {
                    switch (actTarget)
                    {
                        case EnumlocalItems.File:
                            if (File.Exists(oldFilePath))
                            {
                                File.Move(oldFilePath, newFilePath);
                                Console.WriteLine($"本機資料夾名稱已修改:\n{oldFilePath} → {newFilePath}");
                                //ClsCommon.showMsg($"本機檔案名稱已修改:\n{oldFilePath} → {newFilePath}");
                                //MessageBox.Show($"本機檔案名稱已成功修改:{oldFilePath} > {newFilePath}", "訊息通知");
                            }
                            else
                            {
                                Console.WriteLine("原始檔案不存在。");
                            }
                            break;
                        case EnumlocalItems.Directory:
                            if (Directory.Exists(oldFilePath))
                            {
                                Directory.Move(oldFilePath, newFilePath);
                                Console.WriteLine($"本機資料夾名稱已修改:\n{oldFilePath} → {newFilePath}");
                                //ClsCommon.showMsg($"本機資料夾名稱已修改:\n{oldFilePath} → {newFilePath}");
                                //MessageBox.Show($"本機資料夾名稱已成功修改:{oldFilePath} > {newFilePath}", "訊息通知");
                            }
                            else
                            {
                                Console.WriteLine("原始資料夾不存在。");
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("發生錯誤: " + ex.Message);
                }
            }

            /// <summary>
            /// 更改地端檔案最後修改日期時間
            /// </summary>
            /// <param name="fullPath"></param>
            /// <param name="lastModiDate"></param>
            public void SetLastModiDate(string fullPath, DateTime lastModiDate)
            {
                try
                {
                    // 檢查檔案是否存在
                    if (File.Exists(fullPath))
                    {
                        // 修改檔案的最後修改日期和時間
                        File.SetLastWriteTime(fullPath, lastModiDate);
                        Console.WriteLine("檔案的最後修改日期和時間已修改。");
                    }
                    else
                    {
                        Console.WriteLine("指定的檔案不存在。");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"發生錯誤: {ex.Message}");
                }
            }

            // 檔案監控
            private FileSystemWatcher fileWatcher;
            // 資料夾監控
            private FileSystemWatcher folderWatcher;

            /// <summary>
            /// 開始監控'檔案'Renamed動作
            /// 開始監控'資料夾'Renamed動作
            /// </summary>
            public void StartMonitoring()
            {
                // 指定要監控的資料夾路徑為localFolderPath
                fileWatcher = new FileSystemWatcher(localFolderPath);
                // 包含子資料夾
                fileWatcher.IncludeSubdirectories = true;
                // 只監控"檔案"名稱的變更
                fileWatcher.NotifyFilter = NotifyFilters.FileName;
                // 表示監控所有類型的檔案
                fileWatcher.Filter = "*.*";
                // 訂閱FileSystemWatcher物件的Renamed事件，並指定事件處理方法為OnFileRenamed(更新修改時間)
                fileWatcher.Renamed += OnFileRenamed;
                // 啟動監聽
                fileWatcher.EnableRaisingEvents = true;

                folderWatcher = new FileSystemWatcher(localFolderPath);
                // 包含子資料夾
                folderWatcher.IncludeSubdirectories = true;
                // 只監控"資料夾"名稱的變更
                folderWatcher.NotifyFilter = NotifyFilters.DirectoryName;
                // 表示監控所有類型的檔案
                folderWatcher.Filter = "*";
                // 訂閱FileSystemWatcher物件的Renamed事件，並指定事件處理方法為OnFolderRenamed(更新資料夾內所有檔案的最後修改時間)
                folderWatcher.Renamed += OnFolderRenamed;
                // 啟動監聽
                folderWatcher.EnableRaisingEvents = true;
            }

            /// <summary>
            /// 停止監控資料夾
            /// </summary>
            public void StopMonitoring()
            {
                /*FileSystemWatcher watcher = new FileSystemWatcher(localFolderPath);
                watcher.IncludeSubdirectories = true;
                watcher.EnableRaisingEvents = false;*/
                fileWatcher.EnableRaisingEvents = false;
                folderWatcher.EnableRaisingEvents = false;
            }

            /// <summary>
            /// 獲取到檔案Rename的事件，並重設檔案最後修改時間
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            public void OnFileRenamed(object sender, RenamedEventArgs e)
            {
                try
                {
                    // 取得原始檔案的最後修改時間
                    //DateTime lastWriteTime = File.GetLastWriteTime(e.OldFullPath);
                    DateTime lastWriteTime = DateTime.Now;
                    // 更新原始檔案的最後修改時間
                    File.SetLastWriteTime(e.FullPath, lastWriteTime);

                    Console.WriteLine("檔案名稱已修改: " + e.OldName + " -> " + e.Name);
                    Console.WriteLine("檔案的最後修改時間已成功更新。");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("發生錯誤: " + ex.Message);
                }
            }

            /// <summary>
            /// 獲取到資料夾Rename的事件
            /// 並調用UpdateFolderLastWriteTime方法來更新資料夾內所有檔案的最後修改時間
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            public void OnFolderRenamed(object sender, RenamedEventArgs e)
            {
                try
                {
                    DateTime lastWriteTime = DateTime.Now;
                    UpdateFolderLastWriteTime(e.FullPath, lastWriteTime);

                    Console.WriteLine("資料夾名稱已修改: " + e.OldName + " -> " + e.Name);
                    Console.WriteLine("資料夾內所有檔案的最後修改時間已成功更新。");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("發生錯誤: " + ex.Message);
                }
            }

            /// <summary>
            /// 更新資料夾內所有檔案的最後修改時間
            /// 使用遞迴的方式遍歷資料夾內的所有檔案和子資料夾
            /// </summary>
            /// <param name="folderPath"></param>
            /// <param name="lastWriteTime"></param>
            private void UpdateFolderLastWriteTime(string folderPath, DateTime lastWriteTime)
            {
                DirectoryInfo folder = new DirectoryInfo(folderPath);
                folder.LastWriteTime = lastWriteTime;

                foreach (var file in folder.GetFiles())
                {
                    file.LastWriteTime = lastWriteTime;
                }

                foreach (var subFolder in folder.GetDirectories())
                {
                    UpdateFolderLastWriteTime(subFolder.FullName, lastWriteTime);
                }
            }
        }

        /// <summary>
        /// 雲端檔案管理
        /// </summary>
        public class CloudFileManager
        {
            //private string apiUrl = "http://tinnodrvapitn.cminl.oa/api"; 
            private string apiUrl = InxSetting.ClsCommon.WebApiUrl + "/api";// 取得WebApiUrl
            private HttpClient apiClient; // 添加 HttpClient 成員變數
            string baseFolder;
            public CloudFileManager(string localFolderPath)
            {
                //cloudFolderPath = folderPath;
                baseFolder = localFolderPath;
                HttpClient apiClient = new HttpClient(); // 初始化 HttpClient
            }

            /// <summary>
            /// 取得本機使用者名稱
            /// </summary>
            /// <returns></returns>
            public string GetOsUsername()
            {
                return Environment.UserName;
            }

            /// <summary>
            /// 取得WebDavTicket
            /// </summary>
            /// <param name="account"></param>
            /// <returns></returns>
            public async Task<string> GetWebDavTicket()
            {
                string route = "/WebDav/GetWebDavTicket";
                string account = GetOsUsername();
                string postData = "{\"Account\":\"" + account + "\", \"Password\":\"\"}";
                var responseTicket = "";
                HttpClient client = new HttpClient();
                try
                {
                    using HttpRequestMessage request = new HttpRequestMessage();
                    {
                        request.Method = HttpMethod.Post;
                        request.RequestUri = new Uri(apiUrl + route);

                        request.Content = new StringContent(postData, Encoding.UTF8, "application/json");

                        HttpResponseMessage result = await client.SendAsync(request);
                        if (result.IsSuccessStatusCode)
                        {
                            responseTicket = await result.Content.ReadAsStringAsync();
                            // 移除引號
                            //responseTicket = responseTicket.Trim('"');
                            Console.WriteLine($"GetApiTicket result: {responseTicket}");
                        }
                        else
                        {
                            Console.WriteLine($"GetApiTicket request failed with status code: {result.StatusCode}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle any exceptions
                    Console.WriteLine("Error: " + ex.Message);
                }

                return responseTicket;
            }

            /// <summary>
            /// 解析Ticket
            /// </summary>
            /// <returns></returns>
            public async Task<UserInfo> GetUserInfo(string filter)
            {
                string route = $"/Employees?filter={filter}";
                var webDavTicket = await GetWebDavTicket();
                HttpClient client = new HttpClient();
                
                try
                {
                    using HttpRequestMessage request = new HttpRequestMessage();
                    {
                        request.Method = HttpMethod.Get;
                        request.RequestUri = new Uri(apiUrl + route);

                        // 指定header ticket 
                        request.Headers.Add("ticket", webDavTicket);

                        HttpResponseMessage result = await client.SendAsync(request);
                        result.EnsureSuccessStatusCode();
                        var response = await result.Content.ReadAsStringAsync();
                        List<UserInfo> userInfoList = JsonConvert.DeserializeObject<List<UserInfo>>(response);
                        if (userInfoList.Count > 0)
                        {
                            return userInfoList[0];
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle any exceptions
                    Console.WriteLine("Error: " + ex.Message);
                    return null;
                }
            }
            /// <summary>
            /// 以不同模式取得雲端MyDrive檔案列表
            /// </summary>
            /// <param name="mode"> Full | NewFiles | Rename | RenameFolder </param>
            /// <returns></returns>
            public async Task<List<Files>> GetMyDriveCloudItems(string mode, DateTime syncStartTime)
            {
                string route = "/LocalMapping/GetMyDriveCloudItems";
                var webDavTicket = await GetWebDavTicket();
                HttpClient client = new HttpClient();
                try
                {
                    using HttpRequestMessage request = new HttpRequestMessage();
                    {
                        request.Method = HttpMethod.Post;
                        request.RequestUri = new Uri(apiUrl + route);

                        // 將資料物件序列化成 JSON 字串
                        var jsonData = JsonConvert.SerializeObject(new { SyncStartTime = syncStartTime.ToString("yyyy/MM/dd HH:mm:ss"), Mode = mode, BaseFolder = baseFolder });
                        // 指定body
                        request.Content = new StringContent(jsonData, Encoding.UTF8, "application/json");
                        // 指定header ticket 
                        request.Headers.Add("ticket", webDavTicket);

                        HttpResponseMessage result = await client.SendAsync(request);
                        if (result.IsSuccessStatusCode)
                        {
                            string responseContent = await result.Content.ReadAsStringAsync();
                            List<Files> files = JsonConvert.DeserializeObject<List<Files>>(responseContent);
                            foreach (Files file in files)
                            {
                                file.comparePath = System.IO.Path.GetDirectoryName(file.fullPath);
                                //file.comparePath = file.fullPath.Substring(file.fullPath.IndexOf("MyDrive") + "MyDrive".Length + 1);
                            }
                            return files;
                        }
                        else
                        {
                            Console.WriteLine($"GetMyDriveCloudItems request failed with status code: {result.StatusCode}");
                        }
                    }
                }
                catch (Exception ex)
                {

                    Console.WriteLine("Error: " + ex.Message);
                }

                return null;
            }

            /// <summary>
            /// 取得最後同步時間
            /// </summary>
            /// <returns></returns>
            public async Task<DateTime> GetLastSyncTime()
            {
                string route = "/LocalMapping/GetLastSyncTime";
                string responseContent = "";
                DateTime dateTime = DateTime.Now;
                var webDavTicket = await GetWebDavTicket();
                HttpClient client = new HttpClient();
                try
                {
                    using HttpRequestMessage request = new HttpRequestMessage();
                    {
                        request.Method = HttpMethod.Get;
                        request.RequestUri = new Uri(apiUrl + route);

                        // 指定header ticket 
                        request.Headers.Add("ticket", webDavTicket);

                        HttpResponseMessage result = await client.SendAsync(request);
                        if (result.IsSuccessStatusCode)
                        {
                            responseContent = await result.Content.ReadAsStringAsync();
                            responseContent = responseContent.Replace("\"", ""); // 刪除引號
                            dateTime = DateTime.ParseExact(responseContent, "yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
                            //return dateTime;
                        }
                        else
                        {
                            Console.WriteLine($"GetLastSyncTime request failed with status code: {result.StatusCode}");
                        }
                    }
                }
                catch (Exception ex)
                {

                    Console.WriteLine("Error: " + ex.Message);
                }
                return dateTime;  // 傳回預設值
            }

            /// <summary>
            /// 設定最後同步時間
            /// </summary>
            /// <param name="SyncTime">日期時間格式: yyyy/mm/dd h:m:s</param>
            /// <returns></returns>
            public async Task<DateTime> SetLastSyncTime(DateTime syncTime)
            {
                string route = "/LocalMapping/SetLastSyncTime";
                var webDavTicket = await GetWebDavTicket();
                HttpClient client = new HttpClient();
                try
                {
                    using HttpRequestMessage request = new HttpRequestMessage();
                    {
                        request.Method = HttpMethod.Post;
                        request.RequestUri = new Uri(apiUrl + route);

                        // 將資料物件序列化成 JSON 字串
                        var jsonData = JsonConvert.SerializeObject(new { SyncTime = syncTime.ToString("yyyy/MM/dd HH:mm:ss") });
                        // 指定body
                        request.Content = new StringContent(jsonData, Encoding.UTF8, "application/json");
                        // 指定header ticket 
                        request.Headers.Add("ticket", webDavTicket);

                        HttpResponseMessage result = await client.SendAsync(request);
                        if (result.IsSuccessStatusCode)
                        {
                            string responseContent = await result.Content.ReadAsStringAsync();
                            DateTime dateTime;
                            if (DateTime.TryParse(responseContent, out dateTime))
                            {
                                return dateTime;
                            }
                            else
                            {
                                //throw new Exception("Failed to parse DateTime.");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"SetLastSyncTime request failed with status code: {result.StatusCode}");
                        }
                    }
                }
                catch (Exception ex)
                {

                    Console.WriteLine("Error: " + ex.Message);
                }

                return DateTime.Now;
            }

            /// <summary>
            /// 新增雲端資料夾
            /// </summary>
            /// <param name="parentId">父節點ID</param>
            /// <param name="folderName">資料夾名稱</param>
            /// <returns></returns>
            public async Task<string> AddCloudFolder(string parentId, string folderName)
            {
                string route = "/Interface/AddFolder";
                var webDavTicket = await GetWebDavTicket();
                string addFolderResponse = "";
                HttpClient client = new HttpClient();
                try
                {
                    using HttpRequestMessage request = new HttpRequestMessage();
                    {
                        request.Method = HttpMethod.Post;
                        request.RequestUri = new Uri(apiUrl + route);

                        // 將資料物件序列化成 JSON 字串
                        var jsonData = JsonConvert.SerializeObject(new { parentId = parentId, FolderName = folderName });
                        // 指定body
                        request.Content = new StringContent(jsonData, Encoding.UTF8, "application/json");
                        // 指定header ticket 
                        request.Headers.Add("ticket", webDavTicket);

                        // header加入ticket 
                        client.DefaultRequestHeaders.Add("ticket", webDavTicket);

                        // 發送 POST 請求
                        HttpResponseMessage result = await client.SendAsync(request);
                        if (result.IsSuccessStatusCode)
                        {
                            // 讀取回應內容
                            addFolderResponse = await result.Content.ReadAsStringAsync();
                            return addFolderResponse;
                        }
                        else
                        {
                            Console.WriteLine($"addFolderResponse API request failed with status code: {result.StatusCode}");
                            return $"addFolderResponse API request failed with status code: {result.StatusCode}";
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle any exceptions
                    Console.WriteLine("Error: " + ex.Message);
                    return "Error: " + ex.Message;
                }
            }

            /// <summary>
            /// 新增雲端檔案(地端上傳雲端)
            /// </summary>
            /// <param name="parentId">資料夾父結點ID</param>
            /// <param name="overwrite">是否覆寫</param>
            /// <param name="filePath">檔案路徑</param>
            /// <returns></returns>
            public async Task<APIResponse> UploadCloudFile(string parentId, bool overwrite, string filePath, DateTime lastModiDate)
            {
                string route = "/Interface/UploadFiles";
                var webDavTicket = await GetWebDavTicket();
                string responseContent = "";
                int bufSize = 1024 * 1024;   //1MB

                APIResponse uploadResponse = new APIResponse();

                HttpClient client = new HttpClient();
                try
                {
                    // 加入form-data
                    using MultipartFormDataContent content = new MultipartFormDataContent();
                    {
                        // 添加 parentId 和 overwrite 參數
                        content.Add(new StringContent(parentId), "parentId");
                        content.Add(new StringContent(overwrite.ToString()), "overwrite");
                        content.Add(new StringContent(lastModiDate.ToString("yyyy/MM/dd HH:mm:ss"), Encoding.UTF8), "lastmodifydate");
                        //content.Add(new StringContent(file.LastModiDate.ToString("yyyy/MM/dd HH:mm:ss"), Encoding.UTF8), "lastmodifydate");
                        /* 讀取檔案內容
                        //byte[] fileBytes = await File.ReadAllBytesAsync(filePath);
                        // 建立檔案內容的 StreamContent
                        //StreamContent fileContent = new StreamContent(new MemoryStream(fileBytes));
                        // 設定檔案內容的 MediaTypeHeaderValue
                        //fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");*/

                        // 儲存路徑 格式要符合f1/f2/aa.txt
                        string savePath = filePath.Substring(11).Replace('\\', '/');
                        FileStream fStream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);
                        // 添加檔案內容到 multipart form data
                        content.Add(new StreamContent(fStream, bufSize), savePath, filePath);

                        // header加入ticket 
                        client.DefaultRequestHeaders.Add("ticket", webDavTicket);

                        // 發送 POST 請求
                        HttpResponseMessage result = await client.PostAsync(apiUrl + route, content);
                        if (result.IsSuccessStatusCode)
                        {
                            // 讀取回應內容 
                            responseContent = await result.Content.ReadAsStringAsync();
                            //ClsCommon.showMsg($"檔案已成功上傳至InnoDrive:\n{filePath}");

                            uploadResponse.status = (int)result.StatusCode;
                            uploadResponse.message = filePath;

                            return uploadResponse;
                        }
                        else
                        {
                            uploadResponse.status = (int)result.StatusCode;
                            uploadResponse.message = $"檔案上傳InnoDrive失敗:{result.StatusCode}";

                            Console.WriteLine($"API request failed with status code: {result.StatusCode}");
                            //ClsCommon.showMsg($"檔案上傳InnoDrive失敗:{result.StatusCode}");
                            return uploadResponse;
                        }
                    }
                }
                catch (Exception ex)
                {
                    uploadResponse.status = -1;
                    uploadResponse.message = $"Error: " + ex.Message;
                    // Handle any exceptions
                    Console.WriteLine("Error: " + ex.Message);
                    //ClsCommon.showMsg($"檔案上傳InnoDrive失敗:{ex.Message}");
                    return uploadResponse;
                }

                //return responseContent;
            }

            /// <summary>
            /// 取得下載檔案Ticket
            /// </summary>
            /// <param name="fileID"></param>
            /// <returns></returns>
            public async Task<APIResponse> GetDownloadFileTicket(string fileID)
            {
                string route = "/Interface/GetDownloadFileTicket";
                var webDavTicket = await GetWebDavTicket();
                APIResponse response = new APIResponse();
                HttpClient client = new HttpClient();
                try
                {
                    using HttpRequestMessage request = new HttpRequestMessage();
                    {
                        request.Method = HttpMethod.Post;
                        request.RequestUri = new Uri(apiUrl + route);

                        // 將資料物件序列化成 JSON 字串
                        var jsonData = JsonConvert.SerializeObject(new { FileID = fileID });
                        // 指定body
                        request.Content = new StringContent(jsonData, Encoding.UTF8, "application/json");
                        // 指定header ticket 
                        request.Headers.Add("ticket", webDavTicket);

                        // header加入ticket 
                        client.DefaultRequestHeaders.Add("ticket", webDavTicket);

                        // 發送 POST 請求
                        HttpResponseMessage result = await client.SendAsync(request);
                        if (result.IsSuccessStatusCode)
                        {
                            // 讀取回應內容
                            response.status = (int)result.StatusCode;
                            response.message = await result.Content.ReadAsStringAsync();
                            return response;
                        }
                        else
                        {
                            response.status = (int)result.StatusCode;
                            response.message = await result.Content.ReadAsStringAsync();
                            return response;
                            //Console.WriteLine($"GetDownloadFileTicket API request failed with status code: {result.StatusCode}");
                            //return $"GetDownloadFileTicket API request failed with status code: {result.StatusCode}";
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle any exceptions
                    Console.WriteLine("Error: " + ex.Message);
                    return new APIResponse
                    {
                        status = -1, // 或者根據需要設定適當的錯誤狀態碼
                        message = "Error: " + ex.Message
                    };
                }

                //return null;
            }

            /// <summary>
            /// 下載雲端檔案到地端
            /// </summary>
            public async Task<APIResponse> DownloadCloudFile(string fileID)
            {
                // 取得DownloadTicket
                APIResponse response = await GetDownloadFileTicket(fileID);
                APIResponse downloadResponse = new APIResponse();
                if (response.status == 200)
                {
                    // 在成功取得 Ticket 之後才執行下載 API 程式碼
                    // 下載檔案API程式碼
                    string route = "/Interface/DownloadFile";
                    var webDavTicket = await GetWebDavTicket();

                    byte[] downloadFileBytes = null; // 儲存下載的檔案位元組陣列
                    //string downloadFileResponse = "";
                    HttpClient client = new HttpClient();
                    try
                    {
                        using HttpRequestMessage request = new HttpRequestMessage();
                        {
                            request.Method = HttpMethod.Get;
                            request.RequestUri = new Uri(apiUrl + route);

                            // 指定header ticket 
                            request.Headers.Add("ticket", webDavTicket);

                            // 加入Params
                            var queryParams = new Dictionary<string, string>()
                            {
                                { "DownloadTicket", response.message.Trim('"') }
                            };
                            request.RequestUri = new Uri(QueryHelpers.AddQueryString(request.RequestUri.ToString(), queryParams));

                            // 發送 Get 請求
                            HttpResponseMessage result = await client.SendAsync(request);
                            if (result.IsSuccessStatusCode)
                            {
                                // 讀取回應內容並轉換為位元組陣列
                                downloadFileBytes = await result.Content.ReadAsByteArrayAsync();

                                // 將位元組陣列轉換為 Base64 字串
                                string base64String = Convert.ToBase64String(downloadFileBytes);

                                downloadResponse.status = (int)result.StatusCode;
                                downloadResponse.message = base64String;

                                //ClsCommon.showMsg($"檔案已成功下載至本機MyDrive");
                                return downloadResponse;
                            }
                            else
                            {
                                downloadResponse.status = (int)result.StatusCode;
                                downloadResponse.message = $"GetDownloadFile API request failed with status code: {result.StatusCode}";
                                //ClsCommon.showMsg($"檔案下載至本機MyDrive失敗:{result.StatusCode}");
                                return downloadResponse;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        downloadResponse.status = -1;
                        downloadResponse.message = $"Error: " + ex.Message;
                        //ClsCommon.showMsg($"檔案已下載至本機MyDrive失敗:{ex.Message}");
                        return downloadResponse;
                    }
                }
                else
                {
                    downloadResponse.status = response.status;
                    downloadResponse.message = "Failed to get download file ticket.";
                    //ClsCommon.showMsg($"檔案已下載至本機MyDrive失敗:Failed to get download file ticket.");
                    // 處理無法取得 Ticket 的情況
                    return downloadResponse;
                }
            }

            /// <summary>
            /// 刪除雲端檔案/資料夾
            /// </summary>
            /// <param name="itemID">檔案ID、資料夾節點ID</param>
            /// <param name="parentID">父節點ID</param>
            public async Task<APIResponse> DeleteCloudFile(string itemID, string parentID)
            {
                string route = "/Interface/DeleteFile";
                var webDavTicket = await GetWebDavTicket();
                string deleteFileResponse = "";
                APIResponse delFileResponse = new APIResponse();
                HttpClient client = new HttpClient();
                try
                {
                    using HttpRequestMessage request = new HttpRequestMessage();
                    {
                        request.Method = HttpMethod.Post;
                        request.RequestUri = new Uri(apiUrl + route);

                        // 將資料物件序列化成 JSON 字串
                        var jsonData = JsonConvert.SerializeObject(new { fileID = itemID });
                        // 指定body
                        request.Content = new StringContent(jsonData, Encoding.UTF8, "application/json");
                        // 指定header ticket 
                        request.Headers.Add("ticket", webDavTicket);

                        // header加入ticket 
                        client.DefaultRequestHeaders.Add("ticket", webDavTicket);

                        // 發送 POST 請求
                        HttpResponseMessage result = await client.SendAsync(request);
                        if (result.IsSuccessStatusCode)
                        {
                            // 讀取回應內容
                            deleteFileResponse = $"DeleteFileFile API request successed with status code: {result.StatusCode}";

                            delFileResponse.status = (int)result.StatusCode;
                            delFileResponse.message = $"DeleteFileFile API request successed with status code: {result.StatusCode}";
                            //MessageBox.Show($"檔案已成功刪除:{itemID}", "檔案列表");

                            // 檢查該item的"父資料夾"有無其他資料
                            /*var parentGetItemResponse = await GetItems(parentID);

                            // 檢查該item的"父資料夾"無其他資料
                            if (parentGetItemResponse.status == 0)
                            {
                                // 繼續往上檢查 = > 取得parentID的parentID
                                Files itemInfo = await DriveFolders(parentID);

                                // 若parentID不包含其他資料夾或檔案，則刪除該節點
                                await DeleteCloudFile(parentID, itemInfo.parentID);

                            }*/

                            return delFileResponse;
                        }
                        else
                        {
                            delFileResponse.status = (int)result.StatusCode;
                            delFileResponse.message = $"DeleteFileFile API request failed with status code: {result.StatusCode}";
                            Console.WriteLine($"DeleteFileFile API request failed with status code: {result.StatusCode}");
                            return delFileResponse;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Handle any exceptions
                    delFileResponse.status = -1;
                    delFileResponse.message = $"DeleteFileFile API request failed with status code: {ex.Message}";
                    Console.WriteLine("Error: " + ex.Message);
                    return delFileResponse;
                }
            }

            /// <summary>
            /// 檢查parentId底下有無包含資料夾/檔案 包含回傳1 不包含回傳0 其他錯誤狀況回傳-1
            /// </summary>
            /// <param name="parentId"></param>
            /// <returns></returns>
            public async Task<APIResponse> GetItems(string parentId)
            {
                string route = "/DriveItems/GetItems";
                var webDavTicket = await GetWebDavTicket();
                APIResponse response = new APIResponse();
                string getItemResponse = "";
                HttpClient client = new HttpClient();
                try
                {
                    using HttpRequestMessage request = new HttpRequestMessage();
                    {
                        request.Method = HttpMethod.Get;
                        request.RequestUri = new Uri(apiUrl + route);

                        // 指定header ticket 
                        request.Headers.Add("ticket", webDavTicket);

                        // 加入Params
                        var queryParams = new Dictionary<string, string>()
                            {
                                { "parentId", parentId.Trim('"') }
                            };
                        request.RequestUri = new Uri(QueryHelpers.AddQueryString(request.RequestUri.ToString(), queryParams));

                        // 發送 Get 請求
                        HttpResponseMessage result = await client.SendAsync(request);
                        //getItemResponse = getItemResponse.Trim(); // Trim whitespace characters
                        if (result.IsSuccessStatusCode)
                        {
                            // 讀取回應內容
                            getItemResponse = await result.Content.ReadAsStringAsync();
                            if (getItemResponse.Trim() == "[]")
                            {
                                response.status = 0;
                                response.message = "The parentID does not contain any values.";

                                // 刪除沒有包含資料夾或檔案的節點
                                // await DeleteCloudFile(parentId); // 呼叫刪除檔案的方法

                                return response;
                            }
                            else
                            {
                                response.status = 1;
                                response.message = "The parentID contains values.";
                                return response;
                            }
                        }
                        else
                        {
                            response.status = -1;
                            response.message = $"GetItems API request failed with status code: {result.StatusCode}";
                            return response;
                        }
                    }
                }
                catch (Exception ex)
                {
                    return new APIResponse
                    {
                        status = -1, // 或者根據需要設定適當的錯誤狀態碼
                        message = "Error: " + ex.Message
                    };
                }
            }

        }

        /// <summary>
        /// 同步管理
        /// </summary>
        public class SyncManager
        {
            private LocalFileManager localFileManager;
            private CloudFileManager cloudFileManager;
            //private AgentFrom agentForm = new AgentFrom();
            //AgentFrom agentForm = new AgentFrom();
            public SyncManager(string localFolderPath)
            {
                localFileManager = new LocalFileManager(localFolderPath);
                cloudFileManager = new CloudFileManager(localFolderPath);
            }

            // 上一次同步資訊
            public List<SyncFileList> lastSyncFileLists = null;

            /// <summary>
            /// 同步主程式
            /// </summary>
            /// <param name="syncStartTime"></param>
            public async void SyncFiles(DateTime syncStartTime)
            {
                // 取得本機使用者名稱
                string userName = cloudFileManager.GetOsUsername();

                // 取得當前時間作為同步開始時間
                //DateTime syncStartTime = DateTime.Now;

                // 取得最後同步時間
                DateTime lastSyncTime = await cloudFileManager.GetLastSyncTime();

                // [測試]地:取得地端檔案列表
                List<Files> localFiles = localFileManager.GetLocalFiles();

                // [測試]雲:取得雲端檔案列表
                List<Files> cloudFiles = await cloudFileManager.GetMyDriveCloudItems("NewFiles", syncStartTime);

                // 監控本地檔案/資料夾'rename'動作
                localFileManager.StartMonitoring();

                /*正式呼叫同步檔案function*/
                await RenameLocalFiles(syncStartTime); // [地端 Rename 檔案]
                await RenameLocalFolders(syncStartTime); // [地端 Rename 資料夾]
                await DownloadToLocalStorage(lastSyncTime, syncStartTime); // [地端新增檔案] 從雲端下載到地端
                await UploadToCloud(lastSyncTime, syncStartTime); // [雲端新增資料夾檔案] 上傳到雲端 
                await DeleteLocalFiles(lastSyncTime, syncStartTime); // [地端刪除檔案]
                await DeleteCloudFiles(lastSyncTime, syncStartTime); // [雲端刪除檔案]

                // "同步開始時間"設置為"最後同步時間"
                DateTime setLastSyncTimeResult = await cloudFileManager.SetLastSyncTime(syncStartTime);

                // 取得最後同步資訊
                Console.WriteLine(syncFileLists);
                if(syncFileLists.Count > 0)
                {
                    string syncFileListMsg = string.Join(
                        Environment.NewLine, 
                        syncFileLists.Select
                        (
                            item =>
                            {
                                string actionDescription = GetSyncActionDescription(item.action);
                                return $"執行動作 : {actionDescription}[{item.count}]";
                            } 
                        )
                    );
                    ClsCommon.showMsg($"同步完成\n{syncFileListMsg}");
                }

                // 如何保留上一次同步的資訊?有再存
                if( syncFileLists.Count > 0)
                {
                    //LastSyncFileListForm.getLastSyncFileDataList(syncFileLists);
                    //lastSyncFileLists = syncFileLists;
                    SaveLastSyncFileList(syncFileLists);
                }

            }

            // 檔案同步動作數量列表
            List<SyncFileList> syncFileLists = new List<SyncFileList>();
            /// <summary>
            /// 蒐集檔案同步動作
            /// </summary>
            /// <param name="action"></param>
            /// <param name="count"></param>
            /// <returns></returns>
            public async Task FileChangeSummary(DateTime syncTime,SyncAction action,int count,List<Files> files=null)
            {
                SyncFileList syncFileList = new SyncFileList();
                syncFileList.lastSyncTime = syncTime;
                syncFileList.action = action;
                syncFileList.count = count;
                syncFileList.files = files;
                syncFileLists.Add(syncFileList);
            }

            /// <summary>
            /// 據 SyncAction 枚舉值顯示對應的中文訊息
            /// </summary>
            /// <param name="action"></param>
            /// <returns></returns>
            private static string GetSyncActionDescription(SyncAction action)
            {
                switch (action)
                {
                    case SyncAction.LocalAdd:
                        return "地端新增檔案";
                    case SyncAction.LocalDelete:
                        return "地端刪除檔案";
                    case SyncAction.LocalRename:
                        return "地端 Rename";
                    case SyncAction.CloudAdd:
                        return "雲端新增檔案";
                    case SyncAction.CloudDelete:
                        return "雲端刪除檔案";
                    case SyncAction.CloudRename:
                        return "雲端 Rename";
                    default:
                        return action.ToString();
                }
            }

            private static string SYNC_FILE_LIST_CACHE_FILE = @"./syncFileListsCache.json";

            /// <summary>
            /// 儲存最後同步資訊
            /// </summary>
            /// <param name="syncFileLists"></param>
            private static void SaveLastSyncFileList(List<SyncFileList> syncFileLists)
            {
                if(syncFileLists != null)
                {
                    try
                    {
                        // 儲存位置改變
                        //string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DriveAgent");
                        string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SyncData");
                        Directory.CreateDirectory(folderPath);
                        string filePath = Path.Combine(folderPath, "syncFileListsCache.json");
                        string json = JsonConvert.SerializeObject(syncFileLists);
                        File.WriteAllText(filePath, json);

                        /*string json = JsonConvert.SerializeObject(syncFileLists);
                        File.WriteAllText(SYNC_FILE_LIST_CACHE_FILE, json);*/
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error saving syncFileListsCache.json: {ex.Message}");
                    }
                }
            }

            /// <summary>
            /// 取得最後同步資訊
            /// </summary>
            /// <returns></returns>
            public async Task<List<SyncFileList>> LoadLastSyncFileList()
            {
                /*if (File.Exists(SYNC_FILE_LIST_CACHE_FILE))
                {
                    //string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DriveAgent");
                    string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SyncData");
                    Directory.CreateDirectory(folderPath);
                    string filePath = Path.Combine(folderPath, "syncFileListsCache.json");

                    string json = File.ReadAllText(filePath);
                    List<SyncFileList> syncFileListCache = JsonConvert.DeserializeObject<List<SyncFileList>>(json);
                    return syncFileListCache;
                }
                else
                {
                    return null;
                }*/

                // 還沒同步前先提示無同步資料 不能讓他load資料 
                string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SyncData");
                Directory.CreateDirectory(folderPath);
                string filePath = Path.Combine(folderPath, "syncFileListsCache.json");

                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    List<SyncFileList> syncFileListCache = JsonConvert.DeserializeObject<List<SyncFileList>>(json);
                    return syncFileListCache;
                }
                else
                {

                    return null;
                }
            }

            /// <summary>
            /// 用於比較Files物件之間的相等性 
            /// </summary>
            public class FilesComparer : IEqualityComparer<Files>
            {
                /// <summary>
                /// 判斷兩個Files物件是否相等
                /// </summary>
                /// <param name="x"></param>
                /// <param name="y"></param>
                /// <returns></returns>
                public bool Equals(Files x, Files y)
                {
                    if (x == null && y == null)
                        return true;
                    else if (x == null || y == null)
                        return false;
                    else
                        return x.fullPath == y.fullPath;
                }

                /// <summary>
                /// 獲取Files物件的雜湊碼
                /// </summary>
                /// <param name="obj"></param>
                /// <returns></returns>
                public int GetHashCode(Files obj)
                {
                    return obj.fullPath.GetHashCode();
                }

                /// <summary>
                /// 取得兩陣列相同的值，並保有雲端的nodeID、fileID list1 = local  ，  list2 = cloud
                /// </summary>
                /// <param name="list1"></param>
                /// <param name="list2"></param>
                /// <returns></returns>
                public List<Files> GetCommonFilePaths(List<Files> list1, List<Files> list2, string mode)
                {
                    List<Files> commonFiles = null;

                    // 下載檔案 保留雲端資訊 雲端修改日期>地端修改日期
                    if (mode == "greater")
                    {
                        /*commonFiles = list2.Intersect(list1, new FilesComparer()).ToList();
                        commonFiles = commonFiles.Where(file =>
                            list1.FirstOrDefault(f => f.fullPath == file.fullPath)?.lastModiDate <
                            list2.FirstOrDefault(f => f.fullPath == file.fullPath)?.lastModiDate)
                            .ToList();*/

                        var filteredFiles = from localFile in list1
                                            join cloudFile in list2 on localFile.fullPath equals cloudFile.fullPath
                                            where TruncateMilliseconds(cloudFile.lastModiDate) > TruncateMilliseconds(localFile.lastModiDate)
                                            select cloudFile;

                        commonFiles = filteredFiles.ToList();

                        return commonFiles;
                    }
                    else if (mode == "less")  // 上傳檔案 保留地端資訊 雲端修改日期<地端修改日期
                    {
                        var filteredFiles = from localFile in list1
                                            join cloudFile in list2 on localFile.fullPath equals cloudFile.fullPath
                                            where TruncateMilliseconds(cloudFile.lastModiDate) < TruncateMilliseconds(localFile.lastModiDate)
                                            select localFile;

                        commonFiles = filteredFiles.ToList();
                        /*commonFiles = list1.Intersect(list2, new FilesComparer()).ToList();
                        commonFiles = commonFiles.Where(file =>
                            list1.FirstOrDefault(f => f.fullPath == file.fullPath)?.lastModiDate >
                            list2.FirstOrDefault(f => f.fullPath == file.fullPath)?.lastModiDate)
                            .ToList();*/

                        return commonFiles;
                    }

                    return commonFiles;
                }

                /// <summary>
                /// 將DateTime物件的毫秒數設為0
                /// </summary>
                /// <param name="dateTime"></param>
                /// <returns></returns>
                private DateTime TruncateMilliseconds(DateTime dateTime)
                {
                    return new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, dateTime.Hour, dateTime.Minute, dateTime.Second);
                }

            }

            /// <summary>
            ///  【雲端→地端】 Download 到地端 [地端新增檔案]
            /// </summary>
            public async Task DownloadToLocalStorage(DateTime lastSyncTime, DateTime syncStartTime)
            {
                // 宣告FilesComparer 比對用class
                FilesComparer comparer = new FilesComparer();

                // 取得地端檔案列表
                List<Files> localFiles = localFileManager.GetLocalFiles();

                // 取得雲端檔案列表 (and AccessLog.Action in ('Upload', 'UploadOverwrite', 'Recovery', 'Rename', 'Update') and logtime > LastSyncTime)
                List<Files> cloudFiles = await cloudFileManager.GetMyDriveCloudItems("NewFiles", syncStartTime);

                /* 雲端有 地端沒有 的檔案 union 雲端地端都有且 雲端最後修改日期 > 地端檔案最後修改日期 */

                // [比對]雲端有 地端沒有 的檔案(雲端後來新增的檔案 NewFiles)
                List<Files> extraCloudFiles = cloudFiles.Except(localFiles, new FilesComparer()).ToList();

                // [比對]雲端地端都有的檔案(包含nodeID、fileID) 且 雲端最後修改日期>地端檔案最後修改日期 (雲端後來編輯的檔案 => 怎麼測?)
                List<Files> commonFiles = comparer.GetCommonFilePaths(localFiles, cloudFiles, "greater");

                // 合併兩種情況的檔案列表FileList
                List<Files> combinedList = extraCloudFiles.Concat(commonFiles).ToList();

                // 檔案異動數量
                int syncFileCount = 0;
                // 檢查 combinedList 中的元素數量是否大於 0
                if (combinedList.Count > 0)
                {
                    foreach (Files file in combinedList)
                    {
                        /* Download 雲端檔案到地端*/
                        APIResponse fileDownloadResult = await cloudFileManager.DownloadCloudFile(file.fileID);

                        if (fileDownloadResult.status == 200)
                        {
                            // Base64解碼
                            byte[] fileData = Convert.FromBase64String(fileDownloadResult.message);

                            // 儲存到地端指定資料夾路徑
                            localFileManager.SaveFile(file.fullPath, fileData);

                            // 設定地端檔案的最後修改日期與雲端檔案的最後修改日期一致
                            localFileManager.SetLastModiDate(file.fullPath, file.lastModiDate);

                            // 檔案數量
                            syncFileCount++;

                            // 顯示進度訊息
                            //ClsCommon.showHint($"InnoDrive:Download到地端 {syncFileCount}/{combinedList.Count}");
                            //agentForm.ShowCustomMessage($"InnoDrive:Download到地端 {syncFileCount}/{combinedList.Count}",2000);
                            CustomMessageBox.ShowCustomMessage($"InnoDrive:Download到地端 {syncFileCount}/{combinedList.Count}", 3000);
                        }
                        else
                        {
                            MessageBox.Show($"檔案未成功下載到本地", "錯誤提示");
                            // 中斷執行
                            return;
                        }
                    }
                }
                else
                {
                    // 中斷執行
                    return;
                }

                if (syncFileCount > 0)
                {
                    // 紀錄檔案異動動作、總數量
                    FileChangeSummary(syncStartTime,SyncAction.LocalAdd, syncFileCount, combinedList);
                }
            }

            /// <summary>
            /// 【雲端→地端】 刪除地端檔案 [地端刪除檔案]
            /// </summary>
            public async Task DeleteLocalFiles(DateTime lastSyncTime, DateTime syncStartTime)
            {
                // 取得地端檔案列表
                List<Files> localFiles = localFileManager.GetLocalFiles();

                // 取得雲端檔案列表
                List<Files> cloudFiles = await cloudFileManager.GetMyDriveCloudItems("Full", syncStartTime);

                // 取得最後同步時間
                //DateTime lastSyncTime = await cloudFileManager.GetLastSyncTime();

                /* 找出地端有 而雲端沒有 的檔案 and 地端建立日期<LastSyncTime 或 地端最後修改日期<LastSyncTime (雲端刪除檔案) */

                // [比對]找出地端有 而雲端沒有 的檔案 
                List<Files> extraLocalFiles = localFiles.Except(cloudFiles, new FilesComparer()).ToList();

                // 檔案異動數量
                int syncFileCount = 0;

                // 檢查 combinedList 中的元素數量是否大於 0
                if (extraLocalFiles.Count > 0)
                {
                    foreach (Files file in extraLocalFiles)
                    {
                        /* [比對]地端建立日期<LastSyncTime(|| file.createTime < lastSyncTime) and 地端最後修改日期<LastSyncTime(已測試) ---------> 但rename地端檔案，檔案會被刪除*/
                        if (file.lastModiDate < lastSyncTime && file.createTime < lastSyncTime)
                        {
                            /* 刪除地端檔案*/
                            localFileManager.DeleteLocalFile(file.fullPath);

                            syncFileCount++;

                            // 顯示進度訊息
                            ClsCommon.showHint($"InnoDrive:地端刪除檔案 {syncFileCount}/{extraLocalFiles.Count}");
                            CustomMessageBox.ShowCustomMessage($"InnoDrive:地端刪除檔案 {syncFileCount}/{extraLocalFiles.Count}", 3000);
                        }
                    }
                }
                else
                {
                    // 中斷執行
                    return;
                }

                if (syncFileCount > 0)
                {
                    // 紀錄檔案異動動作、總數量
                    FileChangeSummary(syncStartTime,SyncAction.LocalDelete, syncFileCount, extraLocalFiles);
                }
            }

            /// <summary>
            /// 【地端→雲端】 上傳到雲端 [雲端新增資料夾檔案]
            /// </summary>
            public async Task UploadToCloud(DateTime lastSyncTime, DateTime syncStartTime)
            {
                // 保存已经創建過的資料夾及其對應的文件夹ID
                Dictionary<string, string> folderDictionary = new Dictionary<string, string>();
                List<Files> folderList;

                // 宣告FilesComparer 比對用class
                FilesComparer comparer = new FilesComparer();

                // 取得地端檔案列表
                List<Files> localFiles = localFileManager.GetLocalFiles();

                // 取得雲端檔案列表
                List<Files> cloudFiles = await cloudFileManager.GetMyDriveCloudItems("Full", syncStartTime);

                // 宣告雲端儲存節點ID
                
                string ad = cloudFileManager.GetOsUsername();
                UserInfo userInfo = await cloudFileManager.GetUserInfo(ad);
                string nodeId = $"DRV-{userInfo.EmpNo}";

                /* 遍歷雲端檔案列表，提取檔案路徑中的資料夾名稱
                foreach (var file in cloudFiles)
                {
                    // 取得檔案路徑中的所有資料夾名稱
                    string[] folders = file.fullPath.Split('\\');

                    // 提取倒數第二個資料夾名稱作為父資料夾名稱
                    if (folders.Length >= 2)
                    {
                        string parentFolderName = folders[folders.Length - 2];

                        // 如果字典中不包含父資料夾名稱，則將其添加到字典中
                        if (!folderDictionary.ContainsKey(parentFolderName))
                        {
                            folderDictionary.Add(parentFolderName, file.parentID);
                        }
                    }
                }*/

                // 取得最後同步時間
                //DateTime lastSyncTime = await cloudFileManager.GetLastSyncTime();

                /* 找出地端有 雲端沒有 的檔案 and (地端建立日期>LastSyncTime 或 地端最後修改日期>LastSyncTime)  union  雲端地端都有 且 雲端最後修改日期<地端檔案最後修改日期 */

                // [比對]找出地端有 雲端沒有 的檔案 
                List<Files> extraLocalFiles = localFiles.Except(cloudFiles, new FilesComparer()).ToList();

                // [比對](地端建立日期>LastSyncTime 或 地端最後修改日期>LastSyncTime)
                List<Files> filteredFiles = extraLocalFiles.Where(file => file.lastModiDate > lastSyncTime || file.createTime > lastSyncTime).ToList();

                // [比對]雲端地端都有的檔案(包含nodeID、fileID) 且 雲端最後修改日期<地端檔案最後修改日期(地端較新)
                List<Files> commonFiles = comparer.GetCommonFilePaths(localFiles, cloudFiles, "less");

                /*List<Files> commonFiles = new List<Files>();

                // 遍歷cloudFiles列表
                foreach (var cloudFile in cloudFiles)
                {
                    // 在localFiles列表中尋找與cloudFile.fullPath相同的文件
                    var localFile = localFiles.FirstOrDefault(f => f.fullPath == cloudFile.fullPath);

                    // 確保找到了對應的localFile並檢查lastModiDate條件
                    if (localFile != null && cloudFile.lastModiDate < localFile.lastModiDate)
                    {
                        // 將符合條件的文件加入結果列表
                        commonFiles.Add(cloudFile);
                    }
                }*/

                // 合併兩種情況的檔案列表FileList
                List<Files> combinedList = filteredFiles.Concat(commonFiles).ToList();

                // 檔案異動數量
                int syncFileCount = 0;

                // 檢查 combinedList 中的元素數量是否大於 0
                if (combinedList.Count > 0)
                {
                    /* 上傳到雲端 資料夾parentID? 設定雲端檔案的最後修改日期與地端檔案的最後修改日期一致 (目前 Upload API 邏輯) */
                    foreach (Files file in combinedList)
                    {
                        // 上傳檔案到最後的資料夾
                        APIResponse uploadRespons = await cloudFileManager.UploadCloudFile(nodeId, true, file.fullPath, file.lastModiDate);
                        
                        if(uploadRespons.status == 200)
                        {
                            // 檔案數量
                            syncFileCount++;

                            // 顯示進度訊息
                            ClsCommon.showHint($"InnoDrive:上傳到雲端 {syncFileCount}/{combinedList.Count}");
                            CustomMessageBox.ShowCustomMessage($"InnoDrive:上傳到雲端 {syncFileCount}/{combinedList.Count}", 3000);
                        }
                    }
                }
                else
                {
                    // 中斷執行
                    return;
                }

                if (syncFileCount > 0)
                {
                    // 紀錄檔案異動動作、總數量
                    FileChangeSummary(syncStartTime,SyncAction.CloudAdd, syncFileCount, combinedList);
                }
            }

            /// <summary>
            /// 【地端→雲端】 刪除雲端檔案 [雲端刪除檔案] 檢查上層資料夾有無檔案，若無則直接刪除?或保留?
            /// </summary>
            public async Task DeleteCloudFiles(DateTime lastSyncTime, DateTime syncStartTime)
            {
                // 取得地端檔案列表
                List<Files> localFiles = localFileManager.GetLocalFiles();

                // 取得雲端檔案列表 and not (AccessLog.Action in ('Upload', 'UploadOverwrite', 'Recovery', 'Edit') and AccessLog.logtime>LastSyncTime)
                List<Files> cloudFiles = await cloudFileManager.GetMyDriveCloudItems("Full", syncStartTime);

                // 過濾掉"isNewFile":true的檔案列表 取得cloudOldFiles
                List<Files> cloudOldFiles = new List<Files>();
                foreach (Files file in cloudFiles)
                {
                    if (file.isNewFile == false)
                    {
                        cloudOldFiles.Add(file);
                    }
                }

                Console.WriteLine(cloudOldFiles);
                // 取得最後同步時間
                //DateTime lastSyncTime = await cloudFileManager.GetLastSyncTime();

                /* 找出雲端有 地端沒有的檔案 and not (AccessLog.Action in ('Upload', 'UploadOverwrite', 'Recovery', 'Edit') and AccessLog.logtime>LastSyncTime)*/

                // [比對]找出雲端有 而地端沒有 的檔案 (地端後來刪除的檔案 OldFiles)
                List<Files> extraCloudFiles = cloudOldFiles.Except(localFiles, new FilesComparer()).ToList();

                // 檔案異動數量
                int syncFileCount = 0;

                if (extraCloudFiles.Count > 0)
                {
                    foreach (Files file in extraCloudFiles)
                    {
                        /* 刪除雲端檔案 空資料夾順便刪除?*/
                        APIResponse DeleteFileRes = await cloudFileManager.DeleteCloudFile(file.fileID, file.parentID);
                        if(DeleteFileRes.status == 200)
                        {
                            // 檔案數量
                            syncFileCount++;
                            // 顯示進度訊息
                            ClsCommon.showHint($"InnoDrive:雲端刪除檔案 {syncFileCount}/{extraCloudFiles.Count}");
                            CustomMessageBox.ShowCustomMessage($"InnoDrive:雲端刪除檔案 {syncFileCount}/{extraCloudFiles.Count}", 3000);
                        }
                    }
                }
                else
                {
                    // 中斷執行
                    return;
                }

                if(syncFileCount > 0)
                {
                    // 紀錄檔案異動動作、總數量
                    FileChangeSummary(syncStartTime,SyncAction.CloudDelete, syncFileCount, extraCloudFiles);
                }
                
            }

            /// <summary>
            /// 更新地端檔案名稱 [地端 Rename 檔案]
            /// </summary>
            /// <param name="syncStartTime"></param>
            public async Task RenameLocalFiles(DateTime syncStartTime)
            {
                

                // 取得雲端"Rename"檔案列表
                List<Files> cloudRenameFiles = await cloudFileManager.GetMyDriveCloudItems("Rename", syncStartTime);

                // 檔案異動數量
                int syncFileCount = 0;

                foreach (Files file in cloudRenameFiles)
                {
                    string directory = System.IO.Path.GetDirectoryName(file.fullPath);
                    string newFullPath = System.IO.Path.Combine(directory, file.newFileName);

                    //localFileManager.RenameFile(file.fullPath, newFullPath);
                    localFileManager.RenameLocalItem(EnumlocalItems.File, file.fullPath, newFullPath);
                    
                    // 檔案數量
                    syncFileCount++;

                    // 顯示進度訊息
                    ClsCommon.showHint($"InnoDrive:地端Rename {syncFileCount}/{cloudRenameFiles.Count}");
                    CustomMessageBox.ShowCustomMessage($"InnoDrive:地端Rename {syncFileCount}/{cloudRenameFiles.Count}", 3000);
                }

                if(syncFileCount > 0)
                {
                    // 紀錄檔案異動動作、總數量
                    FileChangeSummary(syncStartTime,SyncAction.LocalRename, syncFileCount, cloudRenameFiles);
                }
            }

            /// <summary>
            /// 更新地端資料夾名稱 [地端 Rename 檔案]
            /// </summary>
            /// <param name="syncStartTime"></param>
            public async Task RenameLocalFolders(DateTime syncStartTime)
            {
                List<Files> cloudRenameFolders = await cloudFileManager.GetMyDriveCloudItems("RenameFolder", syncStartTime);

                // 檔案異動數量
                int syncFileCount = 0;

                foreach (Files file in cloudRenameFolders)
                {
                    string newFullPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(file.fullPath), file.newFileName);

                    localFileManager.RenameLocalItem(EnumlocalItems.Directory, file.fullPath, newFullPath);

                    // 檔案數量
                    syncFileCount++;

                    // 顯示進度訊息
                    ClsCommon.showHint($"InnoDrive:地端Rename {syncFileCount}/{cloudRenameFolders.Count}");
                    CustomMessageBox.ShowCustomMessage($"InnoDrive:地端Rename {syncFileCount}/{cloudRenameFolders.Count}", 3000);
                }

                if (syncFileCount > 0)
                {
                    // 紀錄檔案異動動作、總數量
                    FileChangeSummary(syncStartTime,SyncAction.LocalRename, syncFileCount, cloudRenameFolders);
                }
            }


        }
    }
}
