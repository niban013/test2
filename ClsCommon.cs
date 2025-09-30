using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Claims;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using DriveAgent.Job;
using DriveAgent.Tool;
using Fleck;
using Microsoft.Win32;
using Microsoft.VisualBasic.ApplicationServices;
using MQTTnet;
using MQTTnet.Client;
using Newtonsoft.Json;
using NLog;
using WebDav;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace DriveAgent
{
    public enum ExploreShortCut
    {
        WebDav,
        MyDrive
    }
    public enum SendCommand
    {
        WebApiWatch,
        WebApiPut,
        UnLock, //webdav unlock
        Lock,
        Mqtt,
        UnCheckOut,// 一般checkout
        InnoAgent,
        OpenFile
    }

    public enum NotifyIconType
    {
        OnLine,
        Offline,
        Fail,
        OnSync,
        OffSync
    }

    public class WebdavMessage
    {
        public SendCommand command { get; set; }
        public string id { get; set; }
        public string? cmdparams { get; set; }
        public string msg { get; set; }

    }

    /// <summary>
    /// 提供了一組方法,用於讀取、寫入、刪除 INI 檔案中的設定資訊
    /// </summary>
    class IniFile
    {
        string Path; // 用於儲存 INI 檔案的完整路徑
        string EXE = Assembly.GetExecutingAssembly().GetName().Name; // 儲存了當前執行程式的名稱

        // 從 kernel32.dll 匯入 WritePrivateProfileString 函式
        // 用於寫入 INI 檔案中的設定值
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern long WritePrivateProfileString(string Section, string Key, string Value, string FilePath);

        // 從 kernel32.dll 匯入 GetPrivateProfileString 函式
        // 用於讀取 INI 檔案中的設定值
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder RetVal, int Size, string FilePath);

        // 建構函式,接受一個可選的 INI 檔案路徑
        // 如果未指定,則使用與當前程式相同的名稱
        public IniFile(string IniPath = null)
        {
            Path = new FileInfo(IniPath ?? EXE + ".ini").FullName;
        }

        // 讀取 INI 檔案中指定節點和索引鍵的設定值
        public string Read(string Key, string Section = null)
        {
            var RetVal = new StringBuilder(255);
            // 讀取 INI 檔案中的設定值
            GetPrivateProfileString(Section ?? EXE, Key, "", RetVal, 255, Path);
            return RetVal.ToString();
        }

        // 寫入 INI 檔案中指定節點和索引鍵的設定值
        public void Write(string Key, string Value, string Section = null)
        {
            // 寫入 INI 檔案中的設定值
            WritePrivateProfileString(Section ?? EXE, Key, Value, Path);
        }

        // 刪除 INI 檔案中指定節點和索引鍵的設定值
        public void DeleteKey(string Key, string Section = null)
        {
            Write(Key, null, Section ?? EXE);
        }

        // 刪除 INI 檔案中指定節點的所有設定值
        public void DeleteSection(string Section = null)
        {
            Write(null, null, Section ?? EXE);
        }

        // 檢查 INI 檔案中是否存在指定節點和索引鍵的設定值
        public bool KeyExists(string Key, string Section = null)
        {
            return Read(Key, Section).Length > 0;
        }
    }

    class ProcessInfo
    {
        public string ProcessName { get; set; }
        public int PID { get; set; }
        public string CommandLine { get; set; }
        public string UserName { get; set; }
        public string UserDomain { get; set; }
        public string User
        {
            get
            {
                if (string.IsNullOrEmpty(UserName))
                {
                    return "";
                }
                if (string.IsNullOrEmpty(UserDomain))
                {
                    return UserName;
                }
                return string.Format("{0}\\{1}", UserDomain, UserName);
            }
        }
    }

    public static class ClsCommon
    {
        public static CancellationTokenSource ctsMqtt;
        public static CancellationToken ctsMqttToken;

        public static CancellationTokenSource ctsLogin;
        public static CancellationToken ctsLoginToken;

        public static ManagementEventWatcher startWatch;
        public static ManagementEventWatcher stopWatch;

        static List<IWebSocketConnection> ClientSockets = new List<IWebSocketConnection>();
        static List<ProcessInfo> lsProcessInfo = new List<ProcessInfo>();

        public static string product = InxSetting.ClsCommon.Product;
        public static IMqttClient mqttClient;
        // 提示框
        public static NotifyIcon notifyIcon;

        private static int currentRetry;
        private static int maxRetry = InxSetting.ClsCommon.MaxRetry;
        public static int mqttPollingSecs = InxSetting.ClsCommon.MqttPollingSecs;
        private static int mqttWaitSecs = InxSetting.ClsCommon.MqttWaitSecs;
        private static string webSocket = InxSetting.ClsCommon.WebSocket;

        // 取得LocalMapping排程時間(每多久執行一次job)
        public static int LocalMappingSecs = InxSetting.ClsCommon.LocalMappingSecs;

        private static List<string> webDavServer = new List<string>();

        public static Logger logger;
        public static bool reConnection = false;
        private static object ojbProcess = new object();

        private static string webdavMqtt = "";
        public static string autoupdateUrl = "";
        public static string currentWebdavServer = "";

        public static string Version
        {
            get
            {
                return Assembly.GetEntryAssembly().GetName().Version.ToString();
            }
        }

        public async static Task CreateScheduler()
        {
            // 建立Tool/QuartzHelper
            QuartzHelper.CreateScheduler();
            await QuartzHelper.AddJob<LocalMapping>(LocalMappingSecs);
        }

        public static void Init()
        {
            currentRetry = 0;
            mqttClient = null;
            notifyIcon = null;
        }

        public static async Task<bool> isAuth()
        {
            if (currentWebdavServer == "")
            {
                return false;
            }
            else
            {
                var uri = currentWebdavServer.TrimEnd('/');
                uri = uri.Replace("http://", "").Replace(":", "@");
                uri = @"\\" + uri + @"\DavWWWRoot";
                return Directory.Exists(uri);
            }
        }

        [Obsolete("windows 驗証用")]
        public static void Login()
        {
            // 組合出應用程式資料目錄下的 Microsoft\Windows\Network Shortcuts 路徑
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Windows", "Network Shortcuts", ClsCommon.product);
            // 在 path 路徑下找到第一個 .lnk 格式的檔案
            var fileName = Directory.GetFiles(path, "*.lnk").FirstOrDefault();

            // 建立一個新的 ProcessStartInfo 物件，指向找到的 .lnk 檔案路徑
            ProcessStartInfo pInfo = new ProcessStartInfo(Path.Combine(path, fileName));

            // 設定 UseShellExecute 為 true，使用 Windows Shell 執行程式
            pInfo.UseShellExecute = true;

            // 設定視窗樣式為隱藏
            pInfo.WindowStyle = ProcessWindowStyle.Hidden;

            // 使用 using 區塊來確保及時釋放 Process 資源
            using (Process p = new Process())
            {
                p.StartInfo = pInfo; // 將 ProcessStartInfo 設定給 Process
                p.Start(); // 啟動 Process
            }
        }

        /// <summary>
        /// webapi與WebDav是否連通
        /// </summary>
        /// <param name="bCreateShortCut"></param>
        /// <returns></returns>
        public static async Task<(bool bWebApi, bool bWebDav)> isConnection(bool bCreateShortCut = true)
        {
            //notifyIcon = null;
            //notifyIcon = new NotifyIcon();
            try
            {
                // 建LocalMapping快捷處
                registerQuick(ExploreShortCut.MyDrive);
                // 檢查webAPI連線
                var conn1 = ClsCommon.PingHost(InxSetting.ClsCommon.WebApiUrl + @"/" + "health"); 
                if (!conn1)
                {
                    // 如果 WebAPI 連線失敗, 返回 (conn1, false)
                    return (conn1, false);
                }

                // 嘗試重新創建 WebDAV 伺服器
                var diffUrl = await ClsCommon.ReCreateWebdavServer();
                if (diffUrl)
                {
                    // 如果成功重新創建 WebDAV 伺服器, 檢查是否需要建立快捷方式
                    if (bCreateShortCut)
                    {
                        // 如果是IsSuperAdmin, 建立 WebDAV 快捷方式
                        if (InxSetting.ClsCommon.IsSuperAdmin)
                        { 
                            registerQuick(ExploreShortCut.WebDav);
                            // ClsCommon.createQuick(); // 建WebDav快捷
                        }
                        else
                        {
                            registerQuick(ExploreShortCut.WebDav, true);
                            // removeShortcut(product); // 移除WebDav快捷
                        }
                    }
                }

                // 檢查 WebDAV 連線
                var conn2 = false;
                if (ClsCommon.currentWebdavServer == "")
                {
                    conn2 = false;
                }
                else
                {
                    conn2 = ClsCommon.PingHost(ClsCommon.currentWebdavServer + @"/" + "health");
                }
                // 返回 WebAPI 和 WebDAV 的連線狀態
                return (conn1, conn2);

            }
            catch (Exception ex)
            {
                // 如果發生任何異常, 返回 (false, false)
                return (false, false);
            }
        }

        /// <summary>
        /// 停止 MQTT 服務
        /// </summary>
        /// <returns></returns>
        public static async Task<bool> StopMqttService()
        {
            // 建立 MqttClientDisconnectOptions 物件,用於設定斷開連接的選項
            MqttClientDisconnectOptions opt = new MqttClientDisconnectOptions();

            // 設定斷開連接的原因為 NormalDisconnection
            opt.Reason = MqttClientDisconnectOptionsReason.NormalDisconnection;

            // 檢查 ClsCommon.mqttClient 是否為 null
            if (ClsCommon.mqttClient == null)
            {
                // 不做任何操作
            }
            else if (ClsCommon.mqttClient.IsConnected) // 如果 mqttClient 不為 null 且已連接
            {
                // 使用 DisconnectAsync 方法, 以設定好的選項斷開與 MQTT 服務器的連接
                await ClsCommon.mqttClient.DisconnectAsync(opt, ClsCommon.ctsMqttToken);

                // 釋放 mqttClient 物件
                ClsCommon.mqttClient.Dispose();
            }

            // 返回 true, 表示 MQTT 服務已成功停止
            return true;
        }

        /// <summary>
        /// 根據輸入的 bWebApi 參數來初始化相關參數
        /// </summary>
        /// <param name="bWebApi"></param>
        /// <returns></returns>
        public static async Task InitParaForWebApi(bool bWebApi)
        {
            // 如果傳入的 bWebApi 為 false
            if (!bWebApi)
            {
                // 檢查 QuartzHelper.scheduler 是否不為 null
                if (QuartzHelper.scheduler != null)
                {
                    // 如果 scheduler 不為 null, 則調用 Shutdown 方法關閉它
                    QuartzHelper.Shutdown(false);
                }
            }
            else // 如果 bWebApi 為 true
            {
                // 從 InxSetting.ClsCommon 中獲取 WebdavMqtt 和 AutoupdateUrl 屬性的值
                webdavMqtt = InxSetting.ClsCommon.WebdavMqtt;
                autoupdateUrl = InxSetting.ClsCommon.AutoupdateUrl;
            }

            // 如果 bWebApi 為 true, 且 QuartzHelper.scheduler 為 null
            if (bWebApi && QuartzHelper.scheduler == null)
            {
                // 調用 CreateScheduler 方法創建一個新的 scheduler (新增localmapping job)
                await ClsCommon.CreateScheduler();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static async Task<bool> ReCreateWebdavServer()
        {
            // 初始化 _diffUrl 為 false
            var _diffUrl = false;
            webDavServer = InxSetting.ClsCommon.WebdavServer;

            // 如果 webDavServer 列表不為空, 則取列表中的第一個元素
            var _currentWebdavServer = webDavServer.Count > 0 ? webDavServer[0] : "";

            // 如果 currentWebdavServer 為空字符串, 或者與 _currentWebdavServer 不同
            if (currentWebdavServer == "" || (currentWebdavServer != _currentWebdavServer))
            {
                // 更新 currentWebdavServer 為 _currentWebdavServer
                currentWebdavServer = _currentWebdavServer;

                // 設置 _diffUrl 為 true, 表示 Webdav 服務器已經改變
                _diffUrl = true;
            }
            else
            {
                // 如果 currentWebdavServer 與 _currentWebdavServer 相同, 則設置 _diffUrl 為 false
                _diffUrl = false;
            }

            // 返回 _diffUrl, 表示 Webdav 服務器是否已經改變
            return _diffUrl;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ticket"></param>
        /// <returns></returns>
        public static HttpClient GetHttpClient(string ticket)
        {
            HttpClient client = new HttpClient()
            {
                // 設置請求超時時間為20分鐘
                Timeout = TimeSpan.FromMinutes(20)
            };

            // 如果 ticket 不為空字串
            if (ticket != "")
            {
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Add("ticket", ticket);

            }

            // 返回創建好的 HttpClient 實例
            return client;
        }

        /// <summary>
        /// 建快捷
        /// </summary>
        //public static void createQuick()
        //{
        //    registerQuick(product, currentWebdavServer);
        //}

        static string getFileType(string ext)
        {
            var fileType = "";
            switch (ext)
            {
                case ".xls":
                case ".xlsx":
                    fileType = "EXCEL";
                    break;
                case ".doc":
                case ".docx":
                    fileType = "WORD";
                    break;
                case ".ppt":
                case ".pptx":
                    fileType = "POWERPOINT";
                    break;
                case ".txt":
                    fileType = "notepad";
                    break;

            }
            return fileType;
        }

        static string getMatchFile(string ext)
        {
            var fileType = "";
            switch (ext)
            {
                case ".xls":
                case ".xlsx":
                    fileType = "Excel";
                    break;
                case ".doc":
                case ".docx":
                    fileType = "Word";
                    break;
                case ".ppt":
                case ".pptx":
                    fileType = "Powerpoint";
                    break;
                case ".txt":
                    fileType = "Notepad";
                    break;

            }
            return fileType;
        }


        /// <summary>
        /// 註冊websocket伺服器 (app與agent) 處理來自客戶端app的各種命令和請求
        /// 當收到客戶端發送的訊息時,根據不同的命令類型,執行相應的處理操作,例如檢查agent是否啟動、執行命令、開啟文件、檢查文件是否被打開等
        /// 如果需要回應客戶端,則將結果發送給所有連線的客戶端
        /// </summary>
        public static void registerWebSocket()
        {
            // 建立一個WebSocket連線清單，用於存儲所有連線的客戶端
            ClientSockets = new List<IWebSocketConnection>();

            // 建立一個WebSocketServer實例，指定監聽的位址和埠號
            var server = new WebSocketServer(ClsCommon.webSocket);

            // 設置監聽socket的參數
            server.ListenerSocket.NoDelay = true; // 禁用Nagle演算法，減少延遲
            server.RestartAfterListenError = true; // 如果監聽出錯，重新啟動
            // 啟動WebSocketServer
            server.Start(
                socket =>
                {
                    /* 當有新的WebSocket連線建立時，執行以下操作:*/

                    // 1. 連線開啟時，將這個socket加入到連線清單中
                    socket.OnOpen = () =>
                    {
                        //app 連上加入清單
                        ClientSockets.Add(socket);
                    };

                    // 2. 連線關閉時，將這個socket從連線清單中移除
                    socket.OnClose = () =>
                    {
                        //app 離線移除清單
                        ClientSockets.Remove(socket);
                    };

                    // 3. 當收到客戶端發送的訊息時，進行處理
                    socket.OnMessage = async message =>
                    {
                        // 標記是否需要回應客戶端
                        var isSendtoClient = false;

                        // 客戶端發送的訊息
                        var webdavAction = JsonConvert.DeserializeObject<WebdavMessage>(message);

                        /* 根據收到的命令類型進行不同的處理 */
                        if (webdavAction.command == SendCommand.Mqtt) // app 詢問agent是否啟動
                        {
                            // 檢查指定進程是否運行
                            var process = Process.GetProcessesByName(webdavAction.id).FirstOrDefault();// webdavAction.id ="InnoDriveAgent"
                            if (process == null)
                            {
                                webdavAction.msg = "N"; // 進程未運行
                            }
                            else
                            {
                                // 檢查MQTT連線是否成功
                                webdavAction.msg = ClsCommon.mqttClient.IsConnected ? "Y" : "N"; // mqtt是否與webdav serve連線成功
                            }
                            isSendtoClient = true;  // 回給app通知
                        }
                        else if (webdavAction.command == SendCommand.InnoAgent) // 從app 接收cmd呼叫
                        {
                            ExecuteCommand(webdavAction.cmdparams); // 執行命令
                        }
                        else if (webdavAction.command == SendCommand.OpenFile) // 開啟txt命令
                        {
                            OpenFile(webdavAction.cmdparams); // 開啟文件
                        }
                        else if (webdavAction.command == SendCommand.UnLock) // 從app 接收UnLock，判斷檔案是否還開著，開著不能解lock
                        {
                            
                            // 檢查文件是否仍然打開
                            var isOpen = false;
                            // 獲取與檔案類型相關的所有進程
                            var processes = Process.GetProcessesByName(getFileType(Path.GetExtension(webdavAction.id)));
                            // 遍歷所有獲取的進程
                            for (int i = 0; i < processes.Length; i++)
                            {
                                // 窗口標題以 webdavAction.id 開頭 && 以特定格式結尾，
                                if (processes[i].MainWindowTitle.StartsWith(webdavAction.id) && (processes[i].MainWindowTitle.EndsWith(" - " + getMatchFile(Path.GetExtension(webdavAction.id)))))
                                {
                                    isOpen = true;
                                    break;
                                }
                            }
                            webdavAction.msg = isOpen ? "Y" : "N"; // 設置解鎖結果
                            isSendtoClient = true; // 需要回應客戶端

                            //var test = await AddWebdavLog("Agent registerWebSocket Unlock", " ", " ", " ", " ", $"webdavAction.id:{webdavAction.id},isOpen:{isOpen}");
                        }
                        else if (webdavAction.command == SendCommand.Lock)
                        {
                            // 處理加鎖命令
                        }

                        // 如果需要回應客戶端，則將結果發送給所有連線的客戶端
                        if (isSendtoClient) //通知app
                        {
                            ClientSockets.ToList().ForEach(s => s.Send(JsonConvert.SerializeObject(webdavAction)));
                        }
                    };
                }
           );
        }

        /// <summary>
        /// 從 "shell32.dll" 動態鏈接庫中匯入 "ShellExecute" 函數
        /// </summary>
        /// <param name="hwnd"> 指定父窗口的句柄。如果沒有父窗口,可以傳遞 IntPtr.Zero </param>
        /// <param name="lpOperation"> 要執行的操作,例如 "open"、"edit"、"explore" 等 </param>
        /// <param name="lpFile"> 要執行的文件路徑或程式名稱 </param>
        /// <param name="lpParameters"> 傳遞給 lpFile 的參數 </param>
        /// <param name="lpDirectory"> 指定工作目錄 </param>
        /// <param name="nShowCmd"> 指定窗口顯示方式,例如 SW_SHOWNORMAL、SW_SHOWMINIMIZED 等 </param>
        /// <returns></returns>
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern int ShellExecute(IntPtr hwnd, string lpOperation, string lpFile, string lpParameters, string lpDirectory, int nShowCmd);

        public async static void OpenFile(string command)
        {
            try
            {
                // 從傳入的 command 字串中移除 "InxOpenExe:" 和 "&open" 部分
                command = command.Replace("InxOpenExe:", "").Replace("&open", "");

                // 將剩餘的 command 字串作為 URI 建立
                Uri uri = new Uri(command);

                // 獲取所有 Notepad 進程
                var processes = Process.GetProcessesByName("notepad");

                // 檢查是否已經有一個同名的 Notepad 窗口打開
                var exist = processes.Where(p => p.MainWindowTitle.Equals(Path.GetFileName(uri.AbsolutePath) + " - Notepad")).FirstOrDefault();

                // 如果已經有一個同名的 Notepad 窗口打開,拋出已開啟提示異常
                if (exist != null)
                {
                    throw new Exception("open:" + "檔案" + Path.GetFileName(uri.AbsolutePath) + "已開啟!");
                }
                // 將 URI 轉換為 UNC 路徑格式 (傳遞給 lpFile 的參數)
                var f = HttpUtility.UrlDecode(@"\\" + uri.AbsoluteUri.Replace("http://", "").Replace("/", "\\"));
                // 使用 ShellExecute 開啟 Notepad 並打開文件
                ShellExecute(IntPtr.Zero, "open", "notepad.exe", f, null, 1);


                //ProcessStartInfo pInfo = new ProcessStartInfo(@"C:\Windows\System32\notepad.exe");
                //pInfo.Arguments = @"\\" + uri.AbsoluteUri.Replace("http://", "").Replace("/", "\\"); 
                //pInfo.UseShellExecute = false; //是否使用操作系統shell啓動 

                //pInfo.CreateNoWindow = true;//不顯示程序窗口

                //using (Process p = new Process())
                //{
                //    p.StartInfo = pInfo;
                //    p.Start();
                //} 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 事件處理程序,它會在外部程序輸出數據時被調用.
        /// </summary>
        /// <param name="sender"> 觸發事件的物件,在這裡通常是一個Process(外部程序)實例 </param>
        /// <param name="e"> 包含輸出數據的EventArgs對象 </param>
        static void proc_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            // 檢查是否有輸出數據,如果有,則進行以下處理
            if (e.Data != null)
            {
                // Parse the output to get the progress information
                // Update your progress bar with the progress information
            }
        }

        /// <summary>
        /// 執行一個外部命令並顯示該命令的輸出
        /// </summary>
        /// <param name="command2"> 要執行的命令 </param>
        static void ExecuteCommand(string command2)
        {
            // 創建一個空的Form,僅用於保持程序運行
            Form f = new Form();
            f.Show();

            // 構建完整的命令字符串
            string command = @$"/c drvcmd {command2}";

            // 配置ProcessStartInfo,設置要執行的命令和參數
            ProcessStartInfo startInfo = new ProcessStartInfo("CMD", command);
            startInfo.WorkingDirectory = @"D:\\mars\\1002\\InnoDriveCmd\\bin\\Debug\\net7.0\\win-x64"; // 設置工作目錄
            startInfo.RedirectStandardOutput = true; // 重定向標準輸出
            startInfo.UseShellExecute = false; // 設置為false,以便重定向標準輸出
            startInfo.CreateNoWindow = true; // 不顯示命令提示符窗口

            // Create the process and start it
            Process process = new Process();
            process.StartInfo = startInfo; // 設置為上面配置好的StartInfo
            process.EnableRaisingEvents = true; // 啟用事件處理
            process.OutputDataReceived += new DataReceivedEventHandler(proc_OutputDataReceived); // 訂閱標準輸出數據接收事件
            // 啟動Process並開始讀取標準輸出
            process.Start();
            process.BeginOutputReadLine();
        }

        /// <summary>
        /// 將訊息傳回client端
        /// </summary>
        /// <param name="cmd"> 要發送的命令類型 </param>
        /// <param name="id"> 消息的唯一標識符 </param>
        /// <param name="msg"> 要發送的消息內容 </param>
        static void sendtoClient(SendCommand cmd, string id, string msg)
        {
            // 創建一個新的WebdavMessage對象
            WebdavMessage m = new WebdavMessage();
            // 設置WebdavMessage對象的屬性
            m.command = cmd;
            m.id = id;
            m.msg = msg;

            // 遍歷所有客戶端套接字,並向每個套接字發送序列化後的WebdavMessage對象
            ClientSockets.ToList().ForEach(s => s.Send(JsonConvert.SerializeObject(m)));
        }

        /// <summary>
        /// 開啟彈出視窗
        /// </summary>
        /// <typeparam name="T"> 表示要顯示的表單類型,必須是 Form 的子類 </typeparam>
        /// <param name="args"> 表示創建表單實例時需要的參數 </param>
        public static void showPopForm<T>(params object[] args)
         where T : Form
        {
            Form form = null;
            // 獲取所有已打開的表單中屬於 T 類型的表單
            IEnumerable<T> forms = Application.OpenForms.OfType<T>();
            if (forms.Any())
                form = forms.First();
            if (form == null)
                form = (Form)Activator.CreateInstance(typeof(T), args);

            for (int i = 1; i < Application.OpenForms.Count; i++)
            {
                if (Application.OpenForms[i].Name != form.Name)
                {
                    Application.OpenForms[i].Close();
                }
            }

            if (form.WindowState == FormWindowState.Minimized)
                form.WindowState = FormWindowState.Normal;
            else
                form.Show();
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// 註冊MQTT client端連線，接收WebDav訊息命令
        /// 訊息提示元件notifyIcon1
        /// </summary>
        /// <param name="user"></param>
        /// <param name="notifyIcon1"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public static async Task registerMQTT(string user, NotifyIcon notifyIcon1, CancellationToken cancellationToken)
        {
            try
            {
                // 初始化MQTT相關設定
                Init();
                notifyIcon = notifyIcon1;
                ChangeNotifyIcon(NotifyIconType.OnLine, notifyIcon); // 將通知圖標設置為在線 NotifyIconType.Offline
                
                // 如果已經創建了MQTT客戶端,則直接返回
                if (mqttClient != null)
                    return;

                // 設置MQTT連接用戶名和密碼
                var UserName = "webdav";
                var Password = "webdav";

                // 創建MQTT客戶端
                mqttClient = new MqttFactory().CreateMqttClient();

                // 獲取本機的IP地址
                var clientIp = GetLocalIPAddress().ToString(); //取本機10.開頭ip
                if (clientIp == string.Empty)
                {
                    throw new Exception("本機無網路");
                }
                showMsg("Agent啟動...");

                // 創建MQTT client端連接配置訊息
                var options = new MqttClientOptionsBuilder()
                        .WithClientId(user + "@" + clientIp) // 設置客戶端ID (ConnectionADcminl)
                        .WithCredentials(UserName, Password) // 要訪問的mqtt server端的用戶名和密碼
                        .WithKeepAlivePeriod(TimeSpan.FromSeconds(7.5)) // 設置心跳間隔
                        .WithCleanSession() // // 設置是否清理會話
                        .WithProtocolVersion(MQTTnet.Formatter.MqttProtocolVersion.V500) // 設置MQTT協議版本
                        .WithUserProperty(user + "@" + clientIp, clientIp) // 設置用戶屬性
                        //.WithTcpServer("10.56.133.138", 61613)
                        .WithTcpServer(webdavMqtt.Split(':').First(), Convert.ToInt32(webdavMqtt.Split(':').Last())) // 要訪問的mqtt server端的ip和port
                        .Build();

                // When client connected to the server
                // MessageBox.Show("HostName:" + HostName);
                string topic = (user + "_" + clientIp).ToUpper();

                // 連接MQTT服務器時的事件處理程序(客戶端連接事件)
                mqttClient.ConnectedAsync += (async e =>
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        return;
                    }
                    //註冊topic通訊管道(訂閱話題)
                    await mqttClient.SubscribeAsync(new MqttTopicFilterBuilder().WithTopic(topic).Build());
                    sendtoClient(SendCommand.Mqtt, "DriveAgent", "Y"); //通知app agent起動

                    // 顯示連接成功的訊息
                    var msg = new WebdavMessage();
                    msg.msg = "連線服務器成功";
                    ChangeNotifyIcon(NotifyIconType.OnLine, notifyIcon);
                    showMessage(NotifyIconType.OnLine, toBase64(JsonConvert.SerializeObject(msg))); //通知agent訊息起動
                });

                // 接收MQTT服務器發送的訊息時的事件處理程序 (接收 webdav server 傳來訊息)
                mqttClient.ApplicationMessageReceivedAsync += (async e =>
                {
                    // 處理接收到的訊息
                    var payload = e.ApplicationMessage.ConvertPayloadToString();
                    //var test1 = await AddWebdavLog("Agent registerMQTT Unlock start", clientIp, user, " ", " ", $"payload{payload}");
                    var decodeCommand = JsonConvert.DeserializeObject<WebdavMessage>(payload);
                    if (decodeCommand.command == SendCommand.WebApiWatch) //解析一般訊息
                    {
                        if (decodeCommand.msg == "OK") //正常訊息
                        {
                            if (notifyIcon1.BalloonTipIcon == ToolTipIcon.Info)
                            {
                            }
                            else
                            {
                                showMessage(NotifyIconType.OnLine, toBase64(payload));
                            }
                        }
                        else//異常訊息
                        {
                            showMessage(NotifyIconType.Offline, toBase64(payload));
                        }
                    }
                    else if (decodeCommand.command == SendCommand.WebApiPut)//檔案新增或儲存、存檔訊息
                    {
                        showMessage(NotifyIconType.OnLine, toBase64(payload));
                    }
                    else if (decodeCommand.command == SendCommand.UnLock) //自動解lock
                    {
                        //var test2 = await AddWebdavLog("Agent registerMQTT Unlock", clientIp, user, " ", " ", $"payload{payload}");
                        ClientSockets.ToList().ForEach(s => s.Send(payload));  //send to app
                    }
                    else if (decodeCommand.command == SendCommand.Lock) //檔案總管操作無鎖定(暫沒使用)
                    {
                        var result = _client.Lock(currentWebdavServer + decodeCommand.id.Replace("\\", "/").Replace("\"", "")).Result;
                        if (!result.IsSuccessful)
                        {
                            throw new Exception($"檔案遺失，開啟失敗");
                        }
                    }
                });

                // 客戶端連接關閉( 斷開MQTT連接 )事件
                mqttClient.DisconnectedAsync += (async e =>
                {
                    // 通知app agent已停止
                    sendtoClient(SendCommand.Mqtt, "DriveAgent", "N");
                    if (cancellationToken.IsCancellationRequested)
                    {
                        return;
                    }

                    // 將通知圖標設置為離線
                    ChangeNotifyIcon(NotifyIconType.Offline, notifyIcon);

                    // 如果重試次數超過最大重試次數,則顯示錯誤訊息並返回
                    if (currentRetry >= ClsCommon.maxRetry)  //最大重試
                    {
                        var msg = new WebdavMessage();
                        msg.msg = "[Error Code:503]MQTT初始化失敗,請嘗試手動重啟";
                        showMessage(NotifyIconType.Fail, toBase64(JsonConvert.SerializeObject(msg)));
                        return;
                    }

                    // 顯示離線訊息,並在一段時間後嘗試重新連接
                    showMessage(NotifyIconType.Offline);
                    Console.WriteLine("MQTT reconnecting");
                    await Task.Delay(TimeSpan.FromSeconds(ClsCommon.mqttWaitSecs));
                    await mqttClient.ConnectAsync(options, cancellationToken);
                });

                // 連接MQTT服務器
                await mqttClient.ConnectAsync(options, cancellationToken);//註冊連線
            }
            catch (Exception ex)
            {
                if (ex.Message.ToString() == "本機無網路")
                {
                    var msg = new WebdavMessage();
                    msg.msg = "本機無網路";
                    showMessage(NotifyIconType.Offline, toBase64(JsonConvert.SerializeObject(msg)));
                }
            }

        }

        /// <summary>
        /// 依照 NotifyIconType 切換 NotifyIcon(Agent Icon)
        /// </summary>
        /// <param name="notiType"></param>
        /// <param name="notifyIcon"></param>
        public static void ChangeNotifyIcon(NotifyIconType notiType, NotifyIcon notifyIcon)
        {
            try
            {
                notifyIcon.Icon = null;
                //notifyIcon.Icon = (notiType == NotifyIconType.OnLine) ? ((notiType == NotifyIconType.OnSync) ? DriveAgent.Resource1.inxdrive32flash : DriveAgent.Resource1.inxdrive32 ) : DriveAgent.Resource1.inxdrive32g;
                if (notiType == NotifyIconType.OnLine)
                {
                    if (notiType == NotifyIconType.OnSync)
                    {
                        notifyIcon.Icon = DriveAgent.Resource1.inxdrive32flash;
                    }
                    else if(notiType == NotifyIconType.OffSync)
                    {
                        notifyIcon.Icon = DriveAgent.Resource1.inxdrive32;
                    }
                    else
                    {
                        notifyIcon.Icon = DriveAgent.Resource1.inxdrive32;
                    }
                }
                else
                {
                    notifyIcon.Icon = DriveAgent.Resource1.inxdrive32g;
                }

                notifyIcon.BalloonTipIcon = ToolTipIcon.Info;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void showMessage(NotifyIconType notiType, string msg = "")
        {
            if (notifyIcon != null)
            {
                if (notiType == NotifyIconType.Fail)
                {
                    ChangeNotifyIcon(NotifyIconType.Fail, notifyIcon);
                    var msgInfo = JsonConvert.DeserializeObject<WebdavMessage>(fromBase64(msg));
                    mqttClient = null;
                    notifyIcon.Text = "無法連線......";
                    notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
                    notifyIcon.BalloonTipTitle = "[Status Code:503] Client=>WebDav";
                    notifyIcon.BalloonTipText = msgInfo.msg;
                    notifyIcon.ShowBalloonTip(1000);
                    logger.Error($@"WebDav:{notifyIcon.Text}");
                    return;
                }
                if (msg == string.Empty)
                {
                    if (notiType == NotifyIconType.OnLine)
                    {
                        ChangeNotifyIcon(NotifyIconType.OnLine, notifyIcon);
                        notifyIcon.Text = "已連線......";
                        notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
                        notifyIcon.BalloonTipTitle = "[Status Code:200] Client=>WebDav";
                        notifyIcon.BalloonTipText = "";
                        currentRetry = 0;
                        logger.Info($@"WebDav:{notifyIcon.Text}");
                    }
                    else
                    {
                        ChangeNotifyIcon(NotifyIconType.Fail, notifyIcon);
                        currentRetry++;
                        notifyIcon.Text = "重試連線......";
                        notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
                        notifyIcon.BalloonTipTitle = $@"[Status Code:400]服務器斷線(Client=>WebDav)";
                        notifyIcon.BalloonTipText = currentRetry > ClsCommon.maxRetry ? "請連絡IT" : $@"正在重試中{currentRetry}/{ClsCommon.maxRetry}......";
                        notifyIcon.ShowBalloonTip(1000);
                        logger.Error($@"WebDav:{notifyIcon.Text},{notifyIcon.BalloonTipTitle}");
                    }
                }
                else
                {
                    var msgInfo = JsonConvert.DeserializeObject<WebdavMessage>(fromBase64(msg));
                    notifyIcon.Text = "已連線......";
                    notifyIcon.BalloonTipText = (msgInfo.msg.Contains("ConnectionRefused") || msgInfo.msg.Contains("結果:無法連線，因為目標電腦拒絕連線")
                        || msgInfo.msg.Contains("結果:連線嘗試失敗")) ? "InnoDriveApi連線失敗" : msgInfo.msg.Replace(",", "\r\n");
                    //notifyIcon.BalloonTipText = notifyIcon.BalloonTipText == "OK" ? "InnoDriveApi已連線":"";
                    if (notifyIcon.BalloonTipText.Contains("InnoDriveApi連線失敗"))
                    {
                        notifyIcon.Text = "重試連線......";
                        ChangeNotifyIcon(NotifyIconType.Fail, notifyIcon);
                        notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
                        notifyIcon.BalloonTipTitle = "[Status Code:503] WebDav=>InxDriveApi";
                    }
                    else if (notifyIcon.BalloonTipText.Contains("本機無網路"))
                    {
                        notifyIcon.Text = "重試連線......";
                        ChangeNotifyIcon(NotifyIconType.Fail, notifyIcon);
                        notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
                        notifyIcon.BalloonTipTitle = "[Status Code:503]  Client=>WebDav";
                    }
                    else
                    {
                        ChangeNotifyIcon(NotifyIconType.OnLine, notifyIcon);
                        notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
                        notifyIcon.BalloonTipTitle = "[Status Code:200] Client=>WebDav";
                        currentRetry = 0;
                    }
                    notifyIcon.ShowBalloonTip(1000);
                    logger.Info($@"WebDav:{notifyIcon.Text},InnoDriveApi:{notifyIcon.BalloonTipText}");
                }
            }
        }

        public static void showMsg(string message)
        {
            //notifyIcon.Visible = true; // 將 NotifyIcon 設置為可見
            if (notifyIcon != null)
            {
                notifyIcon.Visible = true; // 將 NotifyIcon 設置為可見
                notifyIcon.Text = message;
                notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
                notifyIcon.BalloonTipTitle = "訊息通知";
                notifyIcon.BalloonTipText = message;
                notifyIcon.ShowBalloonTip(2000);
            }
            else
            {
                // 處理 notifyIcon 為 null 的情況
                throw new Exception("notifyIcon is null");
            }
            
        }

        public static void showHint(string message)
        {
            if (notifyIcon != null)
            {
                notifyIcon.Visible = true; // 將 NotifyIcon 設置為可見
                notifyIcon.Text = message;
                notifyIcon.BalloonTipText = " ";
                //notifyIcon.ShowBalloonTip(1000);
            }
            else
            {
                // 處理 notifyIcon 為 null 的情況
                throw new Exception("notifyIcon is null");
            }
        }

        /// <summary>
        /// 將輸入字符串轉換為Base64編碼的字符串
        /// </summary>
        /// <param name="inputText"></param>
        /// <returns></returns>
        public static string toBase64(string inputText)
        {
            // 將輸入字符串轉換為UTF-8編碼的字節數組
            Byte[] bytesEncode = System.Text.Encoding.UTF8.GetBytes(inputText);
            // 將字節數組轉換為Base64編碼的字符串並返回
            return Convert.ToBase64String(bytesEncode);
        }

        /// <summary>
        /// 將Base64編碼的字符串轉換為原始字符串
        /// </summary>
        /// <param name="resultEncode"></param>
        /// <returns></returns>
        private static string fromBase64(string resultEncode)
        {
            // 將Base64編碼的字符串轉換為字節數組
            Byte[] bytesDecode = Convert.FromBase64String(resultEncode);
            // 將字節數組轉換為UTF-8編碼的字符串並返回
            return System.Text.Encoding.UTF8.GetString(bytesDecode);
        }

        /// <summary>
        /// 建立InnoDriveAgent快捷
        /// </summary>
        /// <param name="product"> 產品名稱,用於命名快捷方式 </param>
        /// <param name="server"> 伺服器的 URL 地址 </param>
        //public static void registerQuick(string product, string server)
        //{
        //    // 刪除伺服器 URL 末尾的斜杠
        //    server = server.TrimEnd('/');
        //    server = server.Replace("http://", "").Replace("/", "\\").Replace(":", "@");
        //    // 將轉換後的伺服器 URL 拼接為 \server\DavWWWRoot 路徑
        //    server = @"\\" + server + @"\DavWWWRoot";

        //    // 移除之前的快捷方式
        //    removeShortcut(product);
        //    // 在桌面上創建新的快捷方式
        //    createDesktopini(product, server, 0);
        //}

        /// <summary>
        /// 新增檔案總管MyDrive圖示捷徑
        /// </summary>
        public static void registerQuick(ExploreShortCut shortcut,bool bCancel = false)
        {
            try
            {
                string currentPath = "";
                Process regeditProcess;
                var strRegPath = "";
                string folder = "";
                string regName = "";
                string guid = "";
                string ico = "";
                RegistryKey? key;
                #if (DEBUG)
                     currentPath =Directory.GetCurrentDirectory(); 
                #endif
                #if (!DEBUG)
                     currentPath = Directory.GetParent(Directory.GetCurrentDirectory()).FullName;
                #endif


                switch (shortcut)
                {
                    case ExploreShortCut.WebDav:
                        regName = "DriveAgent";
                        guid = "9435a7b0-c238-400d-b149-b9629886507a";
                        //刪除伺服器 URL 末尾的斜杠
                        var server = currentWebdavServer.TrimEnd('/');
                        server = server.Replace("http://", "").Replace("/", "\\").Replace(":", "@");
                        // 將轉換後的伺服器 URL 拼接為 \server\DavWWWRoot 路徑
                        server = @"\\" + server + @"\DavWWWRoot";
                        folder = server.Replace("\\", "\\\\");
                        #if (DEBUG)           
                               ico = Path.Combine(Directory.GetParent(currentPath).Parent.Parent.Parent.FullName, "InxSetting", "Icon", "inxdrive32.ico").Replace("\\", "\\\\");
                        #endif
                        #if (!DEBUG)
                               ico = Path.Combine(Directory.GetParent(currentPath).FullName, InxSetting.ClsCommon.Product,"Icon", "inxdrive32.ico").Replace("\\", "\\\\");
                        #endif


                        if (bCancel)
                        {
                            string regText = "";
                            strRegPath = Path.Combine(currentPath, "Reg", "un"+ regName + ".reg"); 
                            regeditProcess = Process.Start("regedit.exe", "/s \"" + strRegPath + "\"");
                            regeditProcess.WaitForExit();
                        } 
                        break;
                    case ExploreShortCut.MyDrive: 
                        regName = "LocalMapping";
                        guid = "c0ce2d28-2699-4519-8622-b2c9d6c3c2b0";
                        folder = InxSetting.ClsCommon.GetIni("Config.ini", "Folder", "MyDriver").Replace("\\", "\\\\");
                        #if (DEBUG)
                           ico = Path.Combine(Directory.GetParent(currentPath).Parent.Parent.Parent.FullName, "InxSetting", "Icon", "inxdrive32.ico").Replace("\\", "\\\\");
                        #endif
                        #if (!DEBUG)
                           ico = Path.Combine(Directory.GetParent(currentPath).FullName, InxSetting.ClsCommon.Product,"Icon", "inxdrive32.ico").Replace("\\", "\\\\");
                        #endif
                     
                        if (folder == string.Empty)
                        {
                            return; 
                        } 
                        key = Registry.LocalMachine.OpenSubKey("Software\\Classes\\CLSID\\{"+ guid + "}");
                        if (key != null)
                        {
                            strRegPath = Path.Combine(currentPath, "Reg", "HKLMRemove.reg");
                            regeditProcess = Process.Start("regedit.exe", "/s \"" + strRegPath + "\"");
                            regeditProcess.WaitForExit();
                        } 
                        break; 
                }


                if (!bCancel)
                {
                    bool bReplace = false;
                    if (!File.Exists(Path.Combine(currentPath, "Reg", regName + ".reg")))
                    {
                        bReplace = true;
                    }
                    else
                    {
                        key = Registry.CurrentUser.OpenSubKey("Software\\Classes\\CLSID\\{" + guid + "}\\Instance\\InitPropertyBag");
                        if (key != null)
                        {
                            Object o = key.GetValue("TargetFolderPath");
                            if (o != null)
                            {
                                if ((o).ToString() != folder.Replace("\\\\", "\\"))
                                {
                                    bReplace = true;
                                }
                            }
                        }
                        else
                        {
                            bReplace = true;
                        }
                    }
                    #if (DEBUG)
                       bReplace = true;
                    #endif

                    if (bReplace)
                    {
                        var strUnRegPath = Path.Combine(currentPath, "Reg", "un" + regName + ".reg");
                        if (File.Exists(strUnRegPath))
                        {
                            regeditProcess = Process.Start("regedit.exe", "/s \"" + strUnRegPath + "\"");
                            regeditProcess.WaitForExit();
                        }

                        string regText = "";
                        strRegPath = Path.Combine(currentPath, "Reg", "Template.reg");
                        regText = File.ReadAllText(strRegPath);
                        regText = regText.Replace("#folder", folder);
                        regText = regText.Replace("#guid", guid);
                        regText = regText.Replace("#name", regName);
                        regText = regText.Replace("#ico", ico);
                        strRegPath = Path.Combine(currentPath, "Reg", regName + ".reg");
                        File.WriteAllText(strRegPath, regText);
                        regeditProcess = Process.Start("regedit.exe", "/s \"" + strRegPath + "\"");
                        regeditProcess.WaitForExit();

                        regText = regText.Replace("[", "[-");
                        File.WriteAllText(strUnRegPath, regText);
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 在指定的路徑上創建一個帶有 Desktop.ini 文件的資料夾,並在該資料夾中創建一個快捷方式
        /// </summary>
        /// <param name="name"></param>
        /// <param name="targetPath"> 快捷方式所指向的目標路徑 </param>
        /// <param name="iconNumber"></param>
        private static void createDesktopini(string name, string targetPath, int iconNumber)
        {
            // 獲取 "%APPDATA%\Microsoft\Windows\Network Shortcuts" 路徑
            string networkshortcuts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Windows", "Network Shortcuts");

            // 在 networkshortcuts 路徑下創建一個新的資料夾,名稱為 name 參數
            var newFolder = Directory.CreateDirectory(networkshortcuts + @"\" + name);

            // 將新創建的資料夾設置為唯讀屬性
            newFolder.Attributes |= FileAttributes.ReadOnly;

            // 構建 Desktop.ini 文件的內容
            string desktopiniContents = @"[.ShellClassInfo]" + Environment.NewLine +

                "CLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}" + Environment.NewLine +

                "Flags=2";

            // 獲取 Desktop.ini 文件的路徑
            string shortcutlocation = networkshortcuts + @"\" + name;

            // 寫入 Desktop.ini 文件的內容
            System.IO.File.WriteAllText(shortcutlocation + @"\Desktop.ini", desktopiniContents);

            // 獲取新創建資料夾的路徑
            string targetLNKPath = networkshortcuts + @"\" + name;

            // 在新創建的資料夾中創建一個快捷,目標路徑為 targetPath
            createShortcut(name, targetLNKPath, targetPath);

        }

        /// <summary>
        /// 創建快捷
        /// </summary>
        /// <param name="product"></param>
        /// <param name="shortcutPath"> 存放路徑 </param>
        /// <param name="targetPath"> 指向的目標路徑 </param>
        static void createShortcut(string product, string shortcutPath, string targetPath)
        {
            try
            {
                // 獲取快捷方式的完整路徑,包括檔案名 "target.lnk"
                string shortcutLocation = System.IO.Path.Combine(shortcutPath, "target.lnk");
                // 創建一個 WshShell 物件,用於操作 Windows 上的快捷方式
                var shell = new IWshRuntimeLibrary.WshShell();
                // 創建一個 WshShortcut 物件,表示一個快捷方式
                var shortcut = (IWshRuntimeLibrary.WshShortcut)shell.CreateShortcut(shortcutLocation);

                // 設置快捷方式的描述
                shortcut.Description = product;
                // 設置快捷方式的圖標位置
                shortcut.IconLocation = Path.Combine(@"C:\InnoLux\DriveAgent\Icon", "inxdrive32.ico");
                // 設置快捷方式的目標路徑
                shortcut.TargetPath = targetPath;
                // 保存快捷方式
                shortcut.Save();
            }
            catch (Exception ex)
            {

            }

        }

        /// <summary>
        /// 刪除一個快捷方式
        /// </summary>
        /// <param name="product"></param>
        static void removeShortcut(string product)
        {
            // 獲取快捷方式的存放路徑
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Windows", "Network Shortcuts");
            string shortcut_link = Path.Combine(path, product, "target.lnk");

            // 檢查快捷方式檔案是否存在
            if (System.IO.File.Exists(shortcut_link))
            {
                // 如果存在,則刪除快捷方式檔案
                System.IO.File.Delete(shortcut_link);

                // 獲取包含該快捷方式的資料夾
                var folder = new DirectoryInfo(Path.Combine(path, product));

                // 將資料夾屬性設置為正常
                folder.Attributes = FileAttributes.Normal;

                // 刪除整個資料夾
                Directory.Delete(folder.FullName, true);
            }
            else
            {
                // 如果快捷方式檔案不存在,則直接獲取包含該快捷方式的資料夾
                var folder = new DirectoryInfo(Path.Combine(path, product));
                // 如果資料夾存在,則將其屬性設置為正常,並刪除整個資料夾
                if (System.IO.Directory.Exists(folder.FullName))
                {
                    folder.Attributes = FileAttributes.Normal;
                    Directory.Delete(folder.FullName, true);
                }
            }
        }

        /// <summary>
        /// 檢查指定主機是否可以 ping 通
        /// 改用health偵測
        /// </summary>
        /// <param name="hostUri"></param>
        /// <returns></returns>
        public static bool PingHost(string hostUri)
        {
            bool pingable = false;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMilliseconds(10000); //置 HttpClient 的超時時間為 15 秒
                    client.DefaultRequestHeaders.ConnectionClose = true; //Set KeepAlive to false

                    // 發送一個 GET 請求到指定的 hostUri，等待異步操作完成
                    HttpResponseMessage response = client.GetAsync(hostUri).GetAwaiter().GetResult(); //Make sure it is synchronous
                    // 檢查 HTTP 響應的狀態碼是否為 200 OK
                    if (response.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        pingable = true;
                    }
                    else
                    {
                        pingable = false;
                    }
                }
            }
            catch (SocketException ex)
            {
                pingable = false;
            }
            return pingable;
        }

        [Obsolete("雲端會通，但pod不通，偵測不出來")]
        public static bool PingHost(string hostUri, int portNumber = 80)
        {
            bool pingable = false;
            try
            {
                using (var client = new TcpClient(hostUri, portNumber))
                {
                    pingable = true;
                }
            }
            catch (SocketException ex)
            {

            }
            return pingable;
        }

        /// <summary>
        /// 獲取系統上正在運行的 notepad.exe 進程信息,並監視其啟動和停止事件
        /// </summary>
        public static void NoWebDavSupportFile()
        {
            try
            {
                // 清空 lsProcessInfo 列表
                lsProcessInfo.Clear();

                // 創建一個 ProcessInfo 對象
                ProcessInfo p;

                // 設置查詢語句,獲取名為 'notepad.exe' 的進程
                string query = "Select * From Win32_Process WHERE Name = 'notepad.exe'";
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
                ManagementObjectCollection processList = searcher.Get();

                // 遍歷找到的 notepad.exe 進程
                using (var results = processList)
                {
                    foreach (ManagementObject result in results)
                    {
                        p = new ProcessInfo();
                        try
                        {
                            // 獲取進程 ID、進程名稱和命令行參數
                            p.PID = Convert.ToInt32(result["processid"].ToString());
                            p.ProcessName = result["name"].ToString();
                            p.CommandLine += result["CommandLine"].ToString() + " ";
                        }
                        catch { }
                        try
                        {
                            // 獲取進程所有者信息
                            string[] argList = new string[2];
                            int returnVal = Convert.ToInt32(result.InvokeMethod("GetOwner", argList));
                            if (returnVal == 0)
                            {
                                p.UserDomain = argList[1];
                                p.UserName = argList[0];
                            }
                        }
                        catch { }
                        if (!string.IsNullOrEmpty(p.CommandLine))
                        {
                            // 修剪命令行參數
                            p.CommandLine = p.CommandLine.Trim();
                        }
                        // 將進程信息添加到 lsProcessInfo 列表
                        lsProcessInfo.Add(p);
                    }

                }

                // 設置監視 notepad.exe 進程啟動事件的事件監視器
                startWatch = new ManagementEventWatcher(new WqlEventQuery("SELECT * FROM Win32_ProcessStartTrace WHERE ProcessName='notepad.exe'"));
                startWatch.EventArrived += new EventArrivedEventHandler(startWatch_EventArrived);
                startWatch.Start();

                // 設置監視 notepad.exe 進程停止事件的事件監視器
                stopWatch = new ManagementEventWatcher(new WqlEventQuery("SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName='notepad.exe'"));
                stopWatch.EventArrived += new EventArrivedEventHandler(stopWatch_EventArrived);
                stopWatch.Start();

            }
            catch (Exception ex)
            {

            }
        }

        // 建立一個 WebDAV 客戶端
        static IWebDavClient _client = new WebDavClient();

        /// <summary>
        /// 當 stopWatch 事件發生時觸發的方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <exception cref="Exception"></exception>
        static void stopWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            // 獲取事件中的處理程序資訊
            var proc = GetProcessInfo(e);
            // 在 lsProcessInfo 列表中尋找與該處理程序相匹配的項目
            var item = lsProcessInfo.Where(p => p.PID == proc.PID).FirstOrDefault();
            if (item != null)
            {
                // 鎖定 ojbProcess 物件以安全地從 lsProcessInfo 列表中移除項目
                lock (ojbProcess)
                {
                    lsProcessInfo.Remove(item);
                }

                // 從命令行中提取 URL
                MatchCollection matches = Regex.Matches(item.CommandLine, "\"([^\"]*)\"");
                var url = item.CommandLine.Replace(matches[0].Value, "").Trim();
                if (url.StartsWith("//"))
                {
                    // 如果 URL 以 "//" 開頭, 則發送"UnLock"命令給客戶端
                    sendtoClient(SendCommand.UnLock, Path.GetFileName(url), "UnLock");
                }
                else
                {
                    if (url.StartsWith("\""))
                    {
                        // 如果 URL 以 "\" 開頭, 則解析 URL 並調用 _client.Unlock() 方法解鎖文件
                        var lsUrl = new List<string>();
                        var tmp = url.Split('\\');
                        for (int i = 4; i < tmp.Length; i++)
                        {
                            lsUrl.Add(tmp[i]);
                        }
                        var result = _client.Unlock(currentWebdavServer + "/" + Path.Combine(lsUrl.ToArray()).Replace("\\", "/").Replace("\"", ""), "").Result;
                        if (!result.IsSuccessful)
                        {
                            throw new Exception($"解鎖失敗");
                        }
                    }

                }
            }
        }

        /// <summary>
        /// 當 startWatch 事件發生時觸發的方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void startWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            try
            {
                // 獲取事件中的處理程序資訊
                var proc = GetProcessInfo(e);
                // 在 lsProcessInfo 列表中尋找與該處理程序相匹配的項目
                var item = lsProcessInfo.Where(p => p.PID == proc.PID).FirstOrDefault();
                if (item == null)
                {
                    /*if (proc.CommandLine.StartsWith("\""))
                    //{ 
                    //    var lsUrl = new List<string>();
                    //    var tmp = proc.CommandLine.Split(' ')[2].Split('\\');
                    //    for (int i = 4; i < tmp.Length; i++)
                    //    {
                    //        lsUrl.Add(tmp[i]);
                    //    }
                    //    var result = _client.Lock("http://" + currentWebdavServer + "/" + Path.Combine(lsUrl.ToArray()).Replace("\\", "/").Replace("\"", "")).Result;
                    //    if (!result.IsSuccessful)
                    //    {
                    //        throw new Exception($"檔案遺失，開啟失敗");
                    //    }
                    //}*/

                    // 鎖定 ojbProcess 物件以安全地將 proc 添加到 lsProcessInfo 列表中
                    lock (ojbProcess)
                    {
                        lsProcessInfo.Add(proc);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// // 根據事件參數獲取處理程序資訊
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        static ProcessInfo GetProcessInfo(EventArrivedEventArgs e)
        {
            // 創建一個新的 ProcessInfo 對象
            var p = new ProcessInfo();
            // 嘗試從事件參數中獲取處理程序的ID
            var pid = 0;
            int.TryParse(e.NewEvent.Properties["ProcessID"].Value.ToString(), out pid);
            p.PID = pid;

            // 獲取處理程序的名稱
            p.ProcessName = e.NewEvent.Properties["ProcessName"].Value.ToString();
            try
            {
                // 使用 ManagementObjectSearcher 查詢 Win32_Process 類獲取更多處理程序資訊
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Process WHERE ProcessId = " + pid))
                using (var results = searcher.Get())
                {
                    foreach (ManagementObject result in results)
                    {
                        try
                        {
                            // 獲取處理程序的命令行參數
                            p.CommandLine += result["CommandLine"].ToString() + " ";
                        }
                        catch { }
                        try
                        {
                            // 獲取處理程序的運行用戶資訊
                            var user = result.InvokeMethod("GetOwner", null, null);
                            p.UserDomain = user["Domain"].ToString();
                            p.UserName = user["User"].ToString();
                        }
                        catch { }
                    }
                }

                // 修剪處理程序的命令行參數, 去掉多餘的空格
                if (!string.IsNullOrEmpty(p.CommandLine))
                {
                    p.CommandLine = p.CommandLine.Trim();
                }
            }
            catch (ManagementException) { }

            // 返回獲取到的處理程序資訊
            return p;
        }

        /// <summary>
        /// 獲取本地 IPv4 地址
        /// </summary>
        /// <returns></returns>
        static string GetLocalIPAddress()
        {
            // 獲取主機的所有地址列表
            var host = Dns.GetHostEntry(Dns.GetHostName());
            // 過濾出只包含 IPv4 地址的列表
            var lsIps = host.AddressList.Where(ip => ip.AddressFamily == AddressFamily.InterNetwork);
            // 優先選擇 10.x.x.x 開頭的 IPv4 地址
            var ipAddress = lsIps.Where(p => p.ToString().StartsWith("10.")).FirstOrDefault(); // 移除OrderDescending()[Linq 7版才有]
            // 如果沒有找到 10.x.x.x 開頭的 IPv4 地址,則返回空字符串
            if (ipAddress == null)
            {
                return "";
            }
            // 返回找到的 IPv4 地址
            return ipAddress.ToString();
        }

        /// <summary>
        /// 根據主機名稱或 IP 地址字符串獲取對應的 IPv4 和/或 IPv6 地址列表
        /// </summary>
        /// <param name="hostName"></param>
        /// <param name="ip4Wanted"></param>
        /// <param name="ip6Wanted"></param>
        /// <returns></returns>
        static IPAddress[] GetIPsByName(string hostName, bool ip4Wanted, bool ip6Wanted)
        {
            // 首先嘗試將輸入的字符串解析為 IP 地址
            IPAddress outIpAddress;
            if (IPAddress.TryParse(hostName, out outIpAddress) == true)
                // 如果成功解析,直接返回該 IP 地址
                return new IPAddress[] { outIpAddress };
            // 使用 Dns.GetHostAddresses() 獲取主機的所有 IP 地址列表
            IPAddress[] addresslist = Dns.GetHostAddresses(hostName);
            // 如果沒有找到任何 IP 地址,返回空數組
            if (addresslist == null || addresslist.Length == 0)
                return new IPAddress[0];
            // 根據輸入的 ip4Wanted 和 ip6Wanted 標誌,篩選並返回指定的 IP 地址列表
            if (ip4Wanted && ip6Wanted)
                // 如果同時需要 IPv4 和 IPv6,直接返回所有地址
                return addresslist;
            if (ip4Wanted)
                // 如果只需要 IPv4,使用 Where 過濾出 IPv4 地址
                return addresslist.Where(o => o.AddressFamily == AddressFamily.InterNetwork).ToArray();
            if (ip6Wanted)
                // 如果只需要 IPv6,使用 Where 過濾出 IPv6 地址
                return addresslist.Where(o => o.AddressFamily == AddressFamily.InterNetworkV6).ToArray();
            // 如果沒有指定需要任何類型的 IP 地址,返回空數組
            return new IPAddress[0];
        }        
        public static async Task<bool> AddWebdavLog(string Action,string UserIP, string UserAD,string FileID,string FileName,string Remark)
        {
            string apiUrl = InxSetting.ClsCommon.WebApiUrl + "/api";// 取得WebApiUrl
            string route = "/WebDav/WebdavLog";
            HttpClient client = new HttpClient();
            try
            {
               using HttpRequestMessage request = new HttpRequestMessage();
               {
                    request.Method = HttpMethod.Post;
                    request.RequestUri = new Uri(apiUrl + route);

                    // 將資料物件序列化成 JSON 字串
                    var jsonData = JsonConvert.SerializeObject(new { 
                        Action = Action, 
                        UserIP = UserIP,
                        UserAD = UserAD,
                        FileID = FileID,
                        FileName = FileName,
                        Remark = Remark
                    });
                    // 指定body
                    request.Content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                    HttpResponseMessage result = await client.SendAsync(request);
                    if (result.IsSuccessStatusCode)
                    {
                        string responseContent = await result.Content.ReadAsStringAsync();
                        return true; // 返回成功響應的內容
                    }
                    else
                    {
                        // 讀取錯誤響應內容
                        string errorContent = await result.Content.ReadAsStringAsync();
                        Console.WriteLine($"Error: {result.StatusCode}, Content: {errorContent}");
                        return false; // 或者根據需要返回錯誤信息
                    }
               }
            }
            catch (Exception ex)
            {

               Console.WriteLine("Error: " + ex.Message);
               return false; // 或者根據需要返回錯誤信息
            }
        }
    }
}
