using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Mysqlx.Crud;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace AutoReturn
{
    public partial class Form1 : Form
    {
        private Settings _settings;
        private const string SettingsUrl = "https://techshop.alexbase.net//setting.json";

        /// <summary>
        /// Form1類的構造函數
        /// 初始化組件並設置Excel許可證上下文，加載應用程序設置
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            LoadSettingsAsync();
        }

        /// <summary>
        /// 從遠程JSON文件異步加載應用程序設置
        /// </summary>
        /// <returns>表示異步操作的Task</returns>
        private async Task LoadSettingsAsync()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var jsonResponse = await client.GetStringAsync(SettingsUrl);
                    _settings = JsonSerializer.Deserialize<Settings>(jsonResponse);

                    if (_settings == null)
                    {
                        throw new Exception("無法載入設定檔");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入設定檔時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        /// <summary>
        /// 文件路徑文本框的點擊事件處理程序
        /// 打開文件對話框以選擇Excel文件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件參數</param>
        private void txtFilePath_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = ofd.FileName;
                }
            }
        }

        /// <summary>
        /// 處理按鈕的點擊事件處理程序
        /// 處理選定的Excel文件並啟動退款操作
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件參數</param>
        private async void btnProcess_Click(object sender, EventArgs e)
        {
            txtFilePath.Enabled = false;
            btnStart.Enabled = false;

            string filePath = txtFilePath.Text.Trim();

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("請先選擇一個有效的 Excel 檔案！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                txtOutput.Text += $"\r\n計算退款筆數...\r\n";

                string orders = ReadExcelOrders(filePath);

                DataTable dbResult = GetOrderDataFromDatabase(orders);

                txtOutput.Text += $"\r\n共有{dbResult.Rows.Count}，需要退款...\r\n";

                if (dbResult.Rows.Count > 0)
                {
                    txtOutput.Text += "\r\n開始提交退款請求...\r\n";
                    foreach (DataRow row in dbResult.Rows)
                    {
                        string orderId = row["id"].ToString();
                        string order_no = row["order_no"].ToString();
                        string response = await SendRefundRequest(orderId, order_no);
                        if (response == "退款異常，token無效")
                        {
                            throw new Exception("退款異常，token無效");
                        }
                    }
                }
                else
                {
                    txtOutput.Text += "\r\n資料庫中無符合的訂單";
                }

                MessageBox.Show("批量退款完成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                WriteLog($"批量退款完成", "", "info");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtOutput.Text = $"發生錯誤: {ex.Message}";
            }
        }

        /// <summary>
        /// 寫入日誌記錄
        /// </summary>
        /// <param name="message">日誌消息內容</param>
        /// <param name="order_no">訂單編號</param>
        /// <param name="type">日誌類型，默認為info</param>
        private void WriteLog(string message, string order_no, string type = "info")
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string logMessage = $"[{timestamp}] [{type.ToUpper()}] 訂單編號: {order_no} - {message}";
                string fileName = $"refund_log_{DateTime.Now:yyyyMMdd}.txt";
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }

                string fullPath = Path.Combine(logPath, fileName);
                File.AppendAllText(fullPath, logMessage + Environment.NewLine);
            }
            catch
            {
                // 寫入日誌失敗時不拋出異常
            }
        }

        /// <summary>
        /// 發送退款請求
        /// </summary>
        /// <param name="orderId">訂單ID</param>
        /// <param name="order_no">訂單編號</param>
        /// <param name="maxRetries">最大重試次數，默認為3次</param>
        /// <returns>API響應內容</returns>
        private async Task<string> SendRefundRequest(string orderId, string order_no, int maxRetries = 3)
        {
            int currentRetry = 0;

            while (currentRetry <= maxRetries)
            {
                try
                {
                    var timeoutPeriod = TimeSpan.FromMinutes(20);

                    using (var cts = new CancellationTokenSource(timeoutPeriod))
                    using (var client = new HttpClient(new HttpClientHandler
                    {
                        UseCookies = false,
                        AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.Brotli
                    }))
                    {
                        client.Timeout = timeoutPeriod;

                        client.DefaultRequestHeaders.Add("Cookie", $"token={_settings.Token}");
                        client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
                        client.DefaultRequestHeaders.Add("Accept", "*/*");
                        client.DefaultRequestHeaders.Add("Connection", "keep-alive");

                        var formData = new FormUrlEncodedContent(new[]
                        {
                            new KeyValuePair<string, string>("t", "0.49321006292624636"),
                            new KeyValuePair<string, string>("id", orderId),
                            new KeyValuePair<string, string>("refund_third", "0")
                        });

                        HttpResponseMessage response = await client.PostAsync(_settings.ApiUrl, formData, cts.Token);
                        response.EnsureSuccessStatusCode();

                        string content = await response.Content.ReadAsStringAsync();

                        if(content.Contains("token无效，请重新登录"))
                        {
                            WriteLog($"退款異常，token無效", order_no, "error");
                            return "退款異常，token無效";
                        }

                        if (content.Contains("Lock wait timeout exceeded"))
                        {
                            if (currentRetry < maxRetries)
                            {
                                WriteLog($"退款請求遇到資料庫鎖定，即將進行第 {currentRetry + 1} 次重試", order_no, "warning");
                                await Task.Delay(TimeSpan.FromMinutes(10));
                                currentRetry++;
                                continue;
                            }
                        }
                        else if (content.Contains("订单不支持退款"))
                        {
                            WriteLog($"訂單已退款過或不支持退款", order_no, "warning");
                            return content;
                        }

                        WriteLog($"退款請求成功", order_no, "info");
                        return content;
                    }
                }
                catch (OperationCanceledException)
                {
                    string errorMsg = "退款請求成功";
                    WriteLog(errorMsg, order_no, "info");
                    return errorMsg;
                }
                catch (Exception ex)
                {
                    if (currentRetry < maxRetries)
                    {
                        WriteLog($"退款請求發生錯誤: {ex.Message}，即將進行第 {currentRetry + 1} 次重試", order_no, "error");
                        await Task.Delay(TimeSpan.FromMinutes(10));
                        currentRetry++;
                        continue;
                    }
                    string errorMsg = $"API 呼叫錯誤: {ex.Message}";
                    WriteLog(errorMsg, order_no, "error");
                    return errorMsg;
                }
            }

            string finalErrorMsg = $"退款請求失敗，已重試 {maxRetries} 次";
            WriteLog(finalErrorMsg, order_no, "error");
            return finalErrorMsg;
        }

        /// <summary>
        /// 讀取Excel文件中的訂單信息
        /// </summary>
        /// <param name="filePath">Excel文件路徑</param>
        /// <returns>訂單號列表字符串，格式為用逗號分隔的帶單引號的訂單號</returns>
        private string ReadExcelOrders(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                if (worksheet.Dimension == null)
                    return "Excel 檔案沒有內容";

                int rowCount = worksheet.Dimension.Rows;
                int orderColumnIndex = GetColumnIndex(worksheet, "訂單號");

                if (orderColumnIndex == -1)
                    return "找不到「訂單號」欄位";

                HashSet<string> orderNumbers = new HashSet<string>();

                for (int row = 2; row <= rowCount; row++)
                {
                    string order = worksheet.Cells[row, orderColumnIndex].Text.Trim();
                    if (!string.IsNullOrEmpty(order))
                        orderNumbers.Add(order);
                }

                return orderNumbers.Count > 0 ? string.Join(",", orderNumbers.Select(o => $"'{o}'")) : "無有效的訂單號";
            }
        }

        /// <summary>
        /// 獲取指定列名的列索引
        /// </summary>
        /// <param name="worksheet">Excel工作表</param>
        /// <param name="columnName">列名</param>
        /// <returns>列索引，如果未找到則返回-1</returns>
        private int GetColumnIndex(ExcelWorksheet worksheet, string columnName)
        {
            int colCount = worksheet.Dimension.Columns;
            for (int col = 1; col <= colCount; col++)
            {
                if (worksheet.Cells[1, col].Text.Trim() == columnName)
                    return col;
            }
            return -1;
        }

        /// <summary>
        /// 從數據庫獲取訂單數據
        /// </summary>
        /// <param name="orderNumbers">訂單號列表字符串，格式為用逗號分隔的帶單引號的訂單號</param>
        /// <returns>包含訂單信息的DataTable</returns>
        private DataTable GetOrderDataFromDatabase(string orderNumbers)
        {
            DataTable dt = new DataTable();

            if (string.IsNullOrEmpty(orderNumbers) || orderNumbers == "無有效的訂單號")
                return dt;

            string query = $"SELECT order_no, id FROM lt_order WHERE order_no IN ({orderNumbers});";

            using (MySqlConnection conn = new MySqlConnection(_settings.ConnectionString))
            {
                try
                {
                    conn.Open();
                    using (MySqlCommand cmd = new MySqlCommand(query, conn))
                    {
                        using (MySqlDataAdapter adapter = new MySqlDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"資料庫查詢錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return dt;
        }
    }
}