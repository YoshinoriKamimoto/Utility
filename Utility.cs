using System;
using System.Linq;

/// <summary>
/// 共通処理を実装するクラス
/// </summary>
internal static class Utility
{
    /// <summary>
    /// 設定ファイルからログ出力先を取得するメソッド
    /// </summary>
    public static string GetLogFolderPath()
    {
        try
        {
            // XMLファイルを読み込む
            XElement xml = XElement.Load(Program.configPath);


            // 取得したい要素の存在チェック
            if (IsElement("Log") == false)
            {
                Error($"設定ファイルに<Log>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("LogPath") == false)
            {
                Error($"設定ファイルに<LogPath>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }


            // 要素の値を取得する
            XElement log = xml.Element("Log");
            string logPath = log.Element("LogPath").Value;
            return logPath;
        }
        catch
        {
            throw;
        }
    }


    /// <summary>
    /// 設定ファイルからDB接続先を取得するメソッド
    /// </summary>
    /// <returns>東京DB・大阪DBの接続文字列を返す、key:"東京" or "大阪",value：接続文字列</returns>
    public static Dictionary<string, string> GetConnectionStr()
    {
        try
        {
            // XMLファイルを読み込む
            XElement xml = XElement.Load(Program.configPath);


            // 取得したい要素の存在チェック
            if (IsElement("DB") == false)
            {
                Error($"設定ファイルに<DB>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("Server") == false)
            {
                Error($"設定ファイルに<Server>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("Port") == false)
            {
                Error($"設定ファイルに<Port>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("UserName") == false)
            {
                Error($"設定ファイルに<UserName>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("Password") == false)
            {
                Error($"設定ファイルに<Password>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("TokyoDataBase") == false)
            {
                Error($"設定ファイルに<TokyoDataBase>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("OsakaDataBase") == false)
            {
                Error($"設定ファイルに<OsakaDataBase>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }


            // 要素の値を取得する
            XElement db = xml.Element("DB");
            string server = db.Element("Server").Value;
            string port = db.Element("Port").Value;
            string userName = db.Element("UserName").Value;
            string password = db.Element("Password").Value;
            string tokyoDataBase = db.Element("TokyoDataBase").Value;
            string osakaDataBase = db.Element("OsakaDataBase").Value;

            // 接続文字列を定義
            string connectionStrTokyo = $"Server={server};Port={port};UserName={userName};Password={password};DataBase={tokyoDataBase};";
            string connectionStrOsaka = $"Server={server};Port={port};UserName={userName};Password={password};DataBase={osakaDataBase};";

            // 接続文字列をコレクションに格納して返却
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("Tokyo", connectionStrTokyo);
            dic.Add("Osaka", connectionStrOsaka);
            return dic;
        }
        catch
        {
            throw;
        }
    }



    /// <summary>
    /// 設定ファイルから対象商品アリスグループコードを取得するメソッド
    /// </summary>
    /// <returns>対象にするアリス商品グループコード文字列</returns>
    public static string GetGroupCode()
    {
        try
        {
            // XMLファイルを読み込む
            XElement xml = XElement.Load(Program.configPath);


            // 取得したい要素の存在チェック
            if (IsElement("ItemGroup") == false)
            {
                Error($"設定ファイルに<ItemGroup>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("GroupCode") == false)
            {
                Error($"設定ファイルに<GroupCode>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }


            // 要素の値を取得する
            XElement itemGroup = xml.Element("ItemGroup");
            string groupCode = itemGroup.Element("GroupCode").Value;

            return groupCode;
        }
        catch
        {
            throw;
        }
    }



    /// <summary>
    /// 設定ファイルからオーダーチェック対象を取得するメソッド
    /// </summary>
    /// <returns>オーダーチェック対象の工程コード</returns>
    public static string GetOrderCheckProcCode()
    {
        try
        {
            // XMLファイルを読み込む
            XElement xml = XElement.Load(Program.configPath);


            // 取得したい要素の存在チェック
            if (IsElement("OrderCheck") == false)
            {
                Error($"設定ファイルに<OrderCheck>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("ProcCode") == false)
            {
                Error($"設定ファイルに<ProcCode>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }


            // 要素の値を取得する
            XElement orderCheck = xml.Element("OrderCheck");
            string procCode = orderCheck.Element("ProcCode").Value;

            return procCode;
        }
        catch
        {
            throw;
        }
    }



    /// <summary>
    /// 設定ファイルからCSV出力先を取得するメソッド
    /// </summary>
    public static string GetCSVFolderPath()
    {
        try
        {
            // XMLファイルを読み込む
            XElement xml = XElement.Load(Program.configPath);


            // 取得したい要素の存在チェック
            if (IsElement("CSV") == false)
            {
                Error($"設定ファイルに<CSV>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("CSVPath") == false)
            {
                Error($"設定ファイルに<CSVPath>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }


            // 要素の値を取得する
            XElement csvInfo = xml.Element("CSV");
            string csvFolderPath = csvInfo.Element("CSVPath").Value;
            return csvFolderPath;
        }
        catch
        {
            throw;
        }
    }



    /// <summary>
    /// 設定ファイルからコピー元、コピー先パスを取得するメソッド
    /// </summary>
    /// <returns>コピー元パス、コピー先パス</returns>
    public static Dictionary<string, string> GetCopyFolderPath()
    {
        try
        {
            // XMLファイルを読み込む
            XElement xml = XElement.Load(Program.configPath);


            // 取得したい要素の存在チェック
            if (IsElement("Copy") == false)
            {
                Error($"設定ファイルに<Copy>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("SourcePath") == false)
            {
                Error($"設定ファイルに<SourcePath>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("TargetPath") == false)
            {
                Error($"設定ファイルに<TargetPath>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }


            // 要素の値を取得する
            XElement copyFile = xml.Element("Copy");
            string sourcePath = copyFile.Element("SourcePath").Value;
            string targetPath = copyFile.Element("TargetPath").Value;

            // コレクションに格納して返却
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("SourcePath", sourcePath);
            dic.Add("TargetPath", targetPath);
            return dic;
        }
        catch
        {
            throw;
        }
    }



    /// <summary>
    /// 設定ファイルからAccessファイルのパスを取得するメソッド
    /// </summary>
    public static Dictionary<string, string> GetAccessFilePath()
    {
        try
        {
            // XMLファイルを読み込む
            XElement xml = XElement.Load(Program.configPath);

            // 取得したい要素の存在チェック
            if (IsElement("Access") == false)
            {
                Error($"設定ファイルに<Copy>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("AccessPath") == false)
            {
                Error($"設定ファイルに<AccessPath>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }

            if (IsElement("AccessBackupPath") == false)
            {
                Error($"設定ファイルに<AccessBackupPath>タグが見つかりませんでした。\nファイルを確認してください。\n{Program.configPath}");
            }


            // 要素の値を取得する
            XElement accessFile = xml.Element("Access");
            string accessFilePath = accessFile.Element("AccessPath").Value;
            string accessFileBackupPath = accessFile.Element("AccessBackupPath").Value;

            // コレクションに格納して返却
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("AccessPath", accessFilePath);
            dic.Add("AccessBackupPath", accessFileBackupPath);
            return dic;
        }
        catch
        {
            throw;
        }
    }



    /// <summary>
    /// 要素が設定ファイルに存在するかチェックするメソッド
    /// </summary>
    /// <param name="tagName">チェックしたい要素のタグ名称</param>
    /// <returns>true：要素あり、false：要素なし</returns>
    public static bool IsElement(string tagName)
    {
        // XMLファイルを読み込む
        XmlDocument doc = new XmlDocument();
        doc.Load(Program.configPath);

        // パラメータで渡されたタグ名称の要素数を取得
        XmlNodeList lists = doc.GetElementsByTagName(tagName);

        // 取得した要素数をチェック
        if (lists.Count > 0)
        {
            // 要素あり
            return true;
        }
        else
        {
            // 要素なし
            return false;
        }
    }



    /// <summary>
    /// 明示的にエラーを発生させ、呼び出し元に投げるメソッド
    /// </summary>
    /// <param name="msg">通知したいエラーメッセージ</param>
    public static void Error(string msg)
    {
        Exception ex = new Exception($"\n{msg}\n\n");
        throw ex;
    }

}
