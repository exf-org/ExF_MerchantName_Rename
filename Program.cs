using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Collections.Specialized;
using System.Collections;
using System.Data;
using ExF_iniFileCS;
//プロジェクト→参照の追加
//プロジェクト→NuGetパッケージマネージャー→パッケージの管理→System.Data.SQLite

//【iniFileの使用方法】
//「ExF_iniFile.cs」を「Program.cs」と同じフォルダにコピー
//プロジェクト＞既存項目の追加＞ExF_iniFile.cs
//「Program.cs」上に「using ExF_iniFileCS;」を追加
//「Program.cs」上にインスタンスを生成　変数「var ExF = new ExF_iniFile();」　※newはインスタンス生成 変数名は好きに設定OK
//「Program.cs」上に変数を作成「Encoding CharacterCode = Encoding.GetEncoding("shift_jis");」
//「Program.cs」上で文字コードを指定「ExF.CharacterCode = CharacterCode;」
//「Program.cs」上でiniFileを呼び出す「ExF.ReadINI(iniFilePath);」
//「Program.cs」上でセクション一覧を取得 連想配列「ExF.dicSECs.Keys」
//「Program.cs」上でセクションからKEYに対応したVALを取得　KEYの一覧を連想配列に格納「var dicKeys = (OrderedDictionary)ExF.dicSECs[SECName];」
//Program.cs」上でKEYの一覧から値を取得「dicKeys["KEY名"].ToString();」

namespace ExF_MerchantName_Rename
{
    internal class Program
    {

        public static string ErrorMsgFlg = "1";

        public static string ChangeFLG = "0";

        public static int ret = 0;

        //ログを書き出すためのStringBuilderを宣言
        public static StringBuilder SB_log = new StringBuilder();

        //dic_Outの宣言
        private static OrderedDictionary dic_Out = new OrderedDictionary();

        //出力エンコーディング
        private static Encoding sjis = Encoding.GetEncoding("shift_jis");

        //現在時刻
        public static DateTime DT = DateTime.Now;

        //ログフォルダパスの宣言
        public static string LogFolderPath = "";

        //ログファイルパスの宣言
        public static string LogFilePath = "";

        //iniFileの宣言
        private static ExF_iniFile ExF_iniFile = new ExF_iniFile();

        //iniFilePathの宣言
        private static string iniFilePath = "";

        static int Main(string[] args)
        {
            var ErrMsg = "";

            //iniFilePathが指定されていなければエラー
            if (args[0] == "")
            {
                ErrMsg = "batにiniFilePathが指定されていません";
                CreateLog(ErrMsg);
                MessageBox.Show(ErrMsg, "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

                ret = 1;
                Environment.Exit(ret);
            }
            else
            {
                //フルパスに変換
                iniFilePath = Path.GetFullPath(args[0]);
            }

            //LogFolderPathが指定されていなければエラー
            if (args[1] == "")
            {
                ErrMsg = "batにLogFolderPathが指定されていません";
                CreateLog(ErrMsg);
                MessageBox.Show(ErrMsg, "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

                ret = 1;
                Environment.Exit(ret);
            }
            else
            {
                //フルパスに変換
                LogFolderPath = Path.GetFullPath(args[1]);
            }

            //ログファイルパスの生成
            LogFilePath = Path.Combine(LogFolderPath, Path.GetFileNameWithoutExtension(System.Windows.Forms.Application.ExecutablePath) + "_" + DT.ToString("yyyyMMdd_HHmmss") + ".log");

            CreateLog($"-----------------------------------------------------------------");
            CreateLog($"処理開始");

            //フォルダの存在チェック
            //①batファイルに指定されたパスが存在しなければエラー---------------------------------------------------------
            if (Directory.Exists(LogFolderPath))
            {
            }
            else
            {
                ErrMsg = "batで指定されたLogFolderPathが存在しません";
                MessageBox.Show(ErrMsg, "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
                ret = 1;
                //プログラムを終了させる
                Environment.Exit(ret);
            }

            //iniファイルとして指定されたファイルが存在しなければエラー
            if (File.Exists(iniFilePath))
            {
                //正常
            }
            else
            {
                ErrMsg = "iniファイルとして指定されたファイルが存在しません";
                CreateLog(ErrMsg);
                MessageBox.Show(ErrMsg, "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
                File.WriteAllText(LogFilePath, SB_log.ToString(), sjis);
                ret = 1;
                Environment.Exit(ret);
            }

            //iniFileの読み込み
            ExF_iniFile.ReadINI(iniFilePath);

            //iniファイルのセクションのキーを宣言
            string InFolderPath = "";
            string OutFolderPath = "";
            string DBFolderPath = "";
            string MerchantCode_Col = "";
            string MerchantName_Col = "";
            string StoreName_Col = "";

            //■iniFileの値を変数に格納
            InFolderPath = ((OrderedDictionary)ExF_iniFile.dicSECs["設定"])["InFolderPath"].ToString();
            OutFolderPath = ((OrderedDictionary)ExF_iniFile.dicSECs["設定"])["OutFolderPath"].ToString();
            DBFolderPath = ((OrderedDictionary)ExF_iniFile.dicSECs["設定"])["DBFolderPath"].ToString();
            MerchantCode_Col = ((OrderedDictionary)ExF_iniFile.dicSECs["設定"])["MerchantCode_Col"].ToString();
            MerchantName_Col = ((OrderedDictionary)ExF_iniFile.dicSECs["設定"])["MerchantName_Col"].ToString();
            StoreName_Col = ((OrderedDictionary)ExF_iniFile.dicSECs["設定"])["StoreName_Col"].ToString();

            //InFolderPathのディレクトリが存在しない場合はエラー--------------------------------------------------------------
            if (!Directory.Exists(InFolderPath))
            {
                ErrMsg = "iniFileで指定されたInFolderPathが存在しません";
                CreateLog(ErrMsg);
                MessageBox.Show(ErrMsg, "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
                File.WriteAllText(LogFilePath, SB_log.ToString(), sjis);
                ret = 1;
                //プログラムを終了させる
                Environment.Exit(ret);
            }
            //----------------------------------------------------------------------------------------------------------------------

            //OutFolderPathのディレクトリが存在しない場合はエラー--------------------------------------------------------------
            if (!Directory.Exists(OutFolderPath))
            {
                ErrMsg = "iniFileで指定されたOutFolderPathが存在しません";
                CreateLog(ErrMsg);
                MessageBox.Show(ErrMsg, "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
                File.WriteAllText(LogFilePath, SB_log.ToString(), sjis);
                ret = 1;
                //プログラムを終了させる
                Environment.Exit(ret);
            }
            //----------------------------------------------------------------------------------------------------------------------

            //DBFolderPathのディレクトリが存在しない場合はエラー--------------------------------------------------------------
            if (!Directory.Exists(DBFolderPath))
            {
                ErrMsg = "iniFileで指定されたDBFolderPathが存在しません";
                CreateLog(ErrMsg);
                MessageBox.Show(ErrMsg, "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
                File.WriteAllText(LogFilePath, SB_log.ToString(), sjis);
                ret = 1;
                //プログラムを終了させる
                Environment.Exit(ret);
            }
            //----------------------------------------------------------------------------------------------------------------------

            //ファイルを探しに行く
            string[] infiles = Directory.GetFiles(InFolderPath, "*.t*");
            if (infiles.Length >= 1)
            {
                //正常パターン
                CreateLog($"対象のフォルダ内に{infiles.Length}件のテキストファイルが見つかりました。");

            }
            else 
            {
                ErrMsg = "対象のフォルダ内にテキストファイルが1件も見つかりませんでした。";
                CreateLog(ErrMsg);
                MessageBox.Show(ErrMsg, "エラー",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
                File.WriteAllText(LogFilePath, SB_log.ToString(), sjis);
                ret = 1;
                //プログラムを終了させる
                Environment.Exit(ret);
            }

            //-----------------------------------------------------------------
            try
            {
                CreateLog($"DBの処理開始");
                //DBを開く
                //DB接続設定
                var SynchronizationModes = new SynchronizationModes();
                var SQLiteJournalModeEnum = new SQLiteJournalModeEnum();

                SynchronizationModes = SynchronizationModes.Normal;

                //Write - Ahead Logging(WAL) モードを意味する
                SQLiteJournalModeEnum = SQLiteJournalModeEnum.Wal;

                //DBの名前を指定。ターゲットはTHINCA_COMMON.sqlite3
                var MainDBFileName = "THINCA_COMMON.sqlite3";

                //DBがあるフォルダパスとDBの名前を繋げたものが絶対パスとなる
                var MainDBFilePath = Path.Combine(DBFolderPath, MainDBFileName);

                //SQLiteデータベースへの接続文字列を構築するためのクラス
                var SQLiteCSB = new SQLiteConnectionStringBuilder()
                {
                    Version = 3,
                    DataSource = MainDBFilePath.ToString(),
                    SyncMode = SynchronizationModes,
                    JournalMode = SQLiteJournalModeEnum
                };

                //M_TERMINALのCREATE_DATEを格納するディクショナリーを作成する
                var dic_MerchantName = new OrderedDictionary();

                using (SQLiteConnection conn = new SQLiteConnection(SQLiteCSB.ToString()))
                {
                    conn.Open();

                    var dicDBPaths = new OrderedDictionary();

                    //複数のDBをアタッチする
                    foreach (DictionaryEntry x in dicDBPaths)
                    {

                        if (dicDBPaths["THINCA_COMMON.sqlite3"].ToString() != x.Value.ToString())
                        {
                            using (var cmd = new SQLiteCommand(conn))
                            {
                                cmd.CommandText = "ATTACH DATABASE '" + x.Value.ToString() + "' AS " + x.Key.ToString() + ";";
                                cmd.ExecuteNonQuery();
                            }
                        }

                    }

                    //SQLでの検索開始
                    try
                    {
                        //SQLコマンド生成
                        var sql = new List<string>();
                        sql.Add("SELECT");
                        sql.Add("MERCHANT_CD,MERCHANT_NAME");
                        sql.Add("FROM");
                        sql.Add("M_MERCHANT");
                        sql.Add("WHERE");
                        sql.Add("MERCHANT_CD LIKE '00000002001924%'");//(STORES_924の全加盟店)
                        sql.Add("ORDER BY");
                        sql.Add("MERCHANT_CD ASC");

                        //コマンドを繋げてSQL文を作成する
                        var SQLStr = String.Join(" ", sql);

                        using (SQLiteCommand cmd = conn.CreateCommand())
                        {
                            var SQLiteCommand = new SQLiteCommand();
                            SQLiteCommand.CommandText = SQLStr;
                            SQLiteCommand.Connection = conn;
                            SQLiteCommand.CommandType = CommandType.Text;

                            // Stopwatchクラス生成
                            var sw = new System.Diagnostics.Stopwatch();

                            // 計測開始
                            sw.Start();

                            //SQLの実行。1秒未満で終わる
                            using (var reader = SQLiteCommand.ExecuteReader())

                            {
                                // 計測停止
                                sw.Stop();

                                // 結果表示
                                CreateLog($"■SQL実行にかかった時間");

                                TimeSpan ts = sw.Elapsed;
                                CreateLog($"{ts.Hours}時間 {ts.Minutes}分 {ts.Seconds}秒 {ts.Milliseconds}ミリ秒");

                                var recNo = 0;

                                if (reader.HasRows == true)
                                {
                                    //レコードが取得できた場合
                                    while (reader.Read())
                                    {

                                        recNo++;

                                        var MERCHANT_CD = reader.GetValue(0);
                                        var MERCHANT_NAME = reader.GetValue(1);

                                        dic_Out.Add(MERCHANT_CD, MERCHANT_NAME);

                                    }

                                    CreateLog($"{recNo} 件の加盟店データを M_MERCHANT から取得しました。");
                                    CreateLog($"DBの処理終了");
                                }

                            }

                        }

                    }

                    catch (Exception)
                    {
                        ret = 1;
                        throw;
                    }

                    //finallyとは例外の有無に関わらず、最後に必ず実行される処理
                    finally
                    {
                        //DBを閉じる
                        conn.Close();
                    }

                }

            //SQL関連処理終了
            }
            catch (Exception)
            {
                ret = 1;
                throw;
            }
            //-----------------------------------------------------------------

            //OutFolderPathの初期化(削除)
            DeleteAllContents(OutFolderPath);
            CreateLog($"OutFolderをリセット（削除）しました。");

            //InFolderPathにあるテキストファイルを全て取得し、OutFolderPathへ移動する
            string[] TextFiles = Directory.GetFiles(InFolderPath, "*.t*");

            foreach (string TextFilePath in TextFiles)
            {
                //ファイル名だけを取得
                string TextFileName = Path.GetFileName(TextFilePath);
                //コピー先パスを生成
                string OutFilePath = Path.Combine(OutFolderPath, TextFileName);

                //コピー（上書き不可）
                File.Copy(TextFilePath, OutFilePath, overwrite: false);
            }
            //-----------------------------------------------------------------

            //ファイルを１行単位で読み込んでいく。文字コードは「Shift_JIS」
            //ファイルを探しに行く----------------------------------------------------------------------------------------------------------------
            string[] outfiles = Directory.GetFiles(OutFolderPath, "*.t*");

            foreach (string file in outfiles)
            {
                //ファイル名を取得
                string fileName = Path.GetFileName(file);

                List<string> AllLines = new List<string>();

                CreateLog($"■ファイル処理開始-----------------------------------");
                CreateLog($"ファイル名：{ fileName}");

                //ファイルを１行単位で読み込んでいく。文字コードは「Shift_JIS」
                StreamReader sr = new StreamReader(file, sjis);

                //2行目から読み込みたいのでヘッダー部分は読み飛ばす
                string line = sr.ReadLine();

                //ヘッダー部分は読み飛ばすが、書き込みはしたいので結果リストに追加
                AllLines.Add(line);

                //行数のカウントのために使用
                int No = 0;

                //空になるまで読んでいく
                while ((line = sr.ReadLine()) != null)
                {
                    //行数カウントをプラスしていく
                    No++;

                    //読み込んだinputファイルの行をタブ区切りで分解する
                    var splitline = Microsoft.VisualBasic.Strings.Split(line, "\t");

                    //MerchantCode_Colをint型に変換しその列の値をMerchantCode_Col_Indexという文字列とする
                    string MerchantCode_Col_Index = splitline[int.Parse(MerchantCode_Col) - 1];
                    //MerchantName_Colをint型に変換しその列の値をMerchantName_Col_Indexという文字列とする
                    string MerchantName_Col_Index = splitline[int.Parse(MerchantName_Col) - 1];
                    //StoreName_Colをint型に変換しその列の値をStoreName_Col_Index という文字列とする
                    string StoreName_Col_Index = splitline[int.Parse(StoreName_Col) - 1];

                    if (MerchantCode_Col_Index.Length == 0)
                    {
                        //新規の想定                       
                        //NULLなら書き換え不要。そのまま新しい行を作成する。実際のファイルへの書き込みはファイルを閉じてから行う。
                        AllLines.Add(line);
                    }
                    else if (MerchantCode_Col_Index.Length == 20)
                    {
                        //端末追加・ブランド追加の想定。加盟店を書き換える必要あり
                        //加盟店コードを見てdic_Outに含まれていれば
                        if (dic_Out.Contains(MerchantCode_Col_Index))
                        {

                            if (MerchantName_Col_Index == dic_Out[MerchantCode_Col_Index].ToString())
                            {
                                //加盟店名が一致しているため何もしない
                                AllLines.Add(line);
                            }
                            else
                            {

                                ChangeFLG = "1";
                                CreateLog($"依頼書 {No} レコード目：加盟店コード({MerchantCode_Col_Index})「{MerchantName_Col_Index}」→ DB：「{dic_Out[MerchantCode_Col_Index]}」");

                                //加盟店名が一致していないためDBの加盟店名に書き換える。設置店名も書き換える。加盟店名=設置店名の想定。
                                splitline[int.Parse(MerchantName_Col) - 1] = dic_Out[MerchantCode_Col_Index].ToString();//加盟店名
                                splitline[int.Parse(StoreName_Col) - 1] = dic_Out[MerchantCode_Col_Index].ToString();//設置店名

                                //書き換えた加盟店名で改めて新しい行を作成する。実際のファイルへの書き込みはファイルを閉じてから行う。
                                string newLine = string.Join("\t", splitline);
                                AllLines.Add(newLine);

                            }

                        }
                        else
                        {
                            //20桁の加盟店コードの記載はあるが、dic_Outに一致する加盟店コードが存在しない
                            ErrMsg = $"20桁の加盟店コードの記載はありますが、DBに一致する加盟店コードが存在しません。DBを更新してください。" + 
                            ($"依頼書 {No} レコード目：加盟店コード({MerchantCode_Col_Index})「{MerchantName_Col_Index}」→ DB：「見つかりませんでした」");
                            CreateLog(ErrMsg);
                            MessageBox.Show(ErrMsg, "エラー",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.DefaultDesktopOnly);
                            File.WriteAllText(LogFilePath, SB_log.ToString(), sjis);
                            ret = 1;
                            //プログラムを終了させる
                            Environment.Exit(ret);
                        }
                    }
                    else
                    {
                        //記載はあるが、20桁ではないためエラー
                        ErrMsg = $"加盟店コードの記載はありますが、20桁ではありません。" +
                        ($"依頼書 {No} レコード目：加盟店コード({MerchantCode_Col_Index})「{MerchantName_Col_Index}」→{MerchantCode_Col_Index.Length}桁");
                        CreateLog(ErrMsg);
                        MessageBox.Show(ErrMsg, "エラー",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly);
                        File.WriteAllText(LogFilePath, SB_log.ToString(), sjis);
                        ret = 1;
                        //プログラムを終了させる
                        Environment.Exit(ret);
                    }

                    //1ファイルの読込終了
                }

                //inputファイルを閉じる
                sr.Close();

                //最後に加盟店名書き換え後のnewLineで上書き
                File.WriteAllLines(file, AllLines, sjis);

                CreateLog($"■ファイル処理終了-----------------------------------");

            }

            //ファイル読込終了-----------------------------------------------------------------
            if (ChangeFLG == "1")
            {
                CreateLog($"加盟店名不一致のファイルが存在したため書き換えました。");
            }
            else
            {
                CreateLog($"加盟店名不一致のファイルは存在しませんでした。");
            }

            CreateLog($"処理終了", "Status：" + ret.ToString());
            CreateLog($"-----------------------------------------------------------------");

            File.WriteAllText(LogFilePath, SB_log.ToString(), sjis);

            //戻り値ret
            return ret;

            //mainの終わり
        }

        //ログメッセージを作成し、それを SB_log とコンソールに出力するメソッド
        static string CreateLog(string key,
                                   string str = "",
                                   string[] arry = null)
        {
            try
            {

                var NowDateTime = DateTime.Now;
                var ls = new List<string>();

                ls.Add(NowDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                ls.Add(key);

                //もし str パラメータが空でない場合、それもログメッセージに追加する
                if (str != "")
                {
                    ls.Add(str);
                }

                if (arry != null)
                {
                    ls.Add(String.Join("\t", arry));
                }

                var log = String.Join("\t", ls.ToArray());
                SB_log.AppendLine(log);
                Console.WriteLine(log);

                return log;

            }
            catch (Exception)
            {
                ret = 1;
                throw;
            }
        }

        //指定したフォルダ内の全ファイル・サブフォルダを削除
        static void DeleteAllContents(string folderPath)
        {
            //フォルダ内のすべてのファイルを取得し削除
            foreach (string file in Directory.GetFiles(folderPath))
            {
                File.Delete(file);
            }

            // フォルダ内のすべてのサブフォルダを取得し削除
            foreach (string subFolder in Directory.GetDirectories(folderPath))
            {
                Directory.Delete(subFolder, true); //trueにすることで再帰的に削除
            }
        }

    }
}
