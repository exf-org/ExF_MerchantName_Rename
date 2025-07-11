using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using System.Collections;
using Microsoft.VisualBasic;
//using System.Windows.Forms;
//コマンドライン引数に→を追加".\ExF_CertificateTextRename.ini"

namespace ExF_iniFileCS
{

    public class ExF_iniFile
    {
        const bool DebugFlg = true;
        // iniFileの内容を格納
        public OrderedDictionary dicSECs = new OrderedDictionary();
        public Encoding CharacterCode = Encoding.GetEncoding("shift_jis");
        private List<string> SectionNames = new List<string>();

        private readonly string ClassName = "ExF_iniFile";// GetType(ExF_iniFile).FullName;
        private void DebugPrint(string printmsg, bool flg)
        {
            if (flg == true)
                Debug.Print(printmsg);
        }
        /// <summary>
        ///     ''' ★iniFileよりSection一覧を取得しList(Of String)へ格納★
        ///     ''' </summary>
        ///     ''' <param name="filepath"></param>
        ///     ''' <returns></returns>
        private List<string> GetIniFileSectionNames(string filepath)
        {
            List<string> ret = new List<string>();

            try
            {
                using (System.IO.StreamReader sr = new System.IO.StreamReader(filepath, CharacterCode))
                {
                    while (sr.Peek() >= 0)
                    {
                        string c1 = ""; // 先頭一文字
                        string c2 = ""; // 後方一文字
                        string s = Strings.Trim(sr.ReadLine()); // 左右の空白を削除

                        if (s != "")
                        {
                            c1 = Strings.Left(s, 1);
                            c2 = Strings.Right(s, 1);

                            if (c1 == "#" || c1 == ";" || c1 == "'")
                            {
                            }
                            else if (c1 == "[" && c2 == "]")
                                // セクション行
                                ret.Add(Strings.Trim(Strings.Mid(s, 2, s.Length - 2))); // 左右の[]を外す
                            else
                            {
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret.Clear();

                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したSectionからキーの一覧を取得★
        ///     ''' </summary>
        ///     ''' <param name="filepath"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <returns></returns>
        private List<string> GetIniFileKeyNames(string filepath, string TargetSectionName)
        {
            List<string> ret = new List<string>();
            int flg = 0;

            try
            {
                using (System.IO.StreamReader sr = new System.IO.StreamReader(filepath, CharacterCode))
                {
                    string SectionName = "";

                    while (sr.Peek() >= 0)
                    {
                        string c1 = ""; // 先頭一文字
                        string c2 = ""; // 後方一文字
                        string s = Strings.Trim(sr.ReadLine()); // 左右の空白を削除

                        if (s != "")
                        {
                            c1 = Strings.Left(s, 1);
                            c2 = Strings.Right(s, 1);

                            if (c1 == "#" || c1 == ";" || c1 == "'")
                            {
                            }
                            else if (c1 == "[" && c2 == "]")

                                // セクション行
                                SectionName = Strings.Trim(Strings.Mid(s, 2, s.Length - 2)); // 左右の[]を外す
                            else if (TargetSectionName == SectionName)
                            {

                                // キー行
                                string[] key = Strings.Split(s, "=", -1, Constants.vbBinaryCompare);
                                key[0] = Strings.Trim(key[0]);
                                key[1] = Strings.Trim(key[1]);
                                ret.Add(key[0]);
                            }
                            else
                            {
                            }

                            if (flg == 1 && TargetSectionName != SectionName)
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret.Clear();
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したSectionとキーから値を取得★
        ///     ''' </summary>
        ///     ''' <param name="filepath"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <param name="TargetKeyName"></param>
        private string GetIniFileKeyValue(string filepath, string TargetSectionName, string TargetKeyName)
        {
            string ret = "";

            try
            {
                using (System.IO.StreamReader sr = new System.IO.StreamReader(filepath, CharacterCode))
                {
                    string SectionName = "";

                    while (sr.Peek() >= 0)
                    {
                        string c1 = ""; // 先頭一文字
                        string c2 = ""; // 後方一文字
                        string s = Strings.Trim(sr.ReadLine()); // 左右の空白を削除

                        if (s != "")
                        {
                            c1 = Strings.Left(s, 1);
                            c2 = Strings.Right(s, 1);

                            if (c1 == "#" || c1 == ";" || c1 == "'")
                            {
                            }
                            else if (c1 == "[" && c2 == "]")

                                // セクション行
                                SectionName = Strings.Trim(Strings.Mid(s, 2, s.Length - 2)); // 左右の[]を外す
                            else if (TargetSectionName == SectionName)
                            {

                                // 指定セクションと同一
                                string[] key = Strings.Split(s, "=", -1, Constants.vbBinaryCompare);
                                key[0] = Strings.Trim(key[0]);
                                key[1] = Strings.Trim(key[1]);
                                // 指定キーと同一
                                if (TargetKeyName == key[0])
                                {
                                    ret = key[1];
                                    break;
                                }
                            }
                            else
                            {
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ret = "";
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★iniFileからdicSECsを生成★
        ///     ''' ※正規表現を使用しセクション
        ///     ''' </summary>
        ///     ''' <param name="filepath"></param>
        public void ReadINI(string filepath)
        {

            // OrderedDictionaryの初期化
            dicSECs.Clear();

            try
            {

                // セクションの一覧を取得
                SectionNames = GetIniFileSectionNames(filepath);

                foreach (string SECName in SectionNames)
                {
                    // キーの一覧格納先
                    List<string> KeyNames = GetIniFileKeyNames(filepath, SECName);
                    // セクション単位のキー用OrderedDictionary
                    OrderedDictionary dicSECKEYs = new OrderedDictionary();

                    foreach (string key in KeyNames)
                    {
                        // セクションとキー名から値を取得
                        string KeyValue = GetIniFileKeyValue(filepath, SECName, key);

                        // キー用のOrderedDictionaryへ格納　※キーが存在しない場合
                        if (dicSECKEYs.Contains(key) == false)
                            dicSECKEYs.Add(key, KeyValue);
                        else
                        {
                            // INI KEYが重複している。同一[SECTION]でKEYが同じなのでエラーとする
                            string ErrorMsg;
                            ErrorMsg = "SECTION名[" + SECName + "]内で、同一のKEY(" + key + ")が定義されています。";
                            throw new ApplicationException(ErrorMsg);
                        }
                    }

                    if (dicSECs.Contains(SECName) == false)
                        dicSECs.Add(SECName, dicSECKEYs);
                    else
                    {

                        // INI [SECTION]が重複している
                        string ErrorMsg;
                        ErrorMsg = "SECTION名[" + SECName + "]が重複しています。";
                        throw new ApplicationException(ErrorMsg);
                    }
                }
            }
            catch (Exception ex)
            {
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryの中にあるセクションから値のリストを取得する★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <returns></returns>
        public List<string> GetKeyVals(OrderedDictionary dics, string TargetSectionName, string KeyFilter = "")
        {
            List<string> ret = new List<string>();

            try
            {

                OrderedDictionary Targetdic = (OrderedDictionary)dics[TargetSectionName];

                //int i = -1;

                foreach (DictionaryEntry n in Targetdic)
                {
                    Console.WriteLine(n.Key);   // キー
                    Console.WriteLine(n.Value); // 値

                    string KeyName = (String)n.Key;
                    if (KeyFilter == "*" || KeyFilter == "" || Regex.IsMatch(KeyName, KeyFilter) == true)
                        ret.Add((String)n.Value);
                }

            }
            catch (Exception ex)
            {
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryの中にあるセクションからキーリストを取得する★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <returns></returns>
        public List<string> GetKeyNames(OrderedDictionary dics, string TargetSectionName, string KeyFilter = "")
        {
            List<string> ret = new List<string>();

            try
            {
                ICollection keyCollection = ((OrderedDictionary)dics[TargetSectionName]).Keys;

                foreach (string KeyName in keyCollection)
                {
                    if (KeyFilter == "*" || KeyFilter == "" || Regex.IsMatch(KeyName, KeyFilter) == true)
                            ret.Add(KeyName);
                    }
                }

            catch (Exception ex)
            {
                ret.Clear();
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryの中にあるセクションリストを取得する★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <returns></returns>
        public List<string> GetSectionNames(OrderedDictionary dics, string SectionFilter = "")
        {
            List<string> ret = new List<string>();

            try
            {
                ICollection keyCollection = dics.Keys;

                foreach (string SECName in keyCollection)
                {
                    if (SectionFilter == "*" || SectionFilter == "" || Regex.IsMatch(SECName, SectionFilter) == true)
                        ret.Add(SECName);
                }
            }
            catch (Exception ex)
            {
                ret.Clear();
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryの中にあるセクションを抽出してコピーする★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="SectionFilter"></param>
        ///     ''' <returns></returns>
        public OrderedDictionary DictionaryCopy(OrderedDictionary dics, string SectionFilter = "", string KeyFilter = "")
        {
            OrderedDictionary ret = new OrderedDictionary();

            try
            {
                ICollection keyCollection1 = dics.Keys;

                foreach (string SECName in keyCollection1)
                {
                    if (SectionFilter == "*" || SectionFilter == "" || Regex.IsMatch(SECName, SectionFilter) == true)
                    {
                        OrderedDictionary dickeys;
                        dickeys = (OrderedDictionary)dics[SECName];

                        ICollection keyCollection2 = dickeys.Keys;

                        // Filter用の新しいOrderedDictionary
                        OrderedDictionary NewDicKeys = new OrderedDictionary();

                        foreach (string KEYName in keyCollection2)
                        {
                            if (KeyFilter == "*" || KeyFilter == "" || Regex.IsMatch(KEYName, KeyFilter) == true)
                                NewDicKeys.Add(KEYName, dickeys[KEYName]);
                        }

                        ret.Add(SECName, NewDicKeys);
                    }
                }
            }
            catch (Exception ex)
            {
                ret.Clear();
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            return ret;
        }

        public OrderedDictionary DictionaryCopy_UnmatchSection(OrderedDictionary dics, string SectionFilter = "", string KeyFilter = "")
        {
            OrderedDictionary ret = new OrderedDictionary();

            try
            {
                ICollection keyCollection1 = dics.Keys;

                foreach (string SECName in keyCollection1)
                {
                    if (Regex.IsMatch(SECName, SectionFilter) == false)
                    {
                        OrderedDictionary dickeys;
                        dickeys = (OrderedDictionary)dics[SECName];

                        ICollection keyCollection2 = dickeys.Keys;

                        // Filter用の新しいOrderedDictionary
                        OrderedDictionary NewDicKeys = new OrderedDictionary();

                        foreach (string KEYName in keyCollection2)
                        {
                            if (KeyFilter == "*" || KeyFilter == "" || Regex.IsMatch(KEYName, KeyFilter) == true)
                                NewDicKeys.Add(KEYName, dickeys[KEYName]);
                        }

                        ret.Add(SECName, NewDicKeys);
                    }
                }
            }
            catch (Exception ex)
            {
                ret.Clear();
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            return ret;
        }

        /// <summary>
        ///     ''' ★指定したOrderedDictionaryの中に指定したセクションとキーが存在するか判定する★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <param name="TargetKeyName"></param>
        ///     ''' <returns></returns>
        public bool Exist_Key(OrderedDictionary dics, string TargetSectionName, string TargetKeyName)
        {
            bool ret = false;

            try
            {
                if (dics.Contains(TargetSectionName) == false)
                    ret = false;
                else if (((OrderedDictionary)dics[TargetSectionName]).Contains(TargetKeyName) == false)
                    ret = false;
                else if (((OrderedDictionary)dics[TargetSectionName]).Contains(TargetKeyName) == true)
                    ret = true;
                else
                {
                }
            }
            catch (Exception ex)
            {
                ret = false;
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryの中に指定したセクションが存在するか判定する★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <returns></returns>
        public bool Exist_Section(OrderedDictionary dics, string TargetSectionName)
        {
            bool ret = false;

            try
            {
                if (dics.Contains(TargetSectionName) == true)
                    ret = true;
                else
                {
                }
            }
            catch (Exception ex)
            {
                ret = false;
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryの中にあるキーに値を設定する★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <param name="TargetKeyName"></param>
        ///     ''' <param name="TargetKeyVal"></param>
        ///     ''' <returns></returns>
        public int SetSectionKeyVal(ref OrderedDictionary dics, string TargetSectionName, string TargetKeyName, string TargetKeyVal)
        {
            int ret = 0;

            try
            {
                if (dics.Contains(TargetSectionName) == false)
                {

                    // セクションが存在しない場合
                    OrderedDictionary dicKeys = new OrderedDictionary();
                    dicKeys.Add(TargetKeyName, TargetKeyVal);
                    dics.Add(TargetSectionName, dicKeys);
                    ret = 1;
                }
                else if (((OrderedDictionary)dics[TargetSectionName]).Contains(TargetKeyName) == false)
                {

                    // セクションは存在しキーが存在しない場合
                    ((OrderedDictionary)dics[TargetSectionName]).Add(TargetKeyName, TargetKeyVal);
                    ret = 1;
                }
                else if (((OrderedDictionary)dics[TargetSectionName]).Contains(TargetKeyName) == true)
                {

                    // セクションとキー共に存在する場合 ★★★★★★★
                    ((OrderedDictionary)dics[TargetSectionName])[TargetKeyName] = TargetKeyVal;
                    ret = 1;
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                ret = ex.HResult;
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryの中にあるセクションやキーを削除する★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <param name="TargetKeyName"></param>
        ///     ''' <returns></returns>
        public int DeleteSectionKey(ref OrderedDictionary dics, string TargetSectionName, string TargetKeyName = "")
        {
            int ret = 0;

            try
            {
                if (dics.Contains(TargetSectionName) == true && TargetKeyName == "")
                {

                    // セクションを削除
                    dics.Remove(TargetSectionName);
                    ret = 1;
                }
                else if (dics.Contains(TargetSectionName) == true && ((OrderedDictionary)dics[TargetSectionName]).Contains(TargetKeyName) == true)
                {

                    // セクション内のキーを削除
                    ((OrderedDictionary)dics[TargetSectionName]).Remove(TargetKeyName);
                    ret = 1;
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                ret = ex.HResult;
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryのセクションとキーをiniFileと比較しながら新しいファイルに書き出す★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="filepath"></param>
        public void WriteINI(OrderedDictionary dics, string filepath)
        {

            // 拡張子をBackUpへ変更する
            string BackUpiniFilePath = Path.ChangeExtension(filepath, "BackUp");
            // 出力セクション・キー検索
            List<string> Export_SECsList = new List<string>();
            List<string> Export_KEYsList = new List<string>();

            // バックアップが存在する場合、事前に削除
            if (File.Exists(BackUpiniFilePath) == true)
                File.Delete(BackUpiniFilePath);

            // ファイルがない場合、空ファイルを生成する
            EmptyFileCreate(filepath);

            // iniFileをバックアップする
            File.Move(filepath, BackUpiniFilePath);

            try
            {

                // 新しいiniFileデータ蓄積用
                StringBuilder SB = new StringBuilder();
                StringBuilder SB_comment = new StringBuilder();

                // 現在のiniFileを比較用で読みながら新しいiniを書き出す##########################################################################
                using (System.IO.StreamReader sr = new System.IO.StreamReader(BackUpiniFilePath, CharacterCode))
                {
                    string SectionName = "";
                    //bool SectionDelFlg = false;

                    while (sr.Peek() >= 0)
                    {

                        // Dim outStr As String = ""
                        string c1 = ""; // 先頭一文字
                        string c2 = ""; // 後方一文字
                        string s = Strings.Trim(sr.ReadLine()); // 左右の空白を削除

                        if (s != "")
                        {
                            c1 = Strings.Left(s, 1);
                            c2 = Strings.Right(s, 1);

                            if (c1 == "#" || c1 == ";" || c1 == "'")

                                // コメント行
                                SB_comment.AppendLine(s);
                            else if (c1 == "[" && c2 == "]")
                            {

                                // 一つ前のセクションに存在しないキーを探して出力する##########################
                                // 現在のdicsにセクションが存在する場合
                                if (Exist_Section(dics, SectionName) == true)
                                {

                                    // 指定セクション内のKEY一覧を取得
                                    ICollection keyCollection = ((OrderedDictionary)dics[SectionName]).Keys;

                                    foreach (string KEYName in keyCollection)
                                    {
                                        if (Export_KEYsList.Contains(KEYName) == false)
                                        {
                                            string ADDKEY = KEYName;
                                            string ADDVAL = (String)((OrderedDictionary)dics[SectionName])[KEYName];
       
                                            SB.AppendLine(ADDKEY + "=" + ADDVAL);
                                        }
                                    }
                                }
                                else
                                {
                                }
                                // #######################################################################

                                // コメント行を追加
                                SB.Append(SB_comment.ToString());
                                // コメントを初期化
                                SB_comment.Clear();
                                // セクション行
                                SectionName = Strings.Trim(Strings.Mid(s, 2, s.Length - 2)); // 左右の[]を外す

                                if (dics.Contains(SectionName) == true)
                                {
                                    SB.AppendLine(s);
                                    // セクションリストを作成
                                    Export_SECsList.Add(SectionName);
                                    // キーリストを初期化
                                    Export_KEYsList.Clear();
                                }
                                else
                                {
                                }
                            }
                            else if (dics.Contains(SectionName) == true)
                            {

                                // キー行
                                string[] key = Strings.Split(s, "=", -1, Constants.vbBinaryCompare);
                                key[0] = Strings.Trim(key[0]);
                                key[1] = Strings.Trim(key[1]);

                                OrderedDictionary dicKeys = new OrderedDictionary();
                                dicKeys = (OrderedDictionary)dics[SectionName];

                                if (dicKeys.Contains(key[0]) == true)
                                {

                                    // キーに対するVALを書き換えて出力
                                    string ADDKEY = key[0];
                                    string ADDVAL = (String)dicKeys[key[0]]; // 最新の情報に更新
                                    //string ADDVAL = dicKeys[key[0]]; // 最新の情報に更新

                                    // コメント行を追加
                                    SB.Append(SB_comment.ToString());
                                    // コメントを初期化
                                    SB_comment.Clear();

                                    // キー行を追加
                                    SB.AppendLine(ADDKEY + "=" + ADDVAL);
                                    Export_KEYsList.Add(ADDKEY);
                                }
                                else
                                {
                                }
                            }
                            else
                            {
                            }
                        }
                        else
                            // NULL行
                            SB_comment.AppendLine(s);
                    }

                    // 最終section用の処理
                    // 一つ前のセクションに存在しないキーを探して出力する##########################
                    // 現在のdicsにセクションが存在する場合
                    if (Exist_Section(dics, SectionName) == true)
                    {

                        // 指定セクション内のKEY一覧を取得
                        ICollection keyCollection = ((OrderedDictionary)dics[SectionName]).Keys;

                        foreach (string KEYName in keyCollection)
                        {
                            if (Export_KEYsList.Contains(KEYName) == false)
                            {
                                string ADDKEY = KEYName;
                                string ADDVAL = (String)((OrderedDictionary)dics[SectionName])[KEYName];
                                SB.AppendLine(ADDKEY + "=" + ADDVAL);
                            }
                        }
                    }
                    else
                    {
                    }
                    // #######################################################################

                    // セクションが存在しない場合、新しくセクションを最終行へ書き出す
                    ICollection SECCollection = dics.Keys;

                    foreach (string SECName in SECCollection)
                    {
                        if (Export_SECsList.Contains(SECName) == false)
                        {
                            SB.AppendLine("");
                            SB.AppendLine("[" + SECName + "]");
                            ICollection KEYCollection = ((OrderedDictionary)dics[SECName]).Keys;

                            foreach (string KEYName in KEYCollection)
                            {
                                string ADDKEY = KEYName;
                                string ADDVAL = (String)((OrderedDictionary)dics[SECName])[KEYName];
                                SB.AppendLine(ADDKEY + "=" + ADDVAL);
                            }
                        }
                    }

                    File.AppendAllText(filepath, SB.ToString(), CharacterCode);
                }
                // 現在のiniFileを比較用で読みながら新しいiniを書き出す##########################################################################

                // バックアップしたiniFileを削除
                if (File.Exists(BackUpiniFilePath) == true)
                    File.Delete(BackUpiniFilePath);
            }
            catch (Exception ex)
            {
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }
        }
        private void EmptyFileCreate(string filepath)
        {
            System.IO.FileStream hStream = null;

            if (File.Exists(filepath) == false)
            {
                try
                {
                    // 指定したパスのファイルを作成する
                    hStream = System.IO.File.Create(filepath);
                }
                catch (Exception ex)
                {
                    string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
                }

                finally
                {

                    // 作成時に返される FileStream を利用して閉じる
                    if (hStream != null)
                        hStream.Close();
                    }
            }
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryから指定したセクションのインデックスを取得する★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <returns></returns>
        public int GetSectionIndex(OrderedDictionary dics, string TargetSectionName)
        {
            int ret = -1;
            int i = -1;
            try
            {
                ICollection keyCollection = dics.Keys;

                foreach (string SECName in keyCollection)
                {
                    i += 1;
                    if (SECName == TargetSectionName)
                    {
                        ret = i;
                        break;
                    }
                    else
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                ret = -1;
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }

            return ret;
        }
        /// <summary>
        ///     ''' ★指定したOrderedDictionaryから指定したキーのインデックスを取得する★
        ///     ''' </summary>
        ///     ''' <param name="dics"></param>
        ///     ''' <param name="TargetSectionName"></param>
        ///     ''' <param name="TargetKeyName"></param>
        ///     ''' <returns></returns>
        public int GetKeyIndex(OrderedDictionary dics, string TargetSectionName, string TargetKeyName)
        {
            int ret = -1;
            int i = -1;
            try
            {
                ICollection keyCollection = ((OrderedDictionary)dics[TargetSectionName]).Keys;

                foreach (string KeyName in keyCollection)
                {
                    i += 1;
                    if (KeyName == TargetKeyName)
                    {
                        ret = i;
                        break;
                    }
                    else
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                ret = -1;
                string ErrorLog = ErrorLogCreate(ClassName, MethodBase.GetCurrentMethod().Name, ex.Message, ex.HResult);
            }

            finally
            {
            }

            return ret;
        }
        public String ErrorLogCreate(string CLSName, string MethodName, string ErrorMsg, int ResultCode, bool MsgFlg = true)
        {
            // 名前空間まで含めてクラス名を取得
            string ClassName = "ExF_General";
            string ret = "";
            List<string> LS = new List<string>();

            try
            {
                if (CLSName != "")
                    LS.Add("ClassName:" + CLSName);
                LS.Add("MethodName:" + MethodName);
                LS.Add("ResultCode:" + ResultCode.ToString());
                LS.Add("ErrorMsg:" + ErrorMsg);
                ret = Strings.Join(LS.ToArray(), Constants.vbCrLf);
            }
            catch (Exception ex)
            {
                ret = "ClassName:" + ClassName + Constants.vbCrLf + "MethodName:" + MethodBase.GetCurrentMethod().Name + Constants.vbCrLf + "ResultCode:" + ex.HResult + Constants.vbCrLf + "ErrorMsg:" + ex.Message;
            }

            finally
            {
                //if (MsgFlg == true)
                    //MessageBox.Show(ret, "ExceptionMessage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return ret;
        }


    }

}

