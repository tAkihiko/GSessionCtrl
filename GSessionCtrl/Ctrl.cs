using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace GSessionCtrl
{
    /// <summary>
    /// GSessionを操作するクラスを提供します。
    /// </summary>
    public class Ctrl
    {
        /// <summary>
        /// エンコーダ
        /// </summary>
        static Encoding m_encoder = Encoding.GetEncoding("UTF-8");

        /// <summary>
        /// UserAgent
        /// </summary>
        static string m_useragent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.116 Safari/537.36";

        /// <summary>
        /// GSession ユーザID
        /// </summary>
        static string m_id = "";

        /// <summary>
        /// GSession パスワード
        /// </summary>
        static string m_passwd = "";

        /// <summary>
        /// GSession上のユーザ管理ID
        /// </summary>
        static int m_sid = -1;

        /// <summary>
        /// スケジュールノード
        /// </summary>
        public class ScheduleNode
        {
            /// <summary>
            /// スケジュール対象名
            /// </summary>
            public string Name { get; private set; }

            /// <summary>
            /// スケジュール開始日時
            /// </summary>
            public DateTime Begin { get; private set; }

            /// <summary>
            /// スケジュール終了日時
            /// </summary>
            public DateTime End { get; private set; }

            /// <summary>
            /// スケジュールタイトル
            /// </summary>
            public string Title { get; private set; }

            /// <summary>
            /// スケジュール本文
            /// </summary>
            public string Text { get; private set; }

            /// <summary>
            /// コンストラクタ
            /// </summary>
            /// <param name="name">スケジュール対象名</param>
            /// <param name="begin">スケジュール開始日時</param>
            /// <param name="end">スケジュール終了日時</param>
            /// <param name="title">スケジュールタイトル</param>
            /// <param name="text">スケジュール詳細</param>
            public ScheduleNode(string name, DateTime begin, DateTime end, string title, string text)
            {
                Name = name;
                Begin = begin;
                End = end;
                Title = title;
                Text = text;
            }
        }

        /// <summary>
        /// GETリクエスト送信関数
        /// </summary>
        /// <param name="url">送信先URL</param>
        /// <param name="cc">Cookieコンテナ</param>
        /// <returns>取得テキスト</returns>
        private static string _HttpGet(string url, CookieContainer cc)
        {

            // リクエストの作成
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.CookieContainer = cc;
            req.UserAgent = m_useragent;

            WebResponse res = req.GetResponse();

            // レスポンスの読み取り
            Stream resStream = res.GetResponseStream();
            StreamReader sr = new StreamReader(resStream, m_encoder);
            string result = sr.ReadToEnd();
            sr.Close();
            resStream.Close();

            return result;
        }

        /// <summary>
        /// POSTリクエスト送信関数
        /// </summary>
        /// <param name="url">送信先URL</param>
        /// <param name="vals">送信パラメータ</param>
        /// <param name="cc">Cookieコンテナ</param>
        /// <returns>取得テキスト</returns>
        private static string _HttpPost(string url, Hashtable vals, CookieContainer cc)
        {
            string param = "";
            foreach (string k in vals.Keys)
            {
                param += String.Format("{0}={1}&", k, vals[k]);
            }
            byte[] data = Encoding.ASCII.GetBytes(param);

            // リクエストの作成
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = "application/x-www-form-urlencoded";
            req.ContentLength = data.Length;
            req.CookieContainer = cc;
            req.UserAgent = m_useragent;

            // ポスト・データの書き込み
            Stream reqStream = req.GetRequestStream();
            reqStream.Write(data, 0, data.Length);
            reqStream.Close();

            WebResponse res = req.GetResponse();

            // レスポンスの読み取り
            Stream resStream = res.GetResponseStream();
            StreamReader sr = new StreamReader(resStream, m_encoder);
            string result = sr.ReadToEnd();
            sr.Close();
            resStream.Close();

            return result;
        }

        /// <summary>
        /// ログイン関数
        /// </summary>
        /// <param name="id">ユーザID</param>
        /// <param name="password">パスワード</param>
        /// <param name="cc">Cookieコンテナ</param>
        /// <returns>ログイン成否</returns>
        private static bool _Login(string id, string password, CookieContainer cc)
        {
            string login = "http://172.16.0.5:8080/gsession/common/cmn001.do";
            string html = "";

            // ログイン・ページへのアクセス
            Hashtable vals = new Hashtable();

            // ログイン情報取得
            html = _HttpGet(login, cc);

            // トランザクショントークン取得
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            HtmlAgilityPack.HtmlNodeCollection inputNodes = doc.DocumentNode.SelectNodes("//input");
            string name = "";
            string token = "";
            if (inputNodes == null)
            {
                return false;
            }
            foreach (HtmlAgilityPack.HtmlNode node in inputNodes)
            {
                name = node.GetAttributeValue("name", "");
                if (name == "org.apache.struts.taglib.html.TOKEN")
                {
                    token = node.GetAttributeValue("value", "");
                }
            }

            // パラメータ設定
            vals["CMD"] = "login";
            vals["org.apache.struts.taglib.html.TOKEN"] = token;
            vals["cmn001loginType"] = 1;
            vals["url"] = "";
            vals["cmn001Userid"] = id;
            vals["cmn001Passwd"] = password;

            // ログイン成否取得
            html = _HttpPost(login, vals, cc);

            doc.LoadHtml(html);
            HtmlAgilityPack.HtmlNodeCollection divNodes = doc.DocumentNode.SelectNodes("//div");
            string cls = "";
            bool errflg = false;
            if (divNodes != null)
            {
                foreach (HtmlAgilityPack.HtmlNode node in divNodes)
                {
                    cls = node.GetAttributeValue("class", "");
                    if (cls == "text_error")
                    {
                        // text_errorクラスのdivがあればログインに失敗していると仮定
                        // 長時間ログインせず系のエラーは未対策
                        errflg = true;
                        break;
                    }
                }
            }

            if (errflg)
            {
                return false;
            }

            // GSessionのユーザID取得
            string schmain = "http://172.16.0.5:8080/gsession/schedule/schmain.do";
            html = _HttpGet(schmain, cc);
            doc.LoadHtml(html);
            inputNodes = doc.DocumentNode.SelectNodes("//input");
            m_sid = -1;
            foreach (HtmlAgilityPack.HtmlNode node in inputNodes)
            {
                if (node.GetAttributeValue("name","") == "schSelectUsrSid")
                {
                    m_sid = int.Parse(node.GetAttributeValue("value", ""));
                }
            }

            if (m_sid < 0)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// ログアウト関数
        /// </summary>
        /// <param name="cc">Cookieコンテナ</param>
        /// <returns>ログアウト成否</returns>
        private static bool _Logout(CookieContainer cc)
        {
            string logout = "http://172.16.0.5:8080/gsession/common/cmn001.do?CMD=logout";
            _HttpGet(logout, cc);
            return true;
        }

        /// <summary>
        /// 在席状態変更
        /// </summary>
        /// <param name="status">在席状態 1:在席 2:不在</param>
        /// <param name="cc">Cookieコンテナ</param>
        /// <returns>変更成否</returns>
        private static bool _Zaiseki(int status, string message, CookieContainer cc)
        {
            string zaiseki = "http://172.16.0.5:8080/gsession/zaiseki/zskmain.do";

            if (status != 1)
            {
                status = 2;
            }

            Hashtable vals = new Hashtable();
            vals["CMD"] = "zskEdit";
            vals["zskUioStatus"] = status;
            vals["zskUioBiko"] = message;
            vals["mainReDspFlg"] = 0;

            _HttpPost(zaiseki, vals, cc);

            return true;
        } 
        
        /// <summary>
        /// スケジュール取得
        /// </summary>
        /// <param name="usrid">GSession上の管理ID</param>
        /// <param name="cc">Cookieコンテナ</param>
        /// <returns>スケジュールのリスト</returns>
        private static List<ScheduleNode> _Sch(int usrid, CookieContainer cc)
        {
            string sch = "http://172.16.0.5:8080/gsession/schedule/sch010.do";

            Hashtable vals = new Hashtable();
            DateTime today = DateTime.Today;
            today.ToLocalTime();
            string date = today.ToString("yyyyMMdd");

            vals["CMD"] = "list";
            vals["dspMod"] = 1;
            vals["sch010SelectUsrSid"] = usrid;
            vals["sch010SelectUsrKbn"] = 0;
            vals["cmd"] = "list";
            vals["sch010DspDate"] = date;
            vals["changeDateFlg"] = 0;
            vals["sch010SelectDate"] = "";
            vals["sch010SchSid"] = "";
            vals["sch010DspGpSid"] = 6; // 全社員グループ
            vals["sch010searchWord"] = "";

            // スケジュールデータ取得
            string html = _HttpPost(sch, vals, cc);

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            HtmlAgilityPack.HtmlNodeCollection inputNodes = doc.DocumentNode.SelectNodes("//tr");
            if (inputNodes == null)
            {
                return null;
            }

            List<ScheduleNode> schlist = new List<ScheduleNode>();
            foreach (HtmlAgilityPack.HtmlNode node in inputNodes)
            {
                HtmlAgilityPack.HtmlNode name_n = node.SelectSingleNode(".//td[1]");
                HtmlAgilityPack.HtmlNode b_day_n = node.SelectSingleNode(".//td[2]");
                HtmlAgilityPack.HtmlNode e_day_n = node.SelectSingleNode(".//td[3]");

                if (name_n == null || b_day_n == null || e_day_n == null)
                {
                    continue;
                }

                if (name_n.GetAttributeValue("class","") != "td_type1" && name_n.GetAttributeValue("class","") != "td_type29")
                {
                    continue;
                }

                if (b_day_n.GetAttributeValue("class", "") != "td_type1" && b_day_n.GetAttributeValue("class", "") != "td_type29")
                {
                    continue;
                }

                if (e_day_n.GetAttributeValue("class", "") != "td_type1" && e_day_n.GetAttributeValue("class", "") != "td_type29")
                {
                    continue;
                }

                HtmlAgilityPack.HtmlNode aNode = node.SelectSingleNode(".//a");

                string name = name_n.InnerText;
                string[] b_day = b_day_n.InnerText.Split('　');
                string[] e_day = e_day_n.InnerText.Split('　');
                string title = aNode.InnerText;
                string[] b_day_date = b_day[0].Split('/');
                string[] b_day_time = b_day[1].Split(':');
                string[] e_day_date = e_day[0].Split('/');
                string[] e_day_time = e_day[1].Split(':');

                title = Regex.Replace(title, "[\r\n]", "");
                title = Regex.Replace(title, "<!--.*?-->", "");
                title = Regex.Replace(title, "^[ 　]+", "");
                title = Regex.Replace(title, "[ 　]+$", "");

                DateTimeKind kind = DateTimeKind.Local;
                DateTime begin = new DateTime(int.Parse(b_day_date[0]), int.Parse(b_day_date[1]), int.Parse(b_day_date[2]), int.Parse(b_day_time[0]), int.Parse(b_day_time[1]), 0, kind); ;
                DateTime end = new DateTime(int.Parse(e_day_date[0]), int.Parse(e_day_date[1]), int.Parse(e_day_date[2]), int.Parse(e_day_time[0]), int.Parse(e_day_time[1]), 0, kind);

                schlist.Add(new ScheduleNode(name, begin, end, title, ""));
            }

            return schlist;
        }

        /// <summary>
        /// GSessionパラメータ設定
        /// </summary>
        /// <param name="id">ログインID</param>
        /// <param name="passwd">ログインパスワード</param>
        public static void ParamSetting(string id, string passwd)
        {
            m_id = id;
            m_passwd = passwd;
        }

        /// <summary>
        /// 在席化
        /// </summary>
        /// <returns>成否</returns>
        public static bool Zaiseki(string message = "")
        {
            CookieContainer cc = new CookieContainer();
            bool login = false;
            bool logout = false;
            bool zaiseki = false;

            // ログイン
            login = _Login(m_id, m_passwd, cc);
            if (!login)
            {
                return false;
            }

            // 在席化
            zaiseki = _Zaiseki(1, message, cc);

            // ログアウト
            logout = _Logout(cc);

            return (zaiseki && logout);
        }

        /// <summary>
        /// 不在化
        /// </summary>
        /// <returns>成否</returns>
        public static bool Huzai(string message = "")
        {
            CookieContainer cc = new CookieContainer();
            bool login = false;
            bool logout = false;
            bool zaiseki = false;

            // ログイン
            login = _Login(m_id, m_passwd, cc);
            if (!login)
            {
                return false;
            }

            // 不在化
            zaiseki = _Zaiseki(2, message, cc);

            // ログアウト
            logout = _Logout(cc);

            return (zaiseki && logout);
        }

        /// <summary>
        /// スケジュール取得
        /// </summary>
        /// <returns>スケジュールのリスト</returns>
        public static List<ScheduleNode> Schedule()
        {
            CookieContainer cc = new CookieContainer();
            List<ScheduleNode> schlist = null;

            // ログイン
            if (!_Login(m_id, m_passwd, cc))
            {
                return null;
            }

            // スケジュール取得
            schlist = _Sch(m_sid, cc);

            // ログアウト
            _Logout(cc);

            return schlist;
        }
    }
}
