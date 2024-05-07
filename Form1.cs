using CefSharp;
using CefSharp.WinForms;
using HtmlAgilityPack;
using MySql.Data.MySqlClient;
using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using File = System.IO.File;

namespace ChromeBrowserMacro
{
    public partial class Form1 : Form
    {
        #region 전역
        /// <summary>
        /// Common 클래스 생성
        /// </summary>
        protected Common _oCommon = new Common();

        /// <summary>
        /// 크롬 브라우저 객체
        /// </summary>
        public ChromiumWebBrowser _oChromeBrowser;
        #endregion

        public Form1()
        {
            InitializeComponent();
            InitChromeBrowser();
        }

        #region { InitChromeBrowser : 크롬 브라우저 초기 세팅 }
        /// <summary>
        /// 크롬 브라우저 초기 세팅
        /// </summary>
        protected void InitChromeBrowser()
        {
            CefSettings oCefSettings = new CefSettings();
            Cef.Initialize(oCefSettings);
            _oChromeBrowser = new ChromiumWebBrowser("");
            this.chromiumWebBrowser1.Controls.Add(_oChromeBrowser);
            _oChromeBrowser.Dock = DockStyle.Fill;

            //Alert 컨트롤
            JsDialogHandler oJsDialogHandler = new JsDialogHandler();
            _oChromeBrowser.JsDialogHandler = oJsDialogHandler;

            //페이지 로딩 완료 이벤트
            _oChromeBrowser.FrameLoadEnd += ChromeBrowserFrameLoadEnd;
        }
        #endregion

        #region { btnStart_Click : 시작 버튼 클릭 }
        /// <summary>
        /// 시작 버튼 클릭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnStart_Click(object sender, EventArgs e)
        {
            LoadWebBrowser();
        }
        #endregion

        #region { btnStop_Click : 중지 버튼 클릭 }
        /// <summary>
        /// 중지 버튼 클릭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnStop_Click(object sender, EventArgs e)
        {

        }
        #endregion

        #region { LoadWebBrowser : 크롬 브라우저 바인딩 End }
        /// <summary>
        /// 브라우저 객체에 URL Load
        /// </summary>
        protected void LoadWebBrowser()
        {
            _oChromeBrowser.Load(txtURLSample.Text);
        }
        #endregion

        #region { ChromeBrowserFrameLoadEnd : 크롬 브라우저 바인딩 End }
        /// <summary>
        /// 크롬 브라우저 바인딩 End
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ChromeBrowserFrameLoadEnd(object sender, FrameLoadEndEventArgs e)
        {
            try
            {
                if (e.Frame.IsMain)
                {
                    _oChromeBrowser.GetSourceAsync().ContinueWith(taskHtml =>
                    {
                        var html = taskHtml.Result;
                        var dom = new HtmlAgilityPack.HtmlDocument();
                        dom.LoadHtml(html);

                        if (dom.GetElementbyId("div_tab_1") != null && !string.IsNullOrWhiteSpace(dom.GetElementbyId("div_tab_1").OuterHtml))
                        {

                        }
                        else if (dom.GetElementbyId("id") != null && !string.IsNullOrWhiteSpace(dom.GetElementbyId("id").OuterHtml))
                        {
                            _oChromeBrowser.GetMainFrame().ExecuteJavaScriptAsync("document.getElementById(\"id\").value = \"sfdevil67\";");
                            _oChromeBrowser.GetMainFrame().ExecuteJavaScriptAsync("document.getElementById(\"pw\").value = \"Nn811226@@\";");
                            _oChromeBrowser.GetMainFrame().ExecuteJavaScriptAsync("document.getElementById(\"log.login\").click();");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                Common.WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
            }
        }
        #endregion

        #region { GetHtmlData : Html 불필요한 데이터 제거 }
        /// <summary>
        /// Html 불필요한 데이터 제거
        /// </summary>
        /// <param name="pHtml"></param>
        protected string GetHtmlData(HtmlAgilityPack.HtmlDocument pDom)
        {
            //불필요한 html 제거
            if (pDom.GetElementbyId("lbl_detail_view") != null) pDom.GetElementbyId("lbl_detail_view").ParentNode.Remove();
            if (pDom.GetElementbyId("btn_detail") != null) pDom.GetElementbyId("btn_detail").Remove();
            if (pDom.GetElementbyId("btn_transfer_menu_alter") != null) pDom.GetElementbyId("btn_transfer_menu_alter").Remove();
            if (pDom.GetElementbyId("div_img_preview") != null) pDom.GetElementbyId("div_img_preview").Remove();

            if (pDom.GetElementbyId("appline_summary") != null) pDom.GetElementbyId("appline_summary").SetAttributeValue("style", "display: none");
            if (pDom.GetElementbyId("app_line_display") != null) pDom.GetElementbyId("app_line_display").SetAttributeValue("style", "display: ''");
            
            foreach(HtmlNode node in pDom.DocumentNode.SelectNodes("//table[@id='tbl_doc_info']"))
            {
                if (node.ParentNode.GetAttributeValue("class", string.Empty).Contains("ui-effects-wrapper"))
                {
                    node.ParentNode.SetAttributeValue("style", "");
                    node.ParentNode.RemoveClass();
                    break;
                }
            }

            string formName = pDom.GetElementbyId("lbl_form_nm").ParentNode.OuterHtml;
            string html = pDom.GetElementbyId("div_tab_1").OuterHtml;

            //이미지 src 변경. (도메인 주소 붙여주기)
            Regex.Replace(pDom.GetElementbyId("div_tab_1").OuterHtml, @"<img.*?>", new MatchEvaluator(delegate (Match m1)
            {
                string src = Regex.Match(m1.Value, "src=[\"'](.+?)[\"']", RegexOptions.IgnoreCase).Groups[1].Value;
                if (src.Contains("Common/ViewImage") && !src.Contains("customer_id=25"))
                {
                    //세션없는 인라인 이미지에 customer_id=25 붙여주면 이미지 보임.
                    html = html.Replace("\"" + src + "\"", "\"" + src + "&customer_id=25" + "\"");
                }
                html = html.Replace("\"" + src, "\"" + Common.Domain + src);
                return "";
            }));

            //스크립트 제거
            html = Common.RemoveTag(html, @"<script.*?</script>");

            // 이벤트 제거
            html = Common.RemoveTag(html, @"(onclick|onmouseover|onmouseout|onload|onfocus|onmousedown|onkeydown|onkeypress)\s?=\s?(['""])[\s\S]*?\2");

            //주석제거
            html = Common.RemoveTag(html, @"<!--.*?-->");

            html = Common.HtmlHeader + formName + html + Common.HtmlFooter;
            return html;
        }
        #endregion
    }

    #region { JsDialogHandler : Alert 컨트롤 }
    /// <summary>
    /// Alert 컨트롤
    /// </summary>
    public class JsDialogHandler : IJsDialogHandler
    {
        /// <summary>
        /// 변환 도중 window alert 창 떳을때 로그 저장 후 진행시킨다.
        /// </summary>
        /// <param name="browserControl"></param>
        /// <param name="browser"></param>
        /// <param name="originUrl"></param>
        /// <param name="dialogType"></param>
        /// <param name="messageText"></param>
        /// <param name="defaultPromptText"></param>
        /// <param name="callback"></param>
        /// <param name="suppressMessage"></param>
        /// <returns></returns>
        public bool OnJSDialog(IWebBrowser browserControl, IBrowser browser, string originUrl, CefJsDialogType dialogType, string messageText, string defaultPromptText, IJsDialogCallback callback, ref bool suppressMessage)
        {
            callback.Continue(true);
            return true;
        }

        public bool OnJSBeforeUnload(IWebBrowser browserControl, IBrowser browser, string message, bool isReload, IJsDialogCallback callback)
        {
            return true;
        }

        public void OnResetDialogState(IWebBrowser browserControl, IBrowser browser)
        {

        }

        public void OnDialogClosed(IWebBrowser browserControl, IBrowser browser)
        {

        }

        public bool OnBeforeUnloadDialog(IWebBrowser chromiumWebBrowser, IBrowser browser, string messageText, bool isReload, IJsDialogCallback callback)
        {
            return true;
        }
    }
    #endregion
}