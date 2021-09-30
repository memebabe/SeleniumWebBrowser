using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SeleniumWebBrowser
{
    public class WebBrowser : IDisposable
    {
        public static int DefaultWidth = 1200;
        public static int DefaultHeight = 720;
        public static string CustomChromePath = null;

        private static List<WebBrowser> _runningWebBrowsers = new List<WebBrowser>();

        public static void DisposeAll()
        {
            List<Task> runningTasks = new List<Task>();
            foreach (var browser in _runningWebBrowsers.ToArray())
            {
                if (browser.IsDisposed == false)
                {
                    runningTasks.Add(Task.Run(() =>
                    {
                        try { browser.Dispose(); } catch { }
                    }));
                }
            }

            Task.WaitAll(runningTasks.ToArray());
        }

        //=====================================================================

        public int ProcessId { get; set; }
        private IWebDriver _webDriver;

        public IntPtr Hwnd { get; private set; }

        public string CurrentUrl
        {
            get
            {
                return _webDriver?.Url;
            }
        }

        public string PageSource
        {
            get { return _webDriver.PageSource; }
        }

        public bool IsDisposed { get; private set; } = false;

        //======================================================================
        public WebBrowser(bool isHidden = false, bool isLoadImages = true, ProxyObject proxyObject = null, string userAgent = null, params string[] args)
        {
            var service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            var options = Createptions(isHidden, isLoadImages, proxyObject, userAgent, args);

            _webDriver = new ChromeDriver(service, options);
            this.ProcessId = service.ProcessId;

            if (string.IsNullOrEmpty(proxyObject.User) == false)
            {
                this.GoToUrl("chrome-extension://ggmdpepbjljkkkdaklfihhngmmgmpggp/options.html");

                this.FillText(By.Id("login"),proxyObject.User);
                this.FillText(By.Id("password"),proxyObject.Password);
                this.FillText(By.Id("retry"),"2");

                this.ClickElement(By.Id("save"));
                Sleep(300);
                CloseTab();
                SwitchToTabIndex(0);
            }

            _runningWebBrowsers.Add(this);
        }

        private ChromeOptions Createptions(bool isHidden, bool isLoadImages, ProxyObject proxyObject, string userAgent = null, params string[] args)
        {
            ChromeOptions options = new ChromeOptions();
            if (string.IsNullOrEmpty(CustomChromePath) == false)
                options.BinaryLocation = CustomChromePath;

            options.AddArgument("--disable-infobars");
            options.AddArgument("--disable-notifications");
            options.AddArgument("ignore-certificate-errors");
            options.AddAdditionalCapability("useAutomationExtension", false);
            options.AddArgument($"--window-size={DefaultWidth},{DefaultHeight}");

            if (string.IsNullOrEmpty(userAgent) == false)
                options.AddArgument($"user-agent={userAgent}");

            foreach (var arg in args)
            {
                options.AddArgument(arg);
                if (arg.Contains("profile-directory"))
                {
                    options.AddArgument("--flag-switches-begin");
                    options.AddArgument("--flag-switches-end");
                    options.AddArgument("--enable-audio-service-sandbox");
                }
            }
                
            if (isHidden)
            {
                options.AddArgument("--headless");
                options.AddArgument("no-sandbox");
            }

            if (isLoadImages == false)
                options.AddUserProfilePreference("profile.default_content_setting_values.images", 2);

            if (proxyObject != null  && string.IsNullOrEmpty(proxyObject.User) == false)
            {
                options.AddExtension("Extensions\\Proxy-Auto-Auth_2.0.crx");
                var proxy = new Proxy();
                proxy.Kind = ProxyKind.Manual;
                proxy.IsAutoDetect = false;
                proxy.HttpProxy =
                proxy.SslProxy = $"{proxyObject.Host}:{proxyObject.Port}";
                options.Proxy = proxy;
            }

            return options;
        }

        //==================================================================

        public void GoToUrl(string url)
        {
            try
            {
                _webDriver.Navigate().GoToUrl(url);
            }
            catch { CancelPageLoad(); }
        }

        public void OpenLinkInNewTab(string url)
        {
            this.NewTab(url);
        }


        public void Refresh()
        {
            _webDriver.Navigate().Refresh();
        }
 
        public void CancelPageLoad()
        {
            try
            {
                _webDriver.FindElement(By.TagName("body")).SendKeys(Keys.Escape);
            }
            catch { }
        }

        private bool _stopDelayCancelPageLoad;
        private Stopwatch _watchOfDelayCancelPageLoad = new Stopwatch();
        private long _delayCancelPageLoadTime;
        public void DelayToCancelPageLoad(long delayTime)
        {
            _delayCancelPageLoadTime = delayTime;
            _watchOfDelayCancelPageLoad.Restart();

            if (_stopDelayCancelPageLoad == false)
            {
                _stopDelayCancelPageLoad = false;

                Task.Run(() =>
                {
                    while (_watchOfDelayCancelPageLoad.ElapsedMilliseconds < _delayCancelPageLoadTime)
                    {
                        System.Threading.Thread.Sleep(1);
                        if (_stopDelayCancelPageLoad)
                        {
                            _watchOfDelayCancelPageLoad.Stop();
                            return;
                        }
                    }
                    CancelPageLoad();
                });
            }

        }

        public void StopDelayingToCancelPageLoad()
        {
            _stopDelayCancelPageLoad = true;
        }


        public void NewTab(string url = null)
        {
            this.ExecuteScript($"window.open('{url}');");
        }

        public void CloseTab()
        {
            _webDriver.Close();
        }

        public void SwitchToTabIndex(int index)
        {
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[index]);
        }

        public void SwitchToLastTab()
        {
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());
        }

        public void SwitchToFrame(IWebElement frame)
        {
            _webDriver.SwitchTo().Frame(frame);
        }

        public bool AcceptAlert()
        {
            try
            {
                var alert = _webDriver.SwitchTo().Alert();
                if (alert != null)
                {
                    alert.Accept();
                    return true;
                }
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }


        public IWebElement SwitchToActiveElement()
        {
            return this._webDriver.SwitchTo().ActiveElement();
        }

        public IWebElement WaitUntil_Element_Exists(By by, int timeout = 30000)
        {
            try
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromMilliseconds(timeout));
                return wait.Until(
                    (d) => {
                        try
                        {
                            return d.FindElement(by);
                        }
                        catch
                        {
                            return null;
                        }
                    });
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return null;
            }
        }

        public IWebElement WaitUntil_OneOfElements_Exists(IEnumerable<By> bys, int timeout = 30000)
        {
            try
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromMilliseconds(timeout));
                return wait.Until(
                    (d) => {
                        try
                        {
                            IWebElement element = null;
                            foreach (var by in bys)
                            {
                                try
                                {
                                    element = d.FindElement(by);
                                    if (element != null)
                                        break;
                                }
                                catch { }
                            }
                            return element;
                        }
                        catch
                        {
                            return null;
                        }
                    });
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return null;
            }
        }

        public bool WaitUntil_Element_NoExists(By by, int timeout = 30000)
        {
            try
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromMilliseconds(timeout));
                return wait.Until(
                    (d) => {
                        try
                        {
                            d.FindElement(by);
                            return false;
                        }
                        catch
                        {
                            return true;
                        }
                    });
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return false;
            }
        }

        public IWebElement WaitUntil_Element_Displayed(By by, int timeout = 30000)
        {
            try
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromMilliseconds(timeout));
                return wait.Until(
                    (d) => {
                        try
                        {
                            var elem = d.FindElement(by);
                            if (elem.Displayed)
                                return elem;
                            else
                                return null;
                        }
                        catch
                        {
                            return null;
                        }
                    });
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return null;
            }
        }

        public IWebElement WaitUntil_OneOfElements_Displayed(IEnumerable<By> bys, int timeout = 30000)
        {
            try
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromMilliseconds(timeout));
                return wait.Until(
                    (d) => {
                        try
                        {
                            IWebElement element = null;
                            foreach (var by in bys)
                            {
                                try
                                {
                                    element = d.FindElement(by);
                                    if (element != null && element.Displayed == true)
                                        break;
                                    else
                                        element = null;
                                }
                                catch { }
                            }
                            return element;
                        }
                        catch
                        {
                            return null;
                        }
                    });
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return null;
            }
        }

        public bool WaitUntil_Element_NoDisplayed(By by, int timeout = 30000)
        {
            try
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromMilliseconds(timeout));
                return wait.Until(
                    (d) => {
                        try
                        {
                            var elem = d.FindElement(by);
                            if (elem.Displayed)
                                return false;
                            else
                                return true;
                        }
                        catch
                        {
                            return true;
                        }
                    });
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return false;
            }
        }

        public IWebElement WaitUntil_Element_Enabled(By by, int timeout = 30000)
        {
            try
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromMilliseconds(timeout));
                return wait.Until(
                    (d) => {
                        try
                        {
                            var elem = d.FindElement(by);
                            if (elem.Enabled)
                                return elem;
                            else
                                return null;
                        }
                        catch
                        {
                            return null;
                        }
                    });
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return null;
            }
        }

        public bool WaitUntil_UrlChanged(int timeout = 30000)
        {
            string url = _webDriver.Url;
            var watch = System.Diagnostics.Stopwatch.StartNew();

            do
            {
                System.Threading.Thread.Sleep(100);
            }
            while (url.Equals(_webDriver.Url) &&
                watch.ElapsedMilliseconds < timeout);
            watch.Stop();

            return !url.Equals(_webDriver.Url);
        }


        public IWebElement FindElement(By by)
        {
            try
            {
                return _webDriver.FindElement(by);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public ReadOnlyCollection<IWebElement> FindElements(By by)
        {
            try
            {
                return _webDriver.FindElements(by);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public IWebElement FindOneOfElement(By[] bys)
        {
            try
            {
                IWebElement element = null;
                foreach (var by in bys)
                {
                    try
                    {
                        element = this.FindElement(by);
                        if (element != null && element.Displayed == true)
                            break;
                        else
                            element = null;
                    }
                    catch { }
                }
                return element;
            }
            catch
            {
                return null;
            }
        }


        public IWebElement GetActiveElement()
        {
            return this._webDriver.SwitchTo().ActiveElement();
        }

        public void ScrollToElement(IWebElement element, bool smooth = false)
        {
            ((IJavaScriptExecutor)_webDriver).ExecuteScript($"arguments[0].scrollIntoView({{behavior: \"{(smooth ? "smooth" : "auto")}\", block: \"start\"}});", element);
        }

        public void ScrollToEnd()
        {
            ((IJavaScriptExecutor)_webDriver).ExecuteScript($"window.scrollTo(0, document.body.scrollHeight);");
        }

        #region Fill datas
        public string GetText(By by)
        {
            var element = _webDriver.FindElement(by);
            return element.Text;
        }

        public bool FillText(By elementSelector, string text, bool stepBystep = true)
        {
            var webElement = FindElement(elementSelector);
            if (webElement != null)
            {
                return FillText(webElement, text, stepBystep);
            }
            else
            {
                return false;
            }
        }

        public bool FillText(IWebElement element, string text, bool stepBystep = true)
        {
            if (element != null)
            {
                element.Click();
                Sleep(200);
                element.Clear();
                Sleep(100);

                if (stepBystep)
                    SendKeysStepByStep(element, text);
                else
                    element.SendKeys(text);

                return true;
            }
            else
            {
                return false;
            }
        }


        public bool SetInnerHtml(By elementSelector, string text)
        {
            var webElement = FindElement(elementSelector);
            if (webElement != null)
            {
                return SetInnerHtml(webElement, text);
            }
            else
            {
                return false;
            }
        }

        public bool SetInnerHtml(IWebElement element, string text)
        {
            if (element != null)
            {
                ((IJavaScriptExecutor)_webDriver).ExecuteScript($"var ele=arguments[0]; ele.innerHTML = '{text}';", element);

                return true;
            }
            else
            {
                return false;
            }
        }


        public bool SetAttribute(By elementSelector, string attr, string value)
        {
            var webElement = FindElement(elementSelector);
            if (webElement != null)
            {
                return SetAttribute(webElement, attr, value);
            }
            else
            {
                return false;
            }
        }

        public bool SetAttribute(IWebElement element, string attr, string value)
        {
            if (element != null)
            {
                ((IJavaScriptExecutor)_webDriver).ExecuteScript($"var ele=arguments[0]; ele.setAttribute('{attr}', '{value}');", element);

                return true;
            }
            else
            {
                return false;
            }
        }


        public bool SelectOptionValue(By elementSelector, string value)
        {
            try
            {
                var webElement = FindElement(elementSelector);
                webElement.Click();
                Sleep();

                var selectElement = new SelectElement(webElement);
                selectElement.SelectByValue(value);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion Fill datas


        //=====================================================================
        public bool ClickElement(By elementSelector)
        {
            try
            {
                var webElement = FindElement(elementSelector);
                if (webElement == null)
                    return false;

                return ClickElement(webElement);
            }
            catch
            {
                return false;
            }
        }

        public bool ClickElement(IWebElement element)
        {
            try
            {
                ScrollToElement(element);
                element.Click();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool WaitToClickElement(By elementSelector, int timeout = 30000)
        {
            try
            {
                IWebElement webElement = this.WaitUntil_Element_Exists(elementSelector, timeout);
                if (webElement == null)
                    return false;

                Sleep();

                ClickElement(webElement);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void ClickElementUsingJs(IWebElement element)
        {
            ScrollToElement(element);
            ((IJavaScriptExecutor)_webDriver).ExecuteScript("arguments[0].click();", element);
        }

        public void ClickElementUsingJs(By elementSelector)
        {
            var webElement = FindElement(elementSelector);
            if (webElement == null)
                return;

            ClickElementUsingJs(webElement);
        }


        public bool SendMouseOver(By elementSelector)
        {
            try
            {
                var webElement = FindElement(elementSelector);
                if (webElement == null)
                    return false;

                SendMouseOver(webElement);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool SendMouseOver(IWebElement element)
        {
            try
            {
                Actions actions = new Actions(_webDriver);
                actions.MoveToElement(element);
                actions.Click().Build().Perform();
                return true;
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// Drag 'element' and drop it on 'destinationElememt'.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="destinationElememt"></param>
        /// <returns></returns>
        public bool DragAndDrop(IWebElement element, IWebElement destinationElememt)
        {
            try
            {
                //Creating object of Actions class to build composite actions
                Actions builder = new Actions(_webDriver);

                //Building a drag and drop action
                var dragAndDrop = builder.ClickAndHold(element)
                .MoveToElement(destinationElememt)
                .Release(destinationElememt)
                .Build();

                //Performing the drag and drop action
                dragAndDrop.Perform();
                return true;
            }
            catch
            {
                return false;
            }
        }
        //=====================================================================

        /// <summary>
        /// Fake human typing actions.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="keys"></param>
        public void SendKeysStepByStep(IWebElement element, string keys)
        {
            Random rand = new Random();

            foreach (var key in keys)
            {
                element.SendKeys($"{key}");
                Sleep(rand.Next(10, 80));
            }
        }
        
        public void SendKeys(string keys)
        {
            _webDriver.SwitchTo().ActiveElement().SendKeys(keys);
        }

        public void Patse()
        {
            _webDriver.SwitchTo().ActiveElement().SendKeys(Keys.Control + "v");
        }

        //=====================================================================
       
        private Random _sleepRand = new Random(20);
        public void Sleep(int time = -1)
        {
            if (time == -1)
                System.Threading.Thread.Sleep(_sleepRand.Next(321, 1234));
            else
                System.Threading.Thread.Sleep(time);
        }

        public void Sleep(int mintime, int maxtime)
        {
            System.Threading.Thread.Sleep(this._sleepRand.Next(mintime, maxtime));
        }

        //=====================================================================

        public object ExecuteScript(string script, params object[] inputs)
        {
            return ((IJavaScriptExecutor)_webDriver).ExecuteScript(script, inputs);
        }

        public object Get(string url, params object[] inputs)
        {
            string script = $@"let xmlhttp = new XMLHttpRequest();
            xmlhttp.open('GET', '{url}', false);
            xmlhttp.send();
            return xmlhttp.response;";

            return ExecuteScript(script, inputs);
        }

        public object Post(string url, params object[] inputs)
        {
            string script = $@"let xmlhttp = new XMLHttpRequest();
            xmlhttp.open('POST', '{url}', false);
            xmlhttp.send();
            return xmlhttp.responseText;";

            return ExecuteScript(script, inputs);
        }


        //=====================================================================

        public ReadOnlyCollection<Cookie> GetCookies()
        {
            return this._webDriver.Manage().Cookies.AllCookies;
        }

        public string GetCookiesString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (OpenQA.Selenium.Cookie c in this._webDriver.Manage().Cookies.AllCookies)
                sb.Append($"{c.Name}={c.Value};");

            return sb.ToString();
        }

        public void SetCookiesString(string cookies)
        {
            cookies = HttpUtility.UrlDecode(cookies);
            cookies = cookies.Replace(" ", "");

            foreach (string cookie in cookies.Split(';'))
            {
                if (string.IsNullOrEmpty(cookie))
                    continue;

                string[] values = cookie.Split('=');
                if (values.Length > 1
                    && !string.IsNullOrEmpty(values[0])
                    && !string.IsNullOrEmpty(values[1]))
                {
                    Cookie c = new Cookie(name: values[0], value: values[1]);
                    this._webDriver.Manage().Cookies.AddCookie(c);
                }
            }
        }

        public void ClearCookies()
        {
            _webDriver.Manage().Cookies.DeleteAllCookies();
        }

        //=====================================================================

        private void Log(string msg = "", [System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
#if DEBUG
            Debug.WriteLine($"{this.GetType().Name}\t{methodName}", msg);
#endif
        }

        public void Dispose()
        {
            try
            {
                this.IsDisposed = true;

                _runningWebBrowsers.Remove(this);

                _webDriver.Quit();
                _webDriver.Dispose();
            }
            catch { }
        }
    }
}
