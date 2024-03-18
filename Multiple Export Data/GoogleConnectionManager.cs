using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Threading;

namespace MyGoogleServices
{
    /// <summary>
    /// 구글 API 연결 매니저
    /// </summary>
    public class GoogleConnectionMannager
    {
        public ManagerTypes GetManagerTypes()
        {
            return ManagerTypes.Connection;
        }
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static readonly string[] Scopes = { GmailService.Scope.GmailReadonly, SheetsService.Scope.SpreadsheetsReadonly };

        private static GoogleConnectionMannager instance;
        /// <summary>
        /// 사용할 서비스 메니저들의 리스트.
        /// </summary>
        //private IGoogleManager _activatedManager;
        public UserCredential credential;
        //public IGoogleManager ActivatedManager { get => _activatedManager; set => _activatedManager = value; }

        private IGoogleSheetManager _activateSheetdManager;
        public IGoogleSheetManager ActivatedSheetManager { get => _activateSheetdManager; set => _activateSheetdManager = value; }

        /// <summary>
        /// 싱글톤. 인스턴스 호출
        /// </summary>
        /// <returns></returns>
        public static GoogleConnectionMannager GetInstance()
        {
            if (instance == null)
            {
                instance = new GoogleConnectionMannager("credentials.json");
            }
            return instance;
        }

        /// <summary>
        /// 구글 API 연결 매니저
        /// </summary>
        /// <param name="credentialsPath">크리덴셜 파일 경로</param>
        public GoogleConnectionMannager(string credentialsPath)
        {
            JoinServer(credentialsPath);
            if (instance == null)
                instance = this;
        }
        /// <summary>
        /// 서비스 매니저 의존성 등록.
        /// </summary>
        /// <returns></returns>
        public bool InjectService(IGoogleSheetManager Manager)
        {
            if (Manager == null)
                return false;
            //_activatedManager = Manager;
            _activateSheetdManager = Manager as IGoogleSheetManager;
            return true;
        }
        /// <summary>
        /// Credentials.json의 내용대로 구글 API 접속
        /// </summary>
        /// <param name="credentialsPath"> credentials.json 파일의 경로</param>
        public void JoinServer(string credentialsPath)
        {
            //Credentials.json의 내용대로 구글 API 접속
            using (var stream =
                new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
        }
    }
}
