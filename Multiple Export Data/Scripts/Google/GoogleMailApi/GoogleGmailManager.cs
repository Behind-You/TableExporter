using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
/*
namespace MyGoogleServices
{
    /// <summary>
    /// Google Gmail 매니저. 싱글톤
    /// </summary>
    public class GoogleGmailManager : IGoogleManager
    {
        public ManagerTypes GetManagerTypes()
        {
            return ManagerTypes.Gmail;
        }
        public GoogleGmailManager(UserCredential credential)
        {
            _credential = credential;
        }

        static UserCredential _credential;
        //private GmailService m_GmailService;
        private bool GMailService_Enabled = false;
        public bool IsEnabled { get => GMailService_Enabled; set => GMailService_Enabled = value; }

        public BaseClientService GetService()
        {
            if (!GMailService_Enabled)
                return null;

            return m_GmailService;
        }
        public bool ActivateService()
        {
            //만약 크리덴셜 파일이 없을경우 실패 반환
            if (_credential == null)
                return false;
            // Create Gmail API service.
            m_GmailService = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _credential,
                ApplicationName = "Gmail API .NET Quickstart"
            });
            GMailService_Enabled = true;
            return GMailService_Enabled;
        }
    }
}*/
