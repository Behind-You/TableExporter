using Google.Apis.Services;


namespace MyGoogleServices
{
    public enum ManagerTypes
    {
        Connection,
        Sheet,
        Gmail
    }
    public interface IGoogleManager
    {
        BaseClientService GetService();
        bool ActivateService();
        ManagerTypes GetManagerTypes();

    }

}
