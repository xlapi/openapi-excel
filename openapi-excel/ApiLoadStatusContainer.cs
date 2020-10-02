namespace openapi_excel
{
    public class ApiLoadStatusContainer
    {
        private ApiLoadStatus _Status;
        public ApiLoadStatus Status
        {
            get { return _Status; }
            set {
                _Status = value; 
            }
        }
    }

    public enum ApiLoadStatus
    {
        NotSet = 0,
        Loading = 1,
        Loaded = 2,
        ConnectionFailure = 3,
        UnknownFailure = 4,
        NotYetLoaded = 5
    }
}