namespace ModPlus
{
    // Данные, зависящие от версии
    internal class MpVersionData
    {
#if A2013
        public const string CurCadVers = "2013";
        public const string CurCadInternalVersion = "19.0";
#elif A2014
        public const string CurCadVers = "2014";
        public const string CurCadInternalVersion = "19.1";
#elif A2015
        public const string CurCadVers = "2015";
        public const string CurCadInternalVersion = "20.0";
#elif A2016
        public const string CurCadVers = "2016";
        public const string CurCadInternalVersion = "20.1";
#elif A2017
        public const string CurCadVers = "2017";
        public const string CurCadInternalVersion = "21.0";
#elif A2018
        public const string CurCadVers = "2018";
        public const string CurCadInternalVersion = "22.0";
#elif A2019
        public const string CurCadVers = "2019";
        public const string CurCadInternalVersion = "23.0";
#endif
    }
}
