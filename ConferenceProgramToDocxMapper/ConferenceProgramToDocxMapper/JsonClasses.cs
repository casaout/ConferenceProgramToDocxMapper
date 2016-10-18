using System.Collections.Generic;

/// <summary>
/// Classes necessary to parse the JSON object
/// (generated with: http://json2csharp.com/)
/// </summary>
namespace ConferenceProgramToDocxMapper
{
    public class Organization
    {
        public List<string> GeneralChairs { get; set; }
        public List<string> PCChairs { get; set; }
    }

    public class SocialFeed
    {
        public string Name { get; set; }
        public List<string> Keywords { get; set; }
        public string URL { get; set; }
    }

    public class GPS
    {
        public int Latitude { get; set; }
        public int Longitude { get; set; }
    }

    public class VenueInfo
    {
        public string Name { get; set; }
        public GPS GPS { get; set; }
    }

    public class GPS2
    {
        public int Latitude { get; set; }
        public int Longitude { get; set; }
    }

    public class Element
    {
        public string XamlName { get; set; }
        public string Type { get; set; }
        public GPS2 GPS { get; set; }
        public string MapLabel { get; set; }
        public string URL { get; set; }
    }

    public class InfoPage
    {
        public string xaml { get; set; }
        public List<Element> Elements { get; set; }
    }

    public class Item
    {
        public string Title { get; set; }
        public string Type { get; set; }
        public string Key { get; set; }
        public string URL { get; set; }
        public string URLvideo { get; set; }
        public string URLinfo { get; set; }
        public string URLmaterial { get; set; }
        public string DOI { get; set; }
        public string PersonsString { get; set; }
        public string AffiliationsString { get; set; }
        public List<string> Authors { get; set; }
        public List<string> Affiliations { get; set; }
        public string Abstract { get; set; }
        public string Award { get; set; }
        public string Keywords { get; set; }
    }

    public class Session
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string ShortTitle { get; set; }
        public string Type { get; set; }
        public string ShortType { get; set; }
        public string Key { get; set; }
        public string Day { get; set; }
        public string Time { get; set; }
        public string Location { get; set; }
        public string LocationIndex { get; set; }
        public string ChairsString { get; set; }
        public List<object> Chairs { get; set; }
        public string SponsoredBy { get; set; }
        public bool Workshop { get; set; }
        public string Comment { get; set; }
        public string URL { get; set; }
        public string Abstract { get; set; }
        public List<string> Items { get; set; }
    }

    public class Person
    {
        public string Name { get; set; }
        public string Affiliation { get; set; }
        public string Key { get; set; }
        public string Id { get; set; }
        public string Bio { get; set; }
        public string URLhp { get; set; }
        public string URLphoto { get; set; }
        public string URLgs { get; set; }
        public string URLmsa { get; set; }
        public string URLtw { get; set; }
        public string URLgp { get; set; }
        public string URLfb { get; set; }
        public string URLli { get; set; }
    }

    public class RootObject
    {
        public int DataRevision { get; set; }
        public string Preliminary { get; set; }
        public string Event { get; set; }
        public string Name { get; set; }
        public string NameFull { get; set; }
        public string Description { get; set; }
        public string Date { get; set; }
        public string DateStart { get; set; }
        public string DateEnd { get; set; }
        public string Location { get; set; }
        public string Sponsors { get; set; }
        public string URL { get; set; }
        public string MultiTrack { get; set; }
        public string UseMiniPage { get; set; }
        public string NewPagePerDay { get; set; }
        public int NumOfParallelTracks { get; set; }
        public Organization Organization { get; set; }
        public List<SocialFeed> SocialFeeds { get; set; }
        public VenueInfo VenueInfo { get; set; }
        public InfoPage InfoPage { get; set; }
        public List<string> SessionPriorities { get; set; }
        public List<Item> Items { get; set; }
        public List<Session> Sessions { get; set; }
        public List<Person> People { get; set; }
    }
}
