using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ASCOM.Pulsar2;
using System.Threading;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Pulsar2UnitTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            var ra = telescope.CommandString("GR", false);
            var dec = telescope.CommandString("GD", false);
            var alignmentMode = telescope.AlignmentMode;
            telescope.Dispose();
        }

        [TestMethod]
        public void BufferTest()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            var ygi = telescope.CommandString("YGi", false);
            telescope.CommandBlind("Q", false);
            telescope.TargetRightAscension = 12.5;
            telescope.TargetDeclination = 31;
            var ms = telescope.CommandString("MS", false);
            var ygi2 = telescope.CommandString("YGi", false);
            var ygi3 = telescope.CommandString("YGi", false);
            telescope.CommandBlind("Q", false);
            telescope.Dispose();
        }

        [TestMethod]
        public void TrackingTest()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            telescope.Tracking = true; //YSS1,0
            telescope.TrackingRate = DriveRates.driveSidereal; //YGS
            telescope.TrackingRate = DriveRates.driveLunar; //YGS
            telescope.TrackingRate = DriveRates.driveSolar; //YGS, YSS2,0
            telescope.Tracking = false; //YSS0,0
            telescope.Dispose();
        }

        [TestMethod]
        public void ParkTest()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            telescope.Unpark();
            telescope.Unpark();
            var atPark = telescope.AtPark;
            telescope.Park();
            atPark = telescope.AtPark;
            telescope.Park();
            atPark = telescope.AtPark;
            //telescope.Unpark();
            atPark = telescope.AtPark;
            telescope.Dispose();
        }

        [TestMethod]
        public void RADec()
        {
            var telescope = new Telescope();
            
            telescope.Connected = true;
            telescope.TargetRightAscension = 12.5;
            telescope.TargetDeclination = 31;
            var ra = telescope.TargetRightAscension;
            var dec = telescope.TargetDeclination;

            var ra2 = telescope.RightAscension;
            var dec2 = telescope.Declination;

            telescope.Dispose();
        }

        [TestMethod]
        public void RADecRateGetSet()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            telescope.RightAscensionRate = 0.05;
            telescope.DeclinationRate = 0.05;
            telescope.Dispose();
        }

        [TestMethod]
        public void SideOfPier()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            var a = telescope.SideOfPier;
            telescope.Dispose();
        }

        [TestMethod]
        public void UTCTime()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            telescope.CommandString("GG", false);
            var t = telescope.UTCDate;
            telescope.UTCDate = DateTime.UtcNow;
            telescope.Dispose();
        }

        [TestMethod]
        public void LatitudeLongitude()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            var lat = telescope.SiteLatitude;
            var lng = telescope.SiteLongitude;

            telescope.SiteLatitude = 46.2144350;
            telescope.SiteLongitude = 14.4365550;

            lat = telescope.SiteLatitude;
            lng = telescope.SiteLongitude;

            telescope.Dispose();
        }

        [TestMethod]
        public void SyncTest()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            telescope.Unpark();
            telescope.TargetRightAscension = 9;
            telescope.TargetDeclination = 21;
            telescope.SyncToTarget();
            var ra1 = telescope.RightAscension;
            var dec1 = telescope.Declination;
            var slewing = telescope.Slewing;
            telescope.Dispose();
        }

        [TestMethod]
        public void Version()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            //telescope.CheckVersion(); 
            telescope.Dispose();
        }

        [TestMethod]
        public void SlewToTargetTest()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            var dec = telescope.TargetDeclination;
            var ra = telescope.TargetRightAscension;
            var dec1 = telescope.TargetDeclination;
            var ra1 = telescope.TargetRightAscension;
          telescope.TargetDeclination = 30;
            telescope.TargetRightAscension = 15;
            //TODO
            telescope.Dispose();
        }

        private string[] ParseYGS(string result)
        {
            //requested by Adnras - ygs sometimes returns "1" before RA rate which is wrong (e.g. "1+0.0000000,-0.0000308" instead of "+0.0000000,-0.0000308")
            var matches = Regex.Matches(result, "\\d*([+-][\\d.]+),([+-][\\d.]+)")[0];
            return new string[2] { matches.Groups[1].Value, matches.Groups[2].Value };
        }

        [TestMethod]
        public void MoveAxisTest()
        {
            var a = "1+0.0000000,-0.0000308";
            var b = "+0.0000000,-0.0000308";
            var c = "111+0.0000000,-0.0000308";

            var Ara = ParseYGS(a)[0];
            var Adec = ParseYGS(a)[1];
            var Bra = ParseYGS(b)[0];
            var Bdec = ParseYGS(b)[1];
            var Cra = ParseYGS(c)[0];
            var Cdec = ParseYGS(c)[1];

           
            var telescope = new Telescope();
            //telescope.Connected = true;
            //TODO
            //telescope.Dispose();
            telescope.MoveAxis(TelescopeAxes.axisSecondary, 2);
        }

        [TestMethod]
        public void SetSiderealTime()
        {
            var telescope = new Telescope();
            telescope.Connected = true;
            telescope.CommandBlind("SS " + DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture), false);
            telescope.Dispose();
        }
    }
}
