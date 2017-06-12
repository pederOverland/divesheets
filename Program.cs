using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Xml.Serialization;
using System.Xml;

namespace tmp
{
    class Program
    {
        class Person
        {
            public string firstName { get; set; }
            public string lastName { get; set; }
        }
        static void Main(string[] args)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Competition));
            Competition competition;
            using(XmlReader reader = XmlReader.Create(@"dc.xml")){
                competition = (Competition) serializer.Deserialize(reader);
            }
            var cParticipants = competition.participants.ToList();
            var e = competition.events.Items[0] as Event;
            var ef = competition.events.Items[1] as Event;
            var filename = @"ind.xlsx";
            var file = new FileInfo(filename);
            var skeleton = "Individual dive sheet";
            var names = new Person[]{
                new Person{firstName="Peder Hans", lastName="Øverland"},
                new Person{firstName="Test", lastName="Bruker"},
                new Person{firstName="Kine", lastName="Hellebust"},
                new Person{firstName="Harald", lastName="Pettersen"}
            };
            using (var package = new ExcelPackage(file))
            {
                e.divers.ToList().ForEach(d=>{
                    var participants = d.participant.ToList().Select(pa=>cParticipants.FirstOrDefault(x=>x.id==pa.id));
                    var p = participants.FirstOrDefault();
                    var cells = package.Workbook.Worksheets.Copy(skeleton, p.firstname.value + " " + p.lastname.value).Cells;
                    cells["B7"].Value = p.lastname.value;
                    cells["M7"].Value = p.firstname.value;
                    cells["V7"].Value = "NOR";
                    if(e.series.onem){
                        cells["J10"].Value = "X";
                    }
                    if(e.series.threem){
                        cells["J12"].Value = "X";
                    }
                    if(e.series.platform){
                        cells["J14"].Value = "X";
                    }
                    if(e.series.men){
                        cells["Q12"].Value = "X";
                    }
                    if(e.series.women){
                        cells["Q14"].Value = "X";
                    }
                    var line = 20;
                    d.divelist.dive.ToList().ForEach(dl=>{
                        var arr = dl.dive.ToList();
                        var start = arr.Count() == 5 ? 'C' : 'D';
                        arr.ForEach(dlc=>{
                            cells[(start++)+line.ToString()].Value = dlc.ToString();
                        });
                        cells["I"+line].Value = getHeight(dl.height);
                        cells["K"+line].Value = dl.dd;
                        line += 2;
                    });
                    var finalList = ef.divers.ToList().FirstOrDefault(x=>x.participant[0].id == p.id);
                    line = 45;
                });
                package.Workbook.Worksheets.Delete(skeleton);
                package.SaveAs(new FileInfo("export/"+DateTime.Now.ToString() + "_sheets.xlsx"));
            }
            Console.WriteLine("Hello World!");
        }

        private static object getHeight(height height)
        {
            switch(height){
                case height.Item1:return "1m";
                case height.Item3:return "3m";
                case height.Item5:return "5m";
                case height.Item75:return "7.5m";
                case height.Item10:return "10m";
            }
            return "1m";
        }
    }
}
