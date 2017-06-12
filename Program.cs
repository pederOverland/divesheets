using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Xml.Serialization;
using System.Xml;
using System.Collections.Generic;

namespace tmp
{
    class Program
    {
        class Person
        {
            public string firstName { get; set; }
            public string lastName { get; set; }
        }
        static List<Participant> cParticipants;
        static List<Team> cTeams;
        static void Main(string[] args)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Competition));
            Competition competition;
            using (XmlReader reader = XmlReader.Create(@"dc.xml"))
            {
                competition = (Competition)serializer.Deserialize(reader);
            }
            cParticipants = competition.participants.ToList();
            cTeams = competition.teams.ToList();
            var titles = competition.events.Items.ToList().Select(x =>
            {
                var ev = x as Event;
                if (ev != null)
                {
                    return ev.title.value;
                }
                var es = x as EventSet;
                return es.name.value;
            }).ToList();
            competition.events.Items.ToList().ForEach(ev =>
            {
                var e = ev as Event;
                if (e == null)
                {
                    //createJump(ev as )
                }
                else if (e.synchro)
                {
                    createSynchro(e);
                }
                else
                {
                    createJump(e);
                }
            });
            Console.WriteLine("Export finished!");
        }

        private static void createSynchro(Event e)
        {
            var filename = @"synchro.xlsx";
            var file = new FileInfo(filename);
            var skeleton = "Dive Sheet Synchro";
            using (var package = new ExcelPackage(file))
            {
                e.divers.ToList().ForEach(d =>
                {
                    var participants = d.participant.ToList().Select(pa => cParticipants.FirstOrDefault(x => x.id == pa.id)).ToList();
                    var teamString = String.Join("/", participants.Select(x=>cTeams.FirstOrDefault(t=>t.id==x.team.id).shortname.value));
                    var p1 = participants.FirstOrDefault();
                    var p2 = participants.Skip(1).FirstOrDefault();
                    var cells = package.Workbook.Worksheets.Copy(skeleton, p1.lastname.value + "_" + p2.lastname.value).Cells;
                    cells["M13"].Value = e.title.value;
                    cells["C7"].Value = p1.lastname.value;
                    cells["N7"].Value = p1.firstname.value;
                    cells["T7"].Value = p1.birthyear.value;
                    cells["C10"].Value = p2.lastname.value;
                    cells["N10"].Value = p2.firstname.value;
                    cells["T10"].Value = p2.birthyear.value;
                    if (e.series.men)
                    {
                        cells["E13"].Value = "X";
                    }
                    if (e.series.women)
                    {
                        cells["J13"].Value = "X";
                    }
                    var line = 19;
                    var currentDive = "";
                    d.divelist.dive?.ToList().ForEach(dl =>
                    {
                        if (dl.number == currentDive)
                        {
                            return;
                        }
                        currentDive = dl.number;
                        var arr = dl.dive.ToList();
                        var start = arr.Count() == 5 ? 'D' : 'E';
                        arr.ForEach(dlc =>
                        {
                            cells[(start++) + line.ToString()].Value = dlc.ToString();
                        });
                        cells["J" + line].Value = getHeight(dl.height);
                        cells["L" + line].Value = dl.dd;
                        line += 2;
                    });
                });
                package.Workbook.Worksheets.Copy(skeleton, "Blank Sheet");
                package.Workbook.Worksheets.Delete(skeleton);
                package.SaveAs(new FileInfo("export/" + DateTime.Now.ToString() + "_" + e.title.value + ".xlsx"));
            }
        }
        private static void createJump(Event e)
        {
            var filename = @"jump.xlsx";
            var file = new FileInfo(filename);
            var skeleton = "JUMP EVENT";
            using (var package = new ExcelPackage(file))
            {
                e.divers.ToList().ForEach(d =>
                {
                    var participants = d.participant.ToList().Select(pa => cParticipants.FirstOrDefault(x => x.id == pa.id)).ToList();
                    var teamString = String.Join("/", participants.Select(x=>cTeams.FirstOrDefault(t=>t.id==x.team.id).shortname.value));
                    var sheetName = string.Join("-", participants.Select(p => p.lastname.value)) + "("+d.id+")";
                    var cells = package.Workbook.Worksheets.Copy(skeleton, sheetName).Cells;
                    cells["Q10"].Value = teamString;
                    var line = 7;
                    participants.ForEach(p =>
                    {
                        cells["A" + line].Value = p.lastname?.value + " - " + p.firstname?.value;
                        cells["I" + line].Value = p.birthmonth?.value;
                        cells["J" + line].Value = p.birthyear?.value;
                        cells["L" + line].Value = p.gender?.value;
                        line += 3;
                    });
                    line = 20;
                    var currentDive = "";
                    d.divelist.dive?.ToList().ForEach(dl =>
                    {
                        cells["B"+line].Value = dl.diver;
                        if (dl.number == currentDive)
                        {
                            cells["B"+(line-1)].Value = dl.diver;
                            return;
                        }
                        currentDive = dl.number;
                        cells["F"+line].Value = dl.dive.Substring(0, dl.dive.Count()-1);
                        cells["G"+line].Value = dl.dive.Substring(dl.dive.Count()-1, 1);
                        cells["I" + line].Value = dl.dd;
                        line += 2;
                    });
                });
                package.Workbook.Worksheets.Copy(skeleton, "Blank Sheet");
                package.Workbook.Worksheets.Delete(skeleton);
                package.SaveAs(new FileInfo("export/" + DateTime.Now.ToString() + "_" + e.title.value + ".xlsx"));
            }
        }

        private static void createIndividual(Event e)
        {
            var filename = @"ind.xlsx";
            var file = new FileInfo(filename);
            var skeleton = "Individual dive sheet";
            using (var package = new ExcelPackage(file))
            {
                e.divers.ToList().ForEach(d =>
                {
                    var participants = d.participant.ToList().Select(pa => cParticipants.FirstOrDefault(x => x.id == pa.id));
                    var teamString = String.Join("/", participants.Select(x=>cTeams.FirstOrDefault(t=>t.id==x.team.id).shortname.value));
                    var p = participants.FirstOrDefault();
                    var cells = package.Workbook.Worksheets.Copy(skeleton, p.firstname.value + " " + p.lastname.value).Cells;
                    cells["B7"].Value = p.lastname.value;
                    cells["M7"].Value = p.firstname.value;
                    cells["V7"].Value = teamString;
                    if (e.series.onem)
                    {
                        cells["J10"].Value = "X";
                    }
                    if (e.series.threem)
                    {
                        cells["J12"].Value = "X";
                    }
                    if (e.series.platform)
                    {
                        cells["J14"].Value = "X";
                    }
                    if (e.series.men)
                    {
                        cells["Q12"].Value = "X";
                    }
                    if (e.series.women)
                    {
                        cells["Q14"].Value = "X";
                    }
                    var line = 20;
                    d.divelist.dive?.ToList().ForEach(dl =>
                    {
                        var arr = dl.dive.ToList();
                        var start = arr.Count() == 5 ? 'C' : 'D';
                        arr.ForEach(dlc =>
                        {
                            cells[(start++) + line.ToString()].Value = dlc.ToString();
                        });
                        cells["I" + line].Value = getHeight(dl.height);
                        cells["K" + line].Value = dl.dd;
                        line += 2;
                    });
                });
                package.Workbook.Worksheets.Copy(skeleton, "Blank Sheet");
                package.Workbook.Worksheets.Delete(skeleton);
                package.SaveAs(new FileInfo("export/" + DateTime.Now.ToString() + "_" + e.title.value + ".xlsx"));
            }
        }

        private static object getHeight(height height)
        {
            switch (height)
            {
                case height.Item1: return "1m";
                case height.Item3: return "3m";
                case height.Item5: return "5m";
                case height.Item75: return "7.5m";
                case height.Item10: return "10m";
            }
            return "1m";
        }
    }
}
