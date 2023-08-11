using System.Globalization;
using System.Text;
using CsvHelper;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using TimeZoneConverter;

class ClassScheduleConvertor
{
    static void Main(string[] args)
    {

        if (args.Length != 2)
        {
            Console.WriteLine("Must have 2 args: input file ouput file");
        }
        using (var fs = new FileStream(args[0], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var reader = new StreamReader(fs, Encoding.Default))
        {

            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<ClassEvent>();
                //var calendar = new Ical.Net.Calendar();
                Dictionary<String, Ical.Net.Calendar> cals = new Dictionary<string, Ical.Net.Calendar>();
                foreach (var classev in records)
                {
                    Ical.Net.Calendar calendar;
                    if (!cals.TryGetValue(classev.person, out calendar))
                    {
                        calendar = new Ical.Net.Calendar();
                        cals.Add(classev.person, calendar);
                    }
                    List<WeekDay> days = new List<WeekDay>();
                    String daysSt = classev.days.ToUpper();
                    DateOnly startDay;
                    if (daysSt.Contains("T"))
                    {
                        startDay = DateOnly.Parse("2023-08-29");
                    }
                    else if (daysSt.Contains("W"))
                    {
                        startDay = DateOnly.Parse("2023-08-30");
                    }
                    else if (daysSt.Contains("R"))
                    {
                        startDay = DateOnly.Parse("2023-08-31");
                    }
                    else if (daysSt.Contains("F"))
                    {
                        startDay = DateOnly.Parse("2023-08-25");
                    }
                    else if (daysSt.Contains("M"))
                    {
                        startDay = DateOnly.Parse("2023-08-28");
                    }
                    else
                    {
                        Console.WriteLine("Error: no valid days");
                        continue;
                    }
                    DateTime startDate = startDay.ToDateTime(classev.start);
                    DateTime endDate = startDay.ToDateTime(classev.end);
                    foreach (char day in classev.days.ToUpper())
                    {
                        switch (day)
                        {
                            case 'M':
                                days.Add(new WeekDay(DayOfWeek.Monday));
                                break;
                            case 'T':
                                days.Add(new WeekDay(DayOfWeek.Tuesday));
                                break;
                            case 'W':
                                days.Add(new WeekDay(DayOfWeek.Wednesday));
                                break;
                            case 'R':
                                days.Add(new WeekDay(DayOfWeek.Thursday));
                                break;
                            case 'F':
                                days.Add(new WeekDay(DayOfWeek.Friday));
                                break;
                        }
                    }
                    //days.Sort();

                    // The first instance of an event taking place on July 1, 2016 between 07:00 and 08:00.
                    String name = classev.courseSub;
                    if (classev.courseName.Length != 0)
                    {
                        name += "-" + classev.courseNo;
                    }
                    if (classev.courseNo.Length != 0)
                    {
                        name += ": " + classev.courseName;
                    }
                    String location = "";
                    if (classev.room.Length != 0)
                    {
                        location = classev.building + " " + classev.room;
                    }
                    string tz = TZConvert.WindowsToIana("Eastern Standard Time");
                    // calendar.AddTimeZone(VTimeZone.FromDateTimeZone(tz));
                    var vEvent = new CalendarEvent
                    {
                        Start = new CalDateTime(startDate, tz),
                        End = new CalDateTime(endDate, tz),
                        Location = location,
                        Summary = name,
                        Status = "Free"
                    };
                    vEvent.AddProperty("X-MICROSOFT-CDO-BUSYSTATUS", "FREE");
                    // Recur daily through the end of the day, July 31, 2016
                    var recurrenceRule = new RecurrencePattern(FrequencyType.Weekly, interval: 1)
                    {
                        ByDay = days,
                        Until = DateTime.Parse("2023-12-08T23:59")
                    };

                    vEvent.RecurrenceRules = new List<RecurrencePattern> { recurrenceRule };
                    calendar.Events.Add(vEvent);
                    if (daysSt.Contains("R"))
                    {
                        DateOnly specDate = DateOnly.Parse("2023-08-24");
                        switch (classev.start.Hour)
                        {
                            case 8:
                                {
                                    startDate = specDate.ToDateTime(new TimeOnly(11, 10));
                                    endDate = specDate.ToDateTime(new TimeOnly(11, 40));
                                    break;
                                }
                            case 9:
                                {
                                    startDate = specDate.ToDateTime(new TimeOnly(11, 50));
                                    endDate = specDate.ToDateTime(new TimeOnly(12, 20));
                                    break;
                                }
                            case 10:
                                {
                                    startDate = specDate.ToDateTime(new TimeOnly(12, 30));
                                    endDate = specDate.ToDateTime(new TimeOnly(13, 0));
                                    break;
                                }
                            case 12:
                                {
                                    startDate = specDate.ToDateTime(new TimeOnly(13, 10));
                                    endDate = specDate.ToDateTime(new TimeOnly(13, 40));
                                    break;
                                }
                            case 13:
                                {
                                    startDate = specDate.ToDateTime(new TimeOnly(13, 50));
                                    endDate = specDate.ToDateTime(new TimeOnly(14, 20));
                                    break;
                                }
                            case 15:
                                {
                                    startDate = specDate.ToDateTime(new TimeOnly(14, 30));
                                    endDate = specDate.ToDateTime(new TimeOnly(15, 0));
                                    break;
                                }
                            default:
                                {
                                    continue;
                                }
                        }
                        var vEvent2 = new CalendarEvent
                        {
                            Start = new CalDateTime(startDate, tz),
                            End = new CalDateTime(endDate, tz),
                            Location = location,
                            Summary = name,
                            Status = "Free"
                        };
                        vEvent2.AddProperty("X-MICROSOFT-CDO-BUSYSTATUS", "FREE");
                        calendar.Events.Add(vEvent2);

                    }

                }

                foreach (KeyValuePair<string, Ical.Net.Calendar> kvp in cals)
                {
                    string tz = TZConvert.WindowsToIana("Eastern Standard Time");
                    kvp.Value.AddTimeZone(VTimeZone.FromDateTimeZone(tz));
                    var serializer = new CalendarSerializer();
                    var serializedCalendar = serializer.SerializeToString(kvp.Value);
                    using (var writer = new StreamWriter(args[1] + kvp.Key + ".ics"))
                    {
                        writer.Write(serializedCalendar);
                    }
                }
            }
        }
    }
}