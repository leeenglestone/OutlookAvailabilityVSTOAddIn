using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace OutlookAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnShowAvailability_Click(object sender, RibbonControlEventArgs e)
        {
            var outlookApp = Globals.ThisAddIn.Application;
            var calendar = outlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            var items = calendar.Items;
            items.IncludeRecurrences = true;
            items.Sort("[Start]");

            DateTime weekStart = DateTime.Today;
            while (weekStart.DayOfWeek != DayOfWeek.Monday)
                weekStart = weekStart.AddDays(-1);

            weekStart = weekStart.AddDays(7);

            DateTime weekEnd = weekStart.AddDays(5); // Monday to Friday

            string[] days = { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday" };
            var output = new StringBuilder();
            output.AppendLine($"Availability w/c {weekStart:dd MMMM yyyy}");

            for (int i = 0; i < 5; i++)
            {
                DateTime dayStart = weekStart.AddDays(i).Date.AddHours(9);  // 9 AM
                DateTime dayEnd = weekStart.AddDays(i).Date.AddHours(17);   // 5 PM
                string dayName = days[i];

                string filter = $"[Start] < '{dayEnd:g}' AND [End] > '{dayStart:g}'";
                var busyItems = items.Restrict(filter);

                List<(DateTime, DateTime)> busyTimes = new List<(DateTime, DateTime)>();
                foreach (Outlook.AppointmentItem item in busyItems)
                {
                    busyTimes.Add((item.Start, item.End));
                }

                List<(DateTime, DateTime)> freeTimes = GetFreeTime(dayStart, dayEnd, busyTimes);

                if (freeTimes.Count == 0)
                {
                    output.AppendLine($"{dayName}: Out all day");
                }
                else
                {
                    string times = string.Join(", ", freeTimes.Select(t =>
    $"{t.Item1:h:mmtt} to {t.Item2:h:mmtt}".ToLower()));
                    output.AppendLine($"{dayName}: {times}");
                }
            }

            Clipboard.SetText(output.ToString());

            MessageBox.Show(output.ToString(), "Copied to clipboard");

        }

        private List<(DateTime, DateTime)> GetFreeTime(DateTime workStart, DateTime workEnd, List<(DateTime, DateTime)> busyTimes)
        {
            var freeSlots = new List<(DateTime, DateTime)>();

            var sortedBusy = busyTimes.OrderBy(x => x.Item1).ToList();
            DateTime current = workStart;

            foreach (var (busyStart, busyEnd) in sortedBusy)
            {
                if (busyStart > current)
                {
                    DateTime freeEnd = busyStart < workEnd ? busyStart : workEnd;
                    if (freeEnd > current)
                        freeSlots.Add((current, freeEnd));
                }
                current = busyEnd > current ? busyEnd : current;
            }

            if (current < workEnd)
                freeSlots.Add((current, workEnd));

            // ✅ Merge adjacent free slots
            var merged = new List<(DateTime, DateTime)>();
            foreach (var slot in freeSlots)
            {
                if (merged.Count == 0)
                {
                    merged.Add(slot);
                }
                else
                {
                    var last = merged.Last();
                    if (last.Item2 == slot.Item1)
                    {
                        merged[merged.Count - 1] = (last.Item1, slot.Item2);
                    }
                    else
                    {
                        merged.Add(slot);
                    }
                }
            }

            return merged;
        }
    }
}
