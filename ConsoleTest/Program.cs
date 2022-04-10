using OpenXMLXLSXImporter;
using SSUT;
using System.Diagnostics;
using System.Reflection;


ConsistencyTests cs = new ConsistencyTests();

object l = new object();

Console.WriteLine("Started ConsoleTest");

TimeSpan[] storedDurations = new TimeSpan[ConsistencyTests.LOOPS];




cs.LoopUsingNewDataSet<SpreadSheetInstructionBuilderTest>((i,r) =>
{
    var timer = new Stopwatch();
    timer.Start();
    string fail = null;
    try
    {
        r.FullRangeCellTest();
    }
    catch(Exception ex)
    {
        fail = ex.ToString();
    }

    timer.Stop();
    TimeSpan timeTaken = timer.Elapsed;
    storedDurations[i] = timeTaken;
    if (fail != null)
    {
        string duration = "Test Failed After: " + timeTaken.ToString(@"m\:ss\.fff");
        lock (l)
        {
            if (i % 10000 == 0)
            {
                Console.WriteLine($"Passed {i} Records");
            }
            Console.WriteLine(duration);
            Console.WriteLine(fail);
        }
    }
    else if (i % 10000 == 0)
    {
        lock (l)
        {
            Console.WriteLine($"Passed {i} Records");
            Console.WriteLine($"last test took {timeTaken.ToString(@"m\:ss\.fff")}");
        }
    }
});

Console.WriteLine("end of ConsoleTest");