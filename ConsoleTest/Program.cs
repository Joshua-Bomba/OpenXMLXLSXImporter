using OpenXMLXLSXImporter;
using SSUT;
using System.Diagnostics;
using System.Reflection;


ConsistencyTests cs = new ConsistencyTests();

object l = new object();

Console.WriteLine("Started ConsoleTest");

TimeSpan[] storedDurations = new TimeSpan[ConsistencyTests.LOOPS];

const uint LOG_FREQUENCY = 10000;

static async Task LogResults(TimeSpan[] storedDurations)
{
    Console.WriteLine("calculating statistics");
    Task<TimeSpan> maxResult = Task.Run<TimeSpan>(() =>
    {
        return storedDurations.Max();
    });

    Task<TimeSpan> minResult = Task.Run<TimeSpan>(() =>
    {
        return storedDurations.Min();
    });
    Task<TimeSpan> averageResult = Task.Run<TimeSpan>(() =>
    {
        return TimeSpan.FromSeconds(storedDurations.Average(x => x.TotalSeconds));
    });

    TimeSpan max = await maxResult;
    Console.WriteLine($"Slowest Test Took {max.ToString(@"m\:ss\.fff")}");
    TimeSpan min = await minResult;
    Console.WriteLine($"Fastest Test Tooken {min.ToString(@"m\:ss\.fff")}");
    TimeSpan average = await averageResult;
    Console.WriteLine($"Average Test Took {average}");
}
await LogResults(storedDurations);


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
            if (i % LOG_FREQUENCY == 0)
            {
                Console.WriteLine($"Passed {i} Records");
            }
            Console.WriteLine(duration);
            Console.WriteLine(fail);
        }
    }
    else if (i % LOG_FREQUENCY == 0)
    {
        lock (l)
        {
            Console.WriteLine($"Passed {i} Records");
            Console.WriteLine($"last test took {timeTaken.ToString(@"m\:ss\.fff")}");
        }
    }
});

Console.WriteLine("end of ConsoleTest");
await LogResults(storedDurations);