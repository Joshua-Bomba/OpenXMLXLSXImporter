﻿using OpenXMLXLSXImporter;
using SSUT;
using System.Diagnostics;
using System.Reflection;
using System.Timers;

const uint LOG_FREQUENCY = 10000;


//await Test<SpreadSheetInstructionBuilderTest>(r => r.FullRangeCellTest());
await Test<ConcurrencyTests>(x => x.MultipleSheetsBundlerTest());

static async Task Test<TProp>(Action<TProp> testAction) where TProp : BaseTest, new()
{
    ConsistencyTests cs = new ConsistencyTests();

    object l = new object();

    Console.WriteLine("Started ConsoleTest");

    TimeSpan?[] storedDurations = new TimeSpan?[ConsistencyTests.LOOPS];

    int lastElement = 0;
    CancellationTokenSource cts = new CancellationTokenSource();
    CancellationToken token = cts.Token;
    Task freezeCheck = Task.Run(async () =>
    {
        while(true)
        {
            await Task.Delay(10000, token);
            token.ThrowIfCancellationRequested();
            for (int i = 0; i < storedDurations.Length; i++)
            {
                if (storedDurations[i] == null)
                {
                    if (i == lastElement)
                    {
                        lock (l)
                        {
                            Console.WriteLine($"The Program Might Be Struck since {i} was null 10 seconds ago");
                        }
                    }
                    else
                    {
                        lastElement = i;
                    }
                }
            }
        }
    });

    var wholeTest = new Stopwatch();
    wholeTest.Start();
    cs.LoopUsingNewDataSet<TProp>((i, r) =>
    {
        var timer = new Stopwatch();
        timer.Start();
        string fail = null;
        try
        {
            testAction(r);
        }
        catch (Exception ex)
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
    wholeTest.Stop();
    cts.Cancel();
    Console.WriteLine("end of ConsoleTest");
    await LogResults(storedDurations, wholeTest.Elapsed);
}

static async Task LogResults(TimeSpan?[] storedDurations,TimeSpan totalDuration)
{
    Console.WriteLine("calculating statistics");
    Task<TimeSpan> maxResult = Task.Run<TimeSpan>(() =>
    {
        return storedDurations.Max().Value;
    });

    Task<TimeSpan> minResult = Task.Run<TimeSpan>(() =>
    {
        return storedDurations.Min().Value;
    });
    Task<TimeSpan> averageResult = Task.Run<TimeSpan>(() =>
    {
        return TimeSpan.FromSeconds(storedDurations.Average(x => x.Value.TotalSeconds));
    });

    Console.WriteLine($"Whole Test Took {totalDuration.ToString(@"m\:ss\.ffff")}");
    TimeSpan max = await maxResult;
    Console.WriteLine($"Slowest Test Took {max.ToString(@"m\:ss\.ffff")}");
    TimeSpan min = await minResult;
    Console.WriteLine($"Fastest Test Tooken {min.ToString(@"m\:ss\.ffff")}");
    TimeSpan average = await averageResult;
    Console.WriteLine($"Average Test Took {average.ToString(@"m\:ss\.ffff")}");
}



