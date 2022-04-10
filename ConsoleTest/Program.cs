using OpenXMLXLSXImporter;
using SSUT;
using System.Diagnostics;
using System.Reflection;


ConsistencyTests cs = new ConsistencyTests();

object l = new object();

Console.WriteLine("Started ConsoleTest");
uint loop = 0;

cs.LoopUsingNewDataSet<SpreadSheetInstructionBuilderTest>(r =>
{
    var timer = new Stopwatch();
    timer.Start();
    try
    {
        r.FullRangeCellTest();
        //TimeSpan timeTaken = timer.Elapsed;
        //string duration = "Time taken: " + timeTaken.ToString(@"m\:ss\.fff");
        lock (l)
        {
            loop++;
            if(loop % 10000 == 0)
            {
                Console.WriteLine($"Passed {loop} Records");
            }
        }
    }
    catch(Exception ex)
    {
        TimeSpan timeTaken = timer.Elapsed;
        string duration = "Test Failed After: " + timeTaken.ToString(@"m\:ss\.fff");
        lock (l)
        {
            Console.WriteLine(duration);
            Console.WriteLine(ex.ToString());
        }

    }



});

Console.WriteLine("end of ConsoleTest");