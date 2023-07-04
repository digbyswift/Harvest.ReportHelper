using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Harvest.ReportHelper
{
    public sealed class Program
    {
        public static async Task Main(string[] args)
        {
            try
            {
                var runner = new Runner();
                
                await Console.Out.WriteLineAsync("Using downloads folder: %USERPROFILE%\\Downloads");
                var canContinue = false;

                do
                {
                    var option = await runner.GetOptionAsync();
                    if (option > 0)
                    {
                        if (!await runner.TryGetFileNameAsync())
                            continue;
                    
                        switch (option)
                        {
                            case 1:
                                await runner.RunCheckAsync();
                                break;
                        
                            case 2:
                                await runner.RunCleanAsync();
                                break;
                        
                            case 3:
                                await runner.RunCleanAsync(deleteClientColumn: false, allowPrefix: false);
                                await runner.RunSplitAsync();
                                break;
                        }
                    }

                    await Console.Out.WriteAsync("Continue? [Y/n]: ");
                    
                    var continueValue = Console.ReadLine();
                    canContinue = String.IsNullOrWhiteSpace(continueValue) || (continueValue.ToUpper() ?? "Y") == "Y";
                    
                } while (canContinue);

            }
            catch (Exception ex)
            {
                await Console.Error.WriteLineAsync((ex.InnerException ?? ex).Message);
            }
        }
    }
}
