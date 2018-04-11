using System;

namespace ExcelHelperUnitTests
{
	internal class DiagnosticsHelper
	{
		internal static object GetCallingMethodName(int v)
		{
            return System.Reflection.MethodBase.GetCurrentMethod();
        }
	}
}