using System;
using System.Reflection;

namespace ExcelHelperUnitTests
{
	internal class ReflectionHelper
	{
		internal static string GetExecutingAssemblyFolder(Assembly assembly)
		{
            var location = new Uri(assembly.CodeBase);
            return new System.IO.FileInfo(location.AbsolutePath).Directory.FullName;

            //return assembly.Location;

		}
	}
}