<Query Kind="Program">
  <RuntimeVersion>6.0</RuntimeVersion>
</Query>

using System.Globalization;

void Main()
{
    // 通过年月日参数构建一个DateOnly
	var date = new DateOnly(2021,5,3);
	
	// 从现有DateTime类型的实例创建DateOnly实例
	var currentDateOnly = DateOnly.FromDateTime(DateTime.Now);
	Console.WriteLine(currentDateOnly);
	
	if(DateOnly.TryParse("28/9/1984",new CultureInfo("en-GB"), DateTimeStyles.None, out var result))
	{
		Console.WriteLine($"{result.Year}--{result.Month}--{result.Day}");
	}
	
	// 加减日期
	var newDateOnly = date.AddDays(1).AddMonths(1);
	Console.WriteLine(newDateOnly);
	
	
	// 构建TimeOnly实例
	var startTime = new TimeOnly(23,30);
	Console.WriteLine(startTime);
	var endTime = new TimeOnly(10,00,00);
	
	// 对TimeOnly实例进行数学操作
	var diff = endTime - startTime;
	Console.WriteLine(diff.TotalHours);
	
	// 从现有DateTime类型的实例创建TimeOnly实例	
	var currentTime = TimeOnly.FromDateTime(DateTime.Now);
	
	// 判断某个时间是否在某个时间范围内
	var isBetween = currentTime.IsBetween(startTime,endTime);
	Console.WriteLine($"Current time {(isBetween ? "is" : "is not")} between start and end.");
	
	// IsBetween可接受跨凌晨的时间范围
	var startTime1 = new TimeOnly(22, 00);
	var endTime1 = new TimeOnly(02, 00);
	var now = new TimeOnly(23, 25);
	var isBetween1 = now.IsBetween(startTime1, endTime1);
	Console.WriteLine($"{now} {(isBetween1 ? "is" : "is not")} between start and end.");
	
	// 比较使用循环时钟的时间操作符
	var startTime2 = new TimeOnly(08, 00);
	var endTime2 = new TimeOnly(09, 00);
	Console.WriteLine($"{startTime2 < endTime2}");
	
}

// You can define other methods, fields, classes and namespaces here
