using System;
using CosmosConfigurator;

/// <summary>
/// Auto Generate for Tab File: test_excel.bytes
/// </summary>
public partial class TestExcelConfig : TabRow
{
	public static readonly string TabFilePath = "test_excel.bytes";
	
	public int Id;  // ID Column/编号
	
	public string Name;  // Name/名字
	
	public string[] StrArray;  // ArrayTest/测试数组
	

	public override void Parse(string[] values)
	{
	
		// ID Column/编号
		Id = Get_int(values[0], "0");
	
		// Name/名字
		Name = Get_string(values[1], "");
	
		// ArrayTest/测试数组
		StrArray = Get_string_array(values[2], "");
	
	}

    public override object PrimaryKey
    {
        get
        {
            return Id;
        }
    }
}