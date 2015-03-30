using CosmosConfigurator;

/// <summary>
/// Auto Generate for Tab File: test_excel.bytes
/// </summary>
public class TestExcelConfig : TabRow
{
	public static readonly string TabFilePath = "test_excel.bytes";
	
	public int Id;  // ID Comlun
	
	public string Name;  // Name
	
	public string[] StrArray;  // 
	

	public override void Parse(string[] values)
	{
	
		// ID Comlun
		Id = Get_int(values[0], "0");
	
		// Name
		Name = Get_string(values[1], "");
	
		// 
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