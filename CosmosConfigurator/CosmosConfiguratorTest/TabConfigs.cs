using CosmosConfigurator;

/// <summary>
/// Table File: ./test_excel.xls
/// </summary>
public class TestExcelConfig : TabRow
{
	
	public int Id;
	
	public string Name;
	
	public string[] StrArray;
	

	public override void Parse(string[] values)
	{
	
		Id = Get_int("", "");
	
		Name = Get_string("", "");
	
		StrArray = Get_string_array("", "");
	
	}

    public override object PrimaryKey
    {
        get
        {
            return Id;
        }
    }
}