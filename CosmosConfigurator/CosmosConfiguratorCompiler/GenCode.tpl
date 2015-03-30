
/// <summary>
/// Table File: {{ TableFilePath }}
/// </summary>
public class {{ClassName}}Config : TabRow
{
	{% for field in Fields %}
	public {{ field.Type }} {{ field.Name}};
	{% endfor %}

	public override void Parse(string[] values)
	{
	{% for field in Fields %}
		{{ field.Name}} = Get_{{ field.Type | replace:'\[\]','_array' }}("", "");
	{% endfor %}
	}

    public override object PrimaryKey
    {
        get
        {
            return {{ PrimaryKey }};
        }
    }
}