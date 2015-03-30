
/// <summary>
/// Table File: {{ TabFilePath }}
/// </summary>
public class {{ClassName}}Config : TabRow
{
	public static readonly string TabFilePath = "{{ TabFilePath }}";
	{% for field in Fields %}
	public {{ field.Type }} {{ field.Name}};  // {{ field.Comment }}
	{% endfor %}

	public override void Parse(string[] values)
	{
	{% for field in Fields %}
		// {{ field.Comment }}
		{{ field.Name}} = Get_{{ field.Type | replace:'\[\]','_array' }}(values[{{ field.Index }}], "{{ field.DefaultValue }}");
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