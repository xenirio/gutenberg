namespace Xenirio.Component.Gutenberg.Model
{
	internal abstract class ReportElement
    {
		public string Key { get; set; }

		public object value { get; set; }

		protected T GetValue<T>()
		{
			return (T)value;
		}

		protected void SetValue<T>(T value)
		{
			this.value = value;
		}
	}
}
