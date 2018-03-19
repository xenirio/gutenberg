namespace Handshakes.Api.Report.Model
{
	internal abstract class ReportElement
    {
		public string Key { get; set; }

		protected object value { get; set; }

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
