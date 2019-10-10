using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xenirio.Component.Gutenberg.Extensions
{
    public static class ReportGeneratorJsonExtension
    {
        public static void setJsonObject(this ReportGenerator context, JObject json)
        {
            var flattens = json.DeserializeAndFlatten();
            foreach (var item in flattens)
            {
                var token = item.Value;
                if(token.Type == JTokenType.Array)
                {
                    if(token[0].Type == JTokenType.Array)
                        context.setTableParagraph(item.Key, token.Children().Select(t => t.Children().Select(c => ((JValue)c).Value.ToString()).ToArray()).ToArray());
                    else
                        context.setParagraphs(item.Key, token.Children().Select(t => ((JValue)t).Value.ToString()).ToArray());
                }
                else
                {
                    context.setParagraph(item.Key, ((JValue)item.Value).Value.ToString());
                }
            }
        }
    }

    public static class JObjectExtension
    {
        public static Dictionary<string, JToken> DeserializeAndFlatten(this JObject json)
        {
            Dictionary<string, JToken> dict = new Dictionary<string, JToken>();
            JToken token = JToken.Parse(json.ToString());
            fillDictionaryFromJToken(dict, token, "");
            return dict;
        }

        private static void fillDictionaryFromJToken(Dictionary<string, JToken> dict, JToken token, string prefix)
        {
            switch (token.Type)
            {
                case JTokenType.Object:
                    foreach (JProperty prop in token.Children<JProperty>())
                    {
                        fillDictionaryFromJToken(dict, prop.Value, join(prefix, prop.Name));
                    }
                    break;

                default:
                    dict.Add(prefix, token);
                    break;
            }
        }

        private static string join(string prefix, string name)
        {
            return (string.IsNullOrEmpty(prefix) ? name : prefix + "." + name);
        }
    }
}
