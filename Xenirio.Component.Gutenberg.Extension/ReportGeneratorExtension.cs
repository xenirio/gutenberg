using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xenirio.Component.Gutenberg.Model;

namespace Xenirio.Component.Gutenberg.Extensions
{
    public static class ReportGeneratorJsonExtension
    {
        public static void setJsonObject(this ReportGenerator context, JObject json)
        {
            var flattenContents = new Dictionary<string, JToken>();
            var contentKeys = json.Properties().Where(p => p.Name != "Style").Select(p => p.Name);
            foreach (var key in contentKeys)
            {
                var tokens = json.ContainsKey(key) ? json[key].DeserializeAndFlatten() : new Dictionary<string, JToken>();
                if (tokens.Any())
                {
                    foreach (var token in tokens)
                    {
                        flattenContents.Add(string.Format("{0}.{1}", key, token.Key), token.Value);
                    }
                }
            }
            var flattenStyles = json.ContainsKey("Style") ? json["Style"].DeserializeAndFlatten(1) : new Dictionary<string, JToken>();
            foreach (var item in flattenContents)
            {
                var contentKey = item.Key;
                var token = item.Value;
                if (token.Type == JTokenType.Array)
                {
                    if (token[0].Type == JTokenType.Array)
                    {
                        var values = token.Children().Select(t => t.Children().Select(c => ((JValue)c).Value.ToString()).ToArray()).ToArray();
                        if (flattenStyles.ContainsKey(contentKey))
                        {
                            var styles = flattenStyles[contentKey].ToObject<ReportLabelStyle[][]>();
                            var valueStyles = values.Select((v, iv) => v.Select((e, ie) => new KeyValuePair<string, ReportLabelStyle>(e, styles[iv][ie])).ToArray()).ToArray();
                            context.setTableParagraph(contentKey, valueStyles);
                        }
                        else
                        {
                            context.setTableParagraph(contentKey, values);
                        }
                    }
                    else
                        context.setParagraphs(contentKey, token.Children().Select(t => ((JValue)t).Value.ToString()).ToArray());
                }
                else
                {
                    var value = ((JValue)item.Value).Value.ToString();
                    if (flattenStyles.ContainsKey(contentKey))
                    {
                        var styleToken = flattenStyles[contentKey].ToObject<ReportLabelStyle>();
                        context.setParagraph(contentKey, value, styleToken);
                    }
                    else
                    {
                        context.setParagraph(contentKey, value);
                    }
                }
            }
        }
    }

    public static class JObjectExtension
    {
        public static Dictionary<string, JToken> DeserializeAndFlatten(this JToken token, int level = 0)
        {
            Dictionary<string, JToken> dict = new Dictionary<string, JToken>();
            fillDictionaryFromJToken(dict, token, "");

            var flattenDict = new Dictionary<string, JToken>();
            foreach (var flat in dict)
            {
                var flatKeys = flat.Key.Split('.');
                var keys = new List<string>();
                var val = token;
                for (var i = 0; i < flatKeys.Length; i++)
                {
                    var raw = val[flatKeys[i]];
                    if (i < flatKeys.Length - level || raw.Type == JTokenType.Array)
                    {
                        keys.Add(flatKeys[i]);
                        val = raw;
                    }
                }
                var flattenKey = string.Join(".", keys);
                if (!flattenDict.ContainsKey(flattenKey))
                    flattenDict.Add(flattenKey, val);
            }
            return flattenDict;
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
