using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
namespace MapSharepointMetdata
{
    public class SharepointFields
    {
        private const string ITEM_NAME_FIELD = "FileLeafRef";
        private const string INDEX_EXCEPTION = "Index Exception Occured in: ";
        private const string COMPLETED_RUN_MESSAGE = "Completed Run in: ";

        public int WAIT_TIME_AFTER_COMPLETEION_FOR_DISPLAYING_CONSOLE = 10;
        public string[,] sources { get; set; }
        public string[,] fieldsToMap { get; set; }
        public static void viewAvailableFields(string directoryOfFileToWriteTo, string sourceSiteUrl, string sourceLibrary, string targetSiteUrl, string targetLibrary)
        {
            string[,] mapped_field_names = get_mapped_field_names(sourceSiteUrl, sourceLibrary, targetSiteUrl, targetLibrary);
            write_to_textfile(mapped_field_names, directoryOfFileToWriteTo);
        }
        public string targetLibarary { get; set; }
        public string target_siteUrl { get; set; }
        public string booleanFilterField { get; set; }
        public string[,] lookupFieldsToMap { get; set; }
        public string filter_field { get; set; }
        public bool MapFields()
        {
            try
            {
                System.Diagnostics.Stopwatch watch = System.Diagnostics.Stopwatch.StartNew();
                for (int i = 0; i < sources.GetLength(0); i++)
                {
                    map_fields(sources[i, 1], targetLibarary, sources[i, 0], target_siteUrl, ITEM_NAME_FIELD, lookupFieldsToMap, booleanFilterField, fieldsToMap);
                }
                Console.WriteLine(COMPLETED_RUN_MESSAGE + watch.Elapsed.ToString());
                System.Threading.Thread.Sleep(WAIT_TIME_AFTER_COMPLETEION_FOR_DISPLAYING_CONSOLE * 1000);
                return true;
            }
            catch { return false; }

        }// edit this function to change functionality
        public bool IsFieldEditable(Microsoft.SharePoint.Client.List targ_list, string fname)
        {
            FieldCollection collField = targ_list.Fields;
            Field oneField = collField.GetByInternalNameOrTitle(fname);
            return oneField.ReadOnlyField;
        }

        #region functions not intended to be used externally
        private static int find_lookup_field(string field_name, string[,] lookup_list_names)
        {
            int pos = -2;
            int trys = 0;
            while (trys < lookup_list_names.GetLength(0))
            {
                int pos_f = find_element3(lookup_list_names, field_name);
                if (pos_f != -1) { pos = pos_f; break; }
                else { trys++; }
            }

            return pos;
        }


        //public static bool IsFieldEditableIn2010(Field field_name)
        //{

        //    List List = field_name.ParentList;
        //    field_name.li
        //    FieldLookup fldLookup = field_name as FieldLookup;
        //    bool bCountRelated = fldLookup != null && fldLookup.CountRelated;
        //    bool bMcolLookup = fldLookup != null && fldLookup.IsDependentLookup &&
        //                                    fldLookup.LookupList != "Docs";

        //    FieldType t = field_name.Type;
        //    if (t == FieldType.Computed ||
        //        t == FieldType.File ||
        //        t == FieldType.Recurrence ||
        //        t == FieldType.CrosrojectLink ||
        //        t == FieldType.AllDayEvent)
        //    {
        //        return false;
        //    }

        //    if (!field_name.Reorderable &&
        //        !bCountRelated &&
        //        !(field_name.ReadOnlyField && field_name.Type == FieldType.User) &&
        //        !(bMcolLookup && !field_name.Hidden) &&
        //        !List.HasExternalDataSource)
        //    {
        //        return false;
        //    }


        //    if ((field_name.ReadOnlyField && !bCountRelated && !bMcolLookup) ||
        //        List.HasExternalDataSource)
        //    {
        //        if (field_name.Type == FieldType.Calculated || field_name.Type == FieldType.User)
        //            return true;

        //    }
        //    else
        //        return true;

        //    return false;
        //}
        private static int find_element3(string[,] array, string search_string)
        {
            if (search_string == null) { return -1; }
            for (int i = 0; i < array.GetLength(0); i++)
            {
                try { if ((array[i, 3] == (search_string)) || (array[i, 3].Contains(search_string) && CalcLevenshteinDistance(array[i, 3], search_string) < 3) || (search_string.Contains(array[i, 3]) && CalcLevenshteinDistance(array[i, 3], search_string) < 3)) { return i; } }
                catch (IndexOutOfRangeException) { i = i + 100000; }

            }

            return -1;
        }
        private static int find_element1(string[] array, string search_string)
        {
            for (int i = 0; i < array.GetLength(0); i++)
            {
                try { if (array[i] == (search_string) || (array[i].Contains(search_string) && CalcLevenshteinDistance(search_string, array[i]) < 2)) { return i; } }
                catch (IndexOutOfRangeException) { return -1; }
            }

            return -1;
        }
        private static int CalcLevenshteinDistance(string a, string b)
        {
            if (String.IsNullOrEmpty(a) || String.IsNullOrEmpty(b)) return 0;

            int lengthA = a.Length;
            int lengthB = b.Length;
            var distances = new int[lengthA + 1, lengthB + 1];
            for (int i = 0; i <= lengthA; distances[i, 0] = i++) ;
            for (int j = 0; j <= lengthB; distances[0, j] = j++) ;

            for (int i = 1; i <= lengthA; i++)
                for (int j = 1; j <= lengthB; j++)
                {
                    int cost = b[j - 1] == a[i - 1] ? 0 : 1;
                    distances[i, j] = Math.Min
                        (
                        Math.Min(distances[i - 1, j] + 1, distances[i, j - 1] + 1),
                        distances[i - 1, j - 1] + cost
                        );
                }
            return distances[lengthA, lengthB];
        }
        private static int find_id_pos(int[,] id_map, int search_id)
        {
            for (int i = 0; i < id_map.GetLength(0); i++)
            {

                try
                {
                    if (id_map[i, 1] == search_id)
                    {
                        return i;
                    }
                }

                catch (IndexOutOfRangeException) { IndexException(INDEX_EXCEPTION); }
                catch { }

            }
            return -1;
        }     
        private static void IndexException(string exception)
        {
            Console.WriteLine(INDEX_EXCEPTION + exception);
        }

        private static void write_to_textfile(string[,] data, string path)
        {
            string total = string.Empty;
            for (int i = 0; i < data.GetLength(0); i++)
            {
                total = total + "{\"" + data[i, 0] + "\" , \"" + data[i, 1] + "\"}," + Environment.NewLine;
            }
            System.IO.File.WriteAllText(path, total);
        }
        private static void print(string str)
        {
            Console.WriteLine(str);
        }
        private static void map_fields(string source_library, string target_library, string source_siteUrl, string target_siteUrl, string Identification_Field, string[,] lookup_fields, string filter_field, string[,] mapped_field_names)
        {
            ClientContext source_clientContext = new ClientContext(source_siteUrl);
            List oList = source_clientContext.Web.Lists.GetByTitle(source_library);
            CamlQuery camlQuery = new CamlQuery();

            if (filter_field != null)
            {
                camlQuery.ViewXml = @"<View Scope='RecursiveAll'><Query><Where>
             <Eq>
             <FieldRef Name='" + filter_field + @"' />
             <Value Type='Boolean'>1</Value>
             </Eq>
             </Where>
            </Query></View>";
            }
            ListItemCollection source_itemList = oList.GetItems(camlQuery);
            source_clientContext.Load(source_itemList);
            source_clientContext.ExecuteQuery();

            ClientContext target_clientContext = new ClientContext(target_siteUrl);
            List targ_List = target_clientContext.Web.Lists.GetByTitle(target_library);
            CamlQuery camlQuery_targ = new CamlQuery();
            ListItemCollection target_itemList = targ_List.GetItems(camlQuery_targ);
            target_clientContext.Load(target_itemList);
            target_clientContext.ExecuteQuery();
            int Ofound = -1;

            string[,] brand_titles_ids1 = get_lookupLists_values_and_ids(target_siteUrl, lookup_fields[0, 0]);///
            string[,] brand_titles_idsMS1 = get_lookupLists_values_and_ids(source_siteUrl, lookup_fields[0, 1]);///takeout ----------------
            int[,] mapped_ids1 = map_lookup_ids(brand_titles_ids1, brand_titles_idsMS1);

            for (int i = 0; i < source_itemList.Count; i++)
            {
                Ofound = -1;
                Console.WriteLine("Initiating mapping of Libraries : [" + source_library + "]" + " ---> " + "[" + target_library + "]" + Environment.NewLine);
                for (int j = 0; j < mapped_field_names.GetLength(0); j++)
                {
                    try
                    {
                        int found = find_item(Identification_Field, target_itemList, source_itemList[i][Identification_Field].ToString());
                        Ofound = found;
                        var value = source_itemList[i][mapped_field_names[j, 0]];
                        //bool editbale = IsFieldEditable(targ_List, mapped_field_names[j, 1]);
                        if (found != -1 && value != null)
                        {
                            try
                            {
                                if (value.GetType() == typeof(FieldLookupValue))
                                {
                                    int pos = find_lookup_field(mapped_field_names[j, 0], lookup_fields);
                                    if (pos > -1)
                                    {
                                        string[,] brand_titles_ids = get_lookupLists_values_and_ids(target_siteUrl, lookup_fields[pos, 0]);
                                        string[,] brand_titles_idsMS = get_lookupLists_values_and_ids(source_siteUrl, lookup_fields[pos, 1]);
                                        int[,] mapped_ids = map_lookup_ids(brand_titles_ids, brand_titles_idsMS);

                                        var source = (FieldLookupValue)value;
                                        int id = get_id(source);

                                        int pos1 = find_id_pos(mapped_ids, id);

                                        if (pos1 != -1)
                                        {
                                            source.LookupId = mapped_ids[pos1, 0];
                                            target_itemList[found][mapped_field_names[j, 1]] = source;
                                        }
                                    }
                                }
                                else if (value.GetType() == typeof(FieldText)) { var source = (FieldText)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldUrlValue)) { var source = (FieldUrlValue)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldUserValue)) { var source = (FieldUserValue)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldChoice)) { var source = (FieldChoice)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldMultiChoice)) { var source = (FieldMultiChoice)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldDateTime)) { var source = (FieldDateTime)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldNumber)) { var source = (FieldNumber)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldCollection)) { var source = (FieldCollection)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldCalculated)) { var source = (FieldCalculated)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldComputed)) { var source = (FieldComputed)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldCurrency)) { var source = (FieldCurrency)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(FieldLink)) { var source = (FieldLink)value; target_itemList[found][mapped_field_names[j, 1]] = source; }
                                else if (value.GetType() == typeof(String)) { var source = value; target_itemList[found][mapped_field_names[j, 1]] = source; }

                                target_itemList[found].Update();
                                // target_clientContext.ExecuteQuery();
                                //  Console.WriteLine("Copied [Field: " + mapped_field_names[j, 1] + "] of [item: " + source_itemList[i][Identification_Field] + "] from libraries: " + source_library + "-->" + target_library);
                                //  Console.WriteLine();
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.ToString());
                            }
                        }
                    }
                    catch (Microsoft.SharePoint.Client.PropertyOrFieldNotInitializedException) { }
                }
                try
                {
                    if (Ofound != -1)
                    {
                        target_clientContext.ExecuteQuery();
                        Console.WriteLine("Successfully mapped fields of item : [" + source_itemList[i]["FileLeafRef"].ToString() + "]" + Environment.NewLine);
                    }
                }
                catch { Console.WriteLine("Error! - Failing to map some fields!"); }
            }

            target_clientContext.Dispose();
            source_clientContext.Dispose();
        }
        private static int get_id(FieldLookupValue x)
        {
            var y = (FieldLookupValue)x;
            return y.LookupId;
        }
        private static bool Empty_check(string str) { if (str == null || str == string.Empty || str.Trim() == null || str.Trim() == string.Empty) { return true; } else { return false; } }
        private static dynamic get_item(ListItemCollection c, int i)
        {
            int p = 0;
            foreach (ListItem m in c) { if (p == i) { return m; } }
            return null;
        }
        private static int find_item(string fname, ListItemCollection source_collist, string criteria)
        {

            for (int i = 0; i < source_collist.Count; i++)
            {
                var thing = source_collist[i][fname];
                if (thing != null)
                {
                    if (thing.ToString() == criteria) { return i; }
                }
            }
            return -1;

        }
        private static string[,] get_field_names(string siteUrl, string source_library)
        {

            ClientContext context = new ClientContext(siteUrl);
            Web web = context.Web;
            List list = (List)web.Lists.GetByTitle(source_library);


            context.Load(list.Fields);
            context.ExecuteQuery();


            string[,] array = new string[list.Fields.Count, 2];

            int n = list.Fields.Count;

            for (int i = 0; i < n; i++)
            {
                try
                {
                    array[i, 0] = list.Fields[i].Title.ToString();
                    array[i, 1] = list.Fields[i].InternalName.ToString();
                }
                catch (IndexOutOfRangeException) { i = i + 100000; }
                catch { }
            }

            return array;

        }
        private static string[,] get_mapped_field_names(string source_url, string source_library, string targ_url, string targ_library)
        {
            string[,] source = get_field_names(source_url, source_library);
            string[,] target = get_field_names(targ_url, targ_library);
            int length = source.GetLength(0);
            if (target.GetLength(0) < length) { length = target.GetLength(0); }
            int n_fields = 0;
            for (int i = 0; i < length; i++)
            {
                try
                {
                    if (Empty_check(source[i, 1]) == false && Empty_check(target[i, 1]) == false)
                    {
                        n_fields++;
                    }
                }
                catch (IndexOutOfRangeException) { i = i + 100000; }
            }

            string[,] mapped_fields = new string[n_fields, 2];//source, target

            for (int u = 0; u < n_fields; u++)
            {
                int a = find_element(target, source[u, 0]);
                if (a != -1)
                {
                    mapped_fields[u, 0] = source[u, 1];
                    mapped_fields[u, 1] = target[a, 1];
                }
            }
            return remove_nulls(mapped_fields);
        }
        private static int find_element(string[,] array, string search_string)
        {
            if (search_string == null) { return -1; }
            for (int i = 0; i < array.GetLength(0); i++)
            {
                try { if ((array[i, 0] == (search_string)) || (array[i, 0].Contains(search_string) && CalcLevenshteinDistance(array[i, 0], search_string) < 3) || (search_string.Contains(array[i, 0]) && CalcLevenshteinDistance(array[i, 0], search_string) < 3)) { return i; } }
                catch (IndexOutOfRangeException) { i = i + 100000; }

            }

            return -1;
        }
        private static string[,] remove_nulls(string[,] array)
        {
            int length = 0;
            for (int i = 0; i < array.GetLength(0) - 1; i++)
            {
                try
                {
                    if (Empty_check(array[i, 0]) == false && Empty_check(array[i, 1]) == false)
                    {
                        length++;
                    }
                }
                catch (IndexOutOfRangeException) { i = i + 100000; }
            }

            string[,] refined_array = new string[length, 2];
            int succ = 0;

            for (int u = 0; u < array.GetLength(0); u++)
            {
                try
                {
                    if (Empty_check(array[u, 0]) == false && Empty_check(array[u, 1]) == false)
                    {
                        refined_array[succ, 0] = array[u, 0];
                        refined_array[succ, 1] = array[u, 1];
                        succ++;

                    }
                }
                catch (IndexOutOfRangeException) { u = u + 100000; }
            }

            return refined_array;

        }
        private static string[,] get_lookupLists_values_and_ids(string siteUrl, string field)
        {
            ClientContext clientContext = new ClientContext(siteUrl);
            List lookupFieldList = clientContext.Web.Lists.GetByTitle(field);
            clientContext.Load(lookupFieldList);
            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = lookupFieldList.GetItems(camlQuery);
            clientContext.Load(collListItem);

            List<string> titles = new List<string>();
            List<string> ids = new List<string>();

            clientContext.ExecuteQuery();

            foreach (ListItem item in collListItem)
            {
                if (item != null)
                {
                    titles.Add(item.FieldValues["Title"].ToString());
                    ids.Add(item.FieldValues["ID"].ToString());
                    //Console.WriteLine(item.FieldValues["Title"] + "---" + item.FieldValues["ID"]);
                }

            }
            string[] titlesarray = titles.ToArray();
            string[] idsarray = ids.ToArray();

            string[,] titles_ids = new string[titles.Count, 2];

            for (int i = 0; i < titlesarray.Length; i++)
            {

                titles_ids[i, 0] = titlesarray[i];
                titles_ids[i, 1] = idsarray[i];
            }

            clientContext.Dispose();
            return titles_ids;
        }
        private static int[,] map_lookup_ids(string[,] brand_titles_ids, string[,] brand_titles_idsMS)
        {
            System.Collections.Generic.List<int> first = new System.Collections.Generic.List<int>();
            System.Collections.Generic.List<int> second = new System.Collections.Generic.List<int>();
            int len = 0;
            if (brand_titles_idsMS.GetLength(0) > brand_titles_ids.GetLength(0)) { len = brand_titles_idsMS.GetLength(0); } else { len = brand_titles_ids.GetLength(0); }

            for (int i = 0; i < len - 1; i++)
            {
                try
                {
                    int found = find_element(brand_titles_idsMS, brand_titles_ids[i, 0]);
                    if (found != -1)
                    {
                        first.Add(Convert.ToInt32(brand_titles_ids[i, 1]));
                        second.Add(Convert.ToInt32(brand_titles_idsMS[found, 1]));
                    }
                }
                catch (IndexOutOfRangeException) { i = i + 1000000; }
            }
            int[] firstarray = first.ToArray();
            int[] secondarray = second.ToArray();
            int[,] mapped_array = new int[first.Count, 2];
            for (int i = 0; i < first.Count; i++)
            {
                mapped_array[i, 0] = firstarray[i];
                mapped_array[i, 1] = secondarray[i];
            }

            return mapped_array;

        }
        private static string[,] clean_array(string[,] array)
        {
            int len = 0;
            System.Collections.Generic.List<string> list = new System.Collections.Generic.List<string>(); System.Collections.Generic.List<string> list2 = new System.Collections.Generic.List<string>();
            for (int i = 0; i < array.GetLength(0); i++)
            {
                if (Empty_check(array[i, 0]) == false && Empty_check(array[i, 1]) == false)
                {
                    list.Add(array[i, 0]);
                    list2.Add(array[i, 1]);
                    len++;
                }
            }

            string[] list_array = list.ToArray();
            string[] list_array2 = list2.ToArray();

            string[,] new_array = new string[len, 2];

            for (int u = 0; u < len; u++)
            {
                new_array[u, 0] = list_array[u];
                new_array[u, 1] = list_array2[u];
            }
            return new_array;
        }
        #endregion functions not intended to be used externally
    }
}
