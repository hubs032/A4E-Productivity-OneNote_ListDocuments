using UiPath.CodedWorkflows;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;


namespace AFE_OneNote_ListDocumnets{
    public class GetOneNoteStructure : CodedWorkflow
    {
        
        static Application onenoteApp = new Application();
        static XNamespace ns = null;
        
        static void GetNamespace() {
            
            string xml;

            onenoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);
            var doc = XDocument.Parse(xml);
            ns = doc.Root.Name.Namespace;
        }
        
        
        static List<XElement> GetObjects(string parentId, HierarchyScope scope) {
            string xml;
            onenoteApp.GetHierarchy(parentId, scope, out xml);
    
            var doc = XDocument.Parse(xml);
            var nodeName = "";
    
            switch (scope) {
                case (HierarchyScope.hsNotebooks): nodeName = "Notebook"; break;
                case (HierarchyScope.hsPages): nodeName = "Page"; break;
                case (HierarchyScope.hsSections): nodeName = "Section"; break;
                default:
                    return null;
            }
            var nodes = doc.Descendants(ns + nodeName);
            
            return nodes.ToList();

        }
        
        [Workflow]
        public JArray Execute()
        {
            GetNamespace();
           
            string sectionId, notebookId;
            var notebooks = GetObjects(null, HierarchyScope.hsNotebooks);
            JArray JResult = new JArray();
            
            foreach (var notebook in notebooks) {
                
                notebookId = notebook.Attribute("ID").Value;
                var notebookName = notebook.Attribute("name").Value;
                var sections = GetObjects(notebookId, HierarchyScope.hsSections);
                JArray JSections = new JArray();
                
                foreach (var section in sections) { 
                    
                    JArray JPages = new JArray();                   
                    var sectionName = section.Attribute("name").Value;
                    sectionId = section.Attribute("ID").Value;
                    JSections.Add(new JObject{{"SectionID", sectionId}, {"Section", sectionName}});
                    var pages = GetObjects(sectionId, HierarchyScope.hsPages);
 
                    foreach (var page in pages) { 
                        JPages.Add(new JObject{{"PageID",page.Attribute("ID").Value},{"Page", page.Attribute("name").Value}});
                    }
                    JObject JSection = new JObject{{"SectionID", sectionId}, {"Section", sectionName}, {"Pages", JPages}};               
                    
                    JSections.Add(JSection);
                }
                JResult.Add(new JObject{{"NotebookID", notebookId}, {"Notebook", notebookName}, {"Sections", JSections}});
            }
            //Console.WriteLine(JResult);
            return JResult;
        }
    }
}