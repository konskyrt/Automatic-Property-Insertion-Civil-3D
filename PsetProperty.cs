using Autodesk.Aec.PropertyData.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Properties_Update
{
    public static class PsetProperty
    {
        /// <summary>
        ///Create P set prop
        /// </summary>
        public static void CreatePSetTab(Dictionary<string, string> keyValuePairs, string signal)
        {

            Document document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Transaction trans = document.Database.TransactionManager.StartTransaction())
            {
                DocumentLock LckDoc = document.LockDocument();
                Database db = document.Database;
                PropertySetDefinition propSetDef = new PropertySetDefinition();
                propSetDef.SetToStandard(db);
                propSetDef.SubSetDatabaseDefaults(db);
                propSetDef.Description = signal;
                foreach (var item in keyValuePairs)
                {
                    PropertyDefinition propDefAutomatic = CreatePrpdef(db, item.Key, item.Value);
                    propSetDef.Definitions.Add(propDefAutomatic);
                }

                DictionaryPropertySetDefinitions dictPropSetDef = new DictionaryPropertySetDefinitions(db);
                if (!dictPropSetDef.Has(signal, trans))
                {
                    dictPropSetDef.AddNewRecord(signal, propSetDef);
                    trans.AddNewlyCreatedDBObject(propSetDef, true);
                    trans.Commit();
                }
            }
            PropertySetattachtoEntity(document, signal);

        }

        /// <summary>
        ///Set P set prop to Blocks
        /// </summary>
        private static void PropertySetattachtoEntity(Document document, string signal)
        {
            Dictionary<ObjectId, string> blocks = GetAlltheBlocks(document);

            using (Transaction trans = document.Database.TransactionManager.StartTransaction())
            {
                foreach (ObjectId blk in blocks.Keys)
                {
                    if (blk != null)
                    {
                        DictionaryPropertySetDefinitions dictPropSetDef1 = new DictionaryPropertySetDefinitions(document.Database);
                        ObjectId objectId = dictPropSetDef1.GetAt(signal);//Utils.findStyle(dictPropSetDef, "Property Set By Automation");
                        Entity ent = (Entity)trans.GetObject(blk, OpenMode.ForWrite, false);
                        if (signal == blocks[blk])
                            PropertyDataServices.AddPropertySet(ent, objectId);
                    }

                }
                // Save the new object to the database
                trans.Commit();
            }
        }
        /// <summary>
        ///Get all the blocks from Doc
        /// </summary>
        private static Dictionary<ObjectId, string> GetAlltheBlocks(Document document)
        {
            Dictionary<ObjectId, string> blocks = new Dictionary<ObjectId, string>();
            try
            {
                Database database = document.Database;
                using (Transaction transaction = database.TransactionManager.StartTransaction())
                {
                    BlockTable blockTable = transaction.GetObject(database.BlockTableId, OpenMode.ForWrite) as BlockTable;
                    ObjectId modelId = blockTable[BlockTableRecord.ModelSpace];
                    BlockTableRecord model = transaction.GetObject(modelId, OpenMode.ForWrite) as BlockTableRecord;
                    foreach (ObjectId id in model)
                    {
                        if (id.IsNull || id.IsErased || id.IsEffectivelyErased || !id.IsValid)
                            continue;

                        Entity ent = (Entity)transaction.GetObject(id, OpenMode.ForWrite, false);
                        if (ent.GetType().ToString() == "Autodesk.AutoCAD.DatabaseServices.BlockReference")
                        {
                            BlockReference blockReference = transaction.GetObject(id, OpenMode.ForWrite) as BlockReference;
                            blocks.Add(blockReference.ObjectId, blockReference.Name);
                        }
                    }
                }
            }
            catch { }

            return blocks;
        }
        /// <summary>
        ///Set P set prop Def
        /// </summary>
        private static PropertyDefinition CreatePrpdef(Database db, string PropName, string PropVal)
        {
            PropertyDefinition propDefAutomatic = new PropertyDefinition();
            propDefAutomatic.SetToStandard(db);
            propDefAutomatic.SubSetDatabaseDefaults(db);
            propDefAutomatic.Name = PropName;
            propDefAutomatic.DataType = Autodesk.Aec.PropertyData.DataType.Text;
            propDefAutomatic.DefaultData = PropVal;
            return propDefAutomatic;
        }
    }
}
