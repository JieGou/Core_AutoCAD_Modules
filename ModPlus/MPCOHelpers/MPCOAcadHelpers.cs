using System;
#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ModPlus.MPCOHelpers
{
    public static class AcadHelpers
    {
        //public static DocData DocData = new DocData();
        //public static bool TransientNeedClear { get; set; }
        /// <summary>
        /// БД активного документа
        /// </summary>
        public static Database Database => AcApp.DocumentManager.MdiActiveDocument.Database;
        /// <summary>
        /// Активный документ
        /// </summary>
        public static Document Document => AcApp.DocumentManager.MdiActiveDocument;
        /// <summary>
        /// Редактор активного документа
        /// </summary>
        public static Editor Editor => AcApp.DocumentManager.MdiActiveDocument.Editor;

        /// <summary>
        /// Открыть объект для чтения
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="objectId"></param>
        /// <param name="openErased"></param>
        /// <param name="forceOpenOnLockedLayer"></param>
        /// <returns></returns>
        public static T Read<T>(this ObjectId objectId, bool openErased = false, bool forceOpenOnLockedLayer = true)
            where T : DBObject
        {
            return (T)(objectId.GetObject(0, openErased, forceOpenOnLockedLayer) as T);
        }
        /// <summary>
        /// Открыть объект для записи
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="objectId"></param>
        /// <param name="openErased"></param>
        /// <param name="forceOpenOnLockedLayer"></param>
        /// <returns></returns>
        public static T Write<T>(this ObjectId objectId, bool openErased = false, bool forceOpenOnLockedLayer = true)
            where T : DBObject
        {
            return (T)(objectId.GetObject(OpenMode.ForWrite, openErased, forceOpenOnLockedLayer) as T);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="point"></param>
        /// <param name="enities"></param>
        /// <returns></returns>
        public static ObjectId AddBlock(Point3d point, params Entity[] enities)
        {
            ObjectId objectId;
            BlockTableRecord blockTableRecord = new BlockTableRecord
            {
                Name = "*U"
            };

            Entity[] entityArray = enities;
            for (int i = 0; i < (int)entityArray.Length; i++)
            {
                blockTableRecord.AppendEntity(entityArray[i]);
            }
            using (Document.LockDocument())
            {
                using (Transaction tr = Database.TransactionManager.StartTransaction())
                {
                    using (BlockTable blockTable = Database.BlockTableId.Write<BlockTable>(false, true))
                    {
                        using (BlockReference blockReference = new BlockReference(point, blockTable.Add(blockTableRecord)))
                        {
                            ObjectId item = blockTable[BlockTableRecord.ModelSpace];//&&&&&&&&?????????????????????????????????????????????????paperspace
                            using (BlockTableRecord btr = item.Write<BlockTableRecord>(false, true))
                            {
                                objectId = btr.AppendEntity(blockReference);
                            }
                            tr.AddNewlyCreatedDBObject(blockReference, true);
                        }
                        tr.AddNewlyCreatedDBObject(blockTableRecord, true);
                    }
                    tr.Commit();
                }
            }
            return objectId;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="point"></param>
        /// <param name="enities"></param>
        /// <returns></returns>
        public static BlockReference GetBlockReference(Point3d point, IEnumerable<Entity> enities)
        {
            BlockTableRecord blockTableRecord = new BlockTableRecord
            {
                Name = "*U"
            };
            foreach (Entity enity in enities)
            {
                blockTableRecord.AppendEntity(enity);
            }
            return new BlockReference(point, blockTableRecord.ObjectId);
        }
    }
    /// <summary>
    /// Вспомогательные методы работы с расширенными данными
    /// Есть аналогичные в MpCadHelpers. Некоторые будут совпадать
    /// но все-равно делаю отдельно
    /// </summary>
    public static class ExtendedDataHelpers
    {
        /// <summary>
        /// Добавление регистрации приложения в соответсвующую таблицу чертежа
        /// </summary>
        public static void AddRegAppTableRecord(string appName)
        {
            using (var tr = AcadHelpers.Document.TransactionManager.StartTransaction())
            {
                RegAppTable rat = (RegAppTable)tr.GetObject(AcadHelpers.Database.RegAppTableId, OpenMode.ForRead, false);
                if (!rat.Has(appName))
                {
                    rat.UpgradeOpen();
                    RegAppTableRecord ratr = new RegAppTableRecord
                    {
                        Name = appName
                    };
                    rat.Add(ratr);
                    tr.AddNewlyCreatedDBObject(ratr, true);
                }
                tr.Commit();
            }
        }

        /// <summary>
        /// Проверка поддерживаемости примитива для Overrule
        /// </summary>
        /// <param name="rxObject"></param>
        /// <returns></returns>
        public static bool IsApplicable(RXObject rxObject, string appName)
        {
            DBObject dbObject = rxObject as DBObject;
            if (dbObject == null) return false;
            return IsMPCOentity(dbObject, appName);
        }
        ///// <summary>
        ///// Проверка по XData вхождения блока, что он является конкретным СПДС примитивом
        ///// </summary>
        ///// <param name="SPDSent">Имя примитива</param>
        ///// <param name="blkRef">Вхождение блока</param>
        ///// <returns></returns>
        //public static bool IsSPDSentity(string SPDSent, Entity blkRef)
        //{
        //    ResultBuffer rb = blkRef.GetXDataForApplication(SDPSInitialize.AppName);
        //    if (rb != null)
        //    {
        //        TypedValue[] rvArr = rb.AsArray();
        //        foreach (TypedValue typedValue in rvArr)
        //        {
        //            if ((DxfCode)typedValue.TypeCode == DxfCode.ExtendedDataAsciiString) // 1000
        //                if (((string)typedValue.Value).Equals(SPDSent))
        //                    return true;
        //        }
        //    }
        //    return false;
        //}
        //public static bool IsSPDSentity(string SPDSent, DBObject dbObject)
        //{
        //    ResultBuffer rb = dbObject.GetXDataForApplication(SDPSInitialize.AppName);
        //    if (rb != null)
        //    {
        //        TypedValue[] rvArr = rb.AsArray();
        //        foreach (TypedValue typedValue in rvArr)
        //        {
        //            if ((DxfCode)typedValue.TypeCode == DxfCode.ExtendedDataAsciiString) // 1000
        //                if (((string)typedValue.Value).Equals(SPDSent))
        //                    return true;
        //        }
        //    }
        //    return false;
        //}
        /// <summary>
        /// Проверка по XData вхождения блока, что он является любым СПДС примитивом
        /// </summary>
        /// <param name="blkRef">Вхождение блока</param>
        /// <returns></returns>
        public static bool IsMPCOentity(Entity blkRef, string appName)
        {
            ResultBuffer rb = blkRef.GetXDataForApplication(appName);
            return rb != null;
        }
        public static bool IsMPCOentity(DBObject dbObject, string appName)
        {
            ResultBuffer rb = dbObject.GetXDataForApplication(appName);
            return rb != null;
        }
        ///// <summary>
        ///// Добавление "метки" с названием СПДС примитива в XData вхождения блока
        ///// </summary>
        ///// <param name="SPDSent">Имя примитива</param>
        ///// <param name="blkRef">Вхождение блока</param>
        //public static void SetSPDSentityNameToXData(string SPDSent, BlockReference blkRef)
        //{
        //    // Блок уже открыт на запись!!!
        //    blkRef.XData = new ResultBuffer
        //    {
        //        new TypedValue((int)DxfCode.ExtendedDataRegAppName, SDPSInitialize.AppName), // 1001
        //        new TypedValue((int)DxfCode.ExtendedDataAsciiString, SPDSent) // 1000
        //    };
        //}
        ///// <summary>
        ///// Добавление (обновление) Extension Dictionary Data к объекту
        ///// </summary>
        ///// <param name="buf"></param>
        ///// <param name="objectId"></param>
        ///// <param name="dataKey"></param>
        //public static void AddEDD(ResultBuffer buf, ObjectId objectId, string dataKey)
        //{
        //    SymbolUtilityServices.ValidateSymbolName(dataKey, false); //validates the name is acceptable for the Dictionary, if not throws exception
        //    using (AcadHelpers.Document.LockDocument())
        //    using (var tr = AcadHelpers.Document.TransactionManager.StartTransaction())
        //    {
        //        DBObject dbObject = (DBObject)tr.GetObject(objectId, OpenMode.ForWrite);
        //        if (dbObject.ExtensionDictionary.IsNull)
        //            dbObject.CreateExtensionDictionary();
        //        DBDictionary extDictionary = (DBDictionary)tr.GetObject(dbObject.ExtensionDictionary, OpenMode.ForWrite);
        //        if (extDictionary.Contains(dataKey))
        //        {
        //            ObjectId xRecId = extDictionary.GetAt(dataKey);
        //            Xrecord xrecord = (Xrecord)tr.GetObject(xRecId, OpenMode.ForWrite);
        //            xrecord.Data = buf;
        //        }
        //        else
        //        {
        //            Xrecord xrecord = new Xrecord();
        //            xrecord.Data = buf;
        //            extDictionary.SetAt(dataKey, xrecord);
        //            tr.AddNewlyCreatedDBObject(xrecord, true);
        //        }
        //        tr.Commit();
        //    }
        //}
        ///// <summary>
        ///// Get Extension Dictionary Data from object by dataKey
        ///// </summary>
        ///// <param name="objectId"></param>
        ///// <param name="dataKey"></param>
        ///// <returns></returns>
        //public static ResultBuffer GetEDD(ObjectId objectId, string dataKey)
        //{
        //    SymbolUtilityServices.ValidateSymbolName(dataKey, false); //validates the name is acceptable for the Dictionary, if not throws exception
        //    using (AcadHelpers.Document.LockDocument())
        //    using (var tr = AcadHelpers.Document.TransactionManager.StartTransaction())
        //    {
        //        DBObject dbObject = (DBObject)tr.GetObject(objectId, OpenMode.ForRead);
        //        if (dbObject.ExtensionDictionary.IsNull) return null;
        //        DBDictionary extDictionary = (DBDictionary)tr.GetObject(dbObject.ExtensionDictionary, OpenMode.ForRead);
        //        if (extDictionary.Contains(dataKey))
        //        {
        //            ObjectId xRecId = extDictionary.GetAt(dataKey);
        //            Xrecord xrecord = (Xrecord)tr.GetObject(xRecId, OpenMode.ForWrite);
        //            return xrecord.Data;
        //        }
        //        return null;
        //    }
        //}
    }

    //public class DocData
    //{
    //    private Dictionary<Document, Dictionary<string, object>> data = new Dictionary<Document, Dictionary<string, object>>();

    //    public Dictionary<string, object> Data
    //    {
    //        get
    //        {
    //            if (!data.ContainsKey(AcadHelpers.Document))
    //                Add(AcadHelpers.Document);
    //            return data[AcadHelpers.Document];
    //        }
    //        set
    //        {
    //            if (!data.ContainsKey(AcadHelpers.Document))
    //            {
    //                Add(AcadHelpers.Document); return;
    //            }
    //            data[AcadHelpers.Document] = value;
    //        }
    //    }
    //    public DocData()
    //    {
    //        Initialise();
    //    }

    //    public DocData(Document doc)
    //    {
    //        Initialise();
    //        Add(doc);
    //    }
    //    private void Initialise()
    //    {
    //        AcApp.DocumentManager.DocumentCreated += (this.DocumentManager_DocumentCreated);
    //        AcApp.DocumentManager.DocumentToBeDestroyed += (this.DocumentManager_DocumentToBeDestroyed);
    //    }
    //    private void DocumentManager_DocumentToBeDestroyed(object sender, DocumentCollectionEventArgs e)
    //    {
    //        data.Remove(e.Document);
    //    }

    //    private void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
    //    {
    //        Add(e.Document);
    //    }
    //    public void Add(Document doc)
    //    {
    //        if (!data.ContainsKey(doc))
    //        {
    //            data.Add(doc, new Dictionary<string, object>());
    //            //Dictionary<string, string> dictionary = new Dictionary<string, string>();
    //            //Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
    //            //using (Transaction tr = AcadHelpers.Document.TransactionManager.StartTransaction())
    //            //{
    //            //    TextStyleTable textStyleTable = (TextStyleTable)tr.GetObject(Application.DocumentManager.MdiActiveDocument.Database.TextStyleTableId, OpenMode.ForRead);
    //            //    using (SymbolTableEnumerator enumerator = textStyleTable.GetEnumerator())
    //            //    {
    //            //        while (enumerator.MoveNext())
    //            //        {
    //            //            ObjectId current = enumerator.Current;
    //            //            TextStyleTableRecord textStyleTableRecord = (TextStyleTableRecord)tr.GetObject(current, OpenMode.ForRead);
    //            //            dictionary.Add(textStyleTableRecord.Id.ToString().Replace("(", "").Replace(")", ""), textStyleTableRecord.Name);
    //            //        }
    //            //    }
    //            //    DBDictionary dBDictionary = (DBDictionary)tr.GetObject(AcadHelpers.Database.MLeaderStyleDictionaryId, OpenMode.ForRead);
    //            //    using (DbDictionaryEnumerator enumerator2 = dBDictionary.GetEnumerator())
    //            //    {
    //            //        while (enumerator2.MoveNext())
    //            //        {
    //            //            DBDictionaryEntry current2 = enumerator2.Current;
    //            //            dictionary2.Add(current2.Value.ToString().Replace("(", "").Replace(")", ""), current2.Key);
    //            //        }
    //            //    }
    //            //}
    //            //data[doc].Add("TEXT_STYLES", dictionary);
    //            //data[doc].Add("ARROW_STYLES", dictionary2);
    //        }
    //    }

    //    public void AddDataObject(string key, object dataObject)
    //    {
    //        if (!data[AcApp.DocumentManager.MdiActiveDocument].Keys.Contains(key))
    //        {
    //            data[AcApp.DocumentManager.MdiActiveDocument].Add(key, dataObject);
    //        }
    //    }
    //}
}
