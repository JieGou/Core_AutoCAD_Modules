using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using mpMsg;
using Autodesk.AutoCAD.ApplicationServices;

namespace ModPlus.MPCOHelpers
{
    public static class MPCOEntityHelpers
    {
        /// <summary>
        /// Список СПДС примитивов, входящих в плагин ModPlus
        /// </summary>
        public static List<string> SupportedSPDSEntities = new List<string>
        {
            "mpBreakLine"
        };
    }
    public abstract class MPCOEntity 
    {
        protected MPCOEntity()
        {
            BlockTransform = Matrix3d.Identity;
        }
        /// <summary>
        /// Первая точка примитива в мировой системе координат.
        /// Должна соответствовать точке вставке блока
        /// </summary>
        public Point3d InsertionPoint { get; set; } = Point3d.Origin;
        /// <summary>
        /// Коллекция базовых примитивов, входящих в примитив 
        /// </summary>
        public abstract IEnumerable<Entity> Entities
        {
            get;
        }

        public bool IsValueCreated { get; set; }
        /// <summary>
        /// Матрица трансформации BlockReference
        /// </summary>
        public Matrix3d BlockTransform { get; set; }
        // Значение аннотативности
        public bool _annotative;
        public bool Annotative
        {
            get { return _annotative; }
            set
            {
                _annotative = value;
                using (AcadHelpers.Document.LockDocument(DocumentLockMode.ProtectedAutoWrite, null, null, true))
                {
                    using (var tr = AcadHelpers.Document.TransactionManager.StartTransaction())
                    {
                        //var obj = (BlockTableRecord)tr.GetObject(BlockRecord.Id, OpenMode.ForWrite);
                        var obj = (BlockReference)tr.GetObject(BlockId, OpenMode.ForWrite);
                        obj.Annotative = !_annotative ? AnnotativeStates.False : AnnotativeStates.True;
                        //var blockReference = (BlockReference)tr.GetObject(Id, OpenMode.ForWrite);
                        var blockReference = (BlockReference)tr.GetObject(BlockId, OpenMode.ForWrite);
                        var contextCollection = AcadHelpers.Database.ObjectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES");
                        if (!blockReference.HasContext(contextCollection.GetContext("1:1")))
                        {
                            blockReference.AddContext(contextCollection.GetContext("1:1"));
                        }
                        if (!blockReference.HasContext(AcadHelpers.Database.Cannoscale))
                        {
                            blockReference.AddContext(AcadHelpers.Database.Cannoscale);
                        }
                        tr.Commit();
                    }
                }
            }
        }
        #region Block
        // ObjectId "примитива"
        public ObjectId BlockId { get; set; }

        // Описание блока
        private BlockTableRecord _blockRecord;
        public BlockTableRecord BlockRecord
        {
            get
            {
                try
                {
                    if (!BlockId.IsNull)
                    {
                        using (AcadHelpers.Document.LockDocument())
                        {
                            using (var tr = AcadHelpers.Database.TransactionManager.StartTransaction())
                            {
                                var blkRef = (BlockReference)tr.GetObject(BlockId, OpenMode.ForWrite);
                                _blockRecord = (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForWrite);
                                if (_blockRecord.GetBlockReferenceIds(true, true).Count <= 1)
                                {
                                    foreach (var objectId in _blockRecord)
                                    {
                                        tr.GetObject(objectId, OpenMode.ForWrite).Erase();
                                    }
                                }
                                else
                                {
                                    _blockRecord = new BlockTableRecord { Name = "*U", BlockScaling = BlockScaling.Uniform };
                                    using (var blockTable =
                                        AcadHelpers.Database.BlockTableId.Write<BlockTable>())
                                    {
                                        if (Annotative)
                                        {
                                            _blockRecord.Annotative = AnnotativeStates.True;
                                        }
                                        blockTable.Add(_blockRecord);
                                        tr.AddNewlyCreatedDBObject(_blockRecord, true);
                                    }
                                    blkRef.BlockTableRecord = _blockRecord.Id;
                                }

                                tr.Commit();
                            }
                            using (var tr = AcadHelpers.Database.TransactionManager.StartTransaction())
                            {
                                var blkRef = (BlockReference)tr.GetObject(BlockId, OpenMode.ForWrite);
                                _blockRecord = (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForWrite);
                                _blockRecord.BlockScaling = BlockScaling.Uniform;
                                var matrix3D = Matrix3d.Displacement(-InsertionPoint.TransformBy(BlockTransform.Inverse()).GetAsVector());
                                foreach (var entity in Entities)
                                {
                                    var transformedCopy = entity.GetTransformedCopy(matrix3D);
                                    _blockRecord.AppendEntity(transformedCopy);
                                    tr.AddNewlyCreatedDBObject(transformedCopy, true);
                                }


                                tr.Commit();
                            }
                        }
                    }
                    else if (!IsValueCreated)
                    {
                        var matrix3D = Matrix3d.Displacement(-InsertionPoint.TransformBy(BlockTransform.Inverse()).GetAsVector());
                        foreach (var ent in Entities)
                        {
                            var transformedCopy = ent.GetTransformedCopy(matrix3D);
                            _blockRecord.AppendEntity(transformedCopy);
                        }
                        IsValueCreated = true;
                    }
                }
                catch (Exception exception)
                {
                    MpExWin.Show(exception);
                }
                return _blockRecord;
            }
            set => _blockRecord = value;
        }
        
        public BlockTableRecord GetBlockTableRecordForUndo(BlockReference blkRef)
        {
            BlockTableRecord blockTableRecord;
            using (AcadHelpers.Document.LockDocument())
            {
                using (var tr = AcadHelpers.Database.TransactionManager.StartTransaction())
                {
                    blockTableRecord = new BlockTableRecord { Name = "*U", BlockScaling = BlockScaling.Uniform };
                    using (var blockTable = AcadHelpers.Database.BlockTableId.Write<BlockTable>())
                    {
                        if (Annotative)
                        {
                            blockTableRecord.Annotative = AnnotativeStates.True;
                        }
                        blockTable.Add(blockTableRecord);
                        tr.AddNewlyCreatedDBObject(blockTableRecord, true);
                    }
                    blkRef.BlockTableRecord = blockTableRecord.Id;
                    tr.Commit();
                }
                using (var tr = AcadHelpers.Database.TransactionManager.StartOpenCloseTransaction())
                {
                    blockTableRecord = (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForWrite);
                    blockTableRecord.BlockScaling = BlockScaling.Uniform;
                    var matrix3D = Matrix3d.Displacement(-InsertionPoint.TransformBy(BlockTransform.Inverse()).GetAsVector());
                    foreach (var entity in Entities)
                    {
                        var transformedCopy = entity.GetTransformedCopy(matrix3D);
                        blockTableRecord.AppendEntity(transformedCopy);
                        tr.AddNewlyCreatedDBObject(transformedCopy, true);
                    }
                    tr.Commit();
                }
            }
            _blockRecord = blockTableRecord;
            return blockTableRecord;
        }
        #endregion
        /// <summary>
        /// Получение свойств блока, которые присуще примитиву
        /// </summary>
        /// <param name="entity"></param>
        public void GetParametersFromEntity(Entity entity)
        {
            var blockReference = (BlockReference)entity;
            if (blockReference != null)
            {
                InsertionPoint = blockReference.Position;
                BlockTransform = blockReference.BlockTransform;
            }
        }

        public void Draw(WorldDraw draw)
        {
            var geometry = draw.Geometry;
            foreach (var entity in Entities)
            {
                geometry.Draw(entity);
            }
        }

        public void Erase()
        {
            foreach (var entity in Entities)
            {
                entity.Erase();
            }
        }
        /// <summary>
        /// Расчлинение СПДС примитива
        /// </summary>
        /// <param name="entitySet"></param>
        public void Explode(DBObjectCollection entitySet)
        {
            entitySet.Clear();
            foreach (var entity in Entities)
            {
                entitySet.Add(entity);
            }
        }
        
    }
    public sealed class MyBinder : SerializationBinder
    {
        public override Type BindToType(
          string assemblyName,
          string typeName)
        {
            return Type.GetType($"{typeName}, {assemblyName}");
        }
    }
}
