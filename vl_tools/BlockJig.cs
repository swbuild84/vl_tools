using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;

namespace vl_tools
{
    class AttInfo
    {
        private Point3d _pos;
        private Point3d _aln;
        private bool _aligned;

        public AttInfo(Point3d pos, Point3d aln, bool aligned)
        {
            _pos = pos;
            _aln = aln;
            _aligned = aligned;
        }

        public Point3d Position
        {
            set { _pos = value; }
            get { return _pos; }
        }

        public Point3d Alignment
        {
            set { _aln = value; }
            get { return _aln; }
        }

        public bool IsAligned
        {
            set { _aligned = value; }
            get { return _aligned; }
        }
    }

    class BlockJig : EntityJig
    {
        private Point3d _pos;
        private Dictionary<ObjectId, AttInfo> _attInfo;
        private Transaction _tr;

        public BlockJig(Transaction tr, BlockReference br, Dictionary<ObjectId, AttInfo> attInfo)
            : base(br)
        {
            _pos = br.Position;
            _attInfo = attInfo;
            _tr = tr;
        }

        protected override bool Update()
        {
            BlockReference br = Entity as BlockReference;
            br.Position = _pos;
            if (br.AttributeCollection.Count != 0)
            {
                foreach (ObjectId id in br.AttributeCollection)
                {
                    DBObject obj = _tr.GetObject(id, OpenMode.ForRead);
                    AttributeReference ar = obj as AttributeReference;

                    // Apply block transform to att def position
                    if (ar != null)
                    {
                        ar.UpgradeOpen();
                        AttInfo ai = _attInfo[ar.ObjectId];
                        ar.Position =
                          ai.Position.TransformBy(br.BlockTransform);
                        if (ai.IsAligned)
                        {
                            ar.AlignmentPoint =
                              ai.Alignment.TransformBy(br.BlockTransform);
                        }
                        if (ar.IsMTextAttribute)
                        {
                            ar.UpdateMTextAttribute();
                        }
                    }
                }
            }
            return true;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions opts = new JigPromptPointOptions("\nSelect insertion point:");
            opts.BasePoint = new Point3d(0, 0, 0);
            opts.UserInputControls = UserInputControls.NoZeroResponseAccepted;

            PromptPointResult ppr = prompts.AcquirePoint(opts);

            if (_pos == ppr.Value)
            {
                return SamplerStatus.NoChange;
            }

            _pos = ppr.Value;

            return SamplerStatus.OK;
        }

        public PromptStatus Run()
        {
            Document doc =
              Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptResult promptResult = ed.Drag(this);
            return promptResult.Status;
        }
    }
}
