using System;
using System.IO;
using Autodesk.Civil.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

using System.Windows.Forms;
using DBTransMan = Autodesk.AutoCAD.DatabaseServices.TransactionManager;
using System.Diagnostics;
using Autodesk.Civil.DatabaseServices;


[assembly: CommandClass(typeof(AeccDotNetAPIDemo.DemoSprint1))]
[assembly: Autodesk.AutoCAD.Runtime.ExtensionApplication(typeof(AeccDotNetAPIDemo.DemoSprint1))]

//Sprit2 Demo
namespace AeccDotNetAPIDemo
{
    public class DemoSprint1 : IExtensionApplication
    {

        private Autodesk.AutoCAD.EditorInput.Editor m_editor =
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

        private CivilDocument doc = CivilApplication.ActiveDocument;
        private Database db = HostApplicationServices.WorkingDatabase;
        private DBTransMan tm;
        private Transaction ts;


        public void Initialize()
        {
        }
        public void Terminate()
        {
        }

        //Get complex aligment properties.
        private void GetStationLocationPoint(String alignmentName)
        {

            ObjectId alignOid = doc.GetSitelessAlignmentId(alignmentName);
            Alignment align = ts.GetObject(alignOid, OpenMode.ForRead) as Alignment;

            m_editor.WriteMessage("\n------User Story 1 :Station should return location as Point2d ----------------\n");
            int stationIndex = 0;
            Station[] stationSet = align.GetStationSet(StationTypes.Major, 100);

            m_editor.WriteMessage("Alignment Station Location Type: {0} \n\n", stationSet[0].Location.GetType());
            foreach (Station MajorStation in stationSet)
            {
                m_editor.WriteMessage("Alignment station:{0}   Location: {1} \n", stationIndex, stationSet[stationIndex++].Location);
            }

            m_editor.WriteMessage("\n\n");

        }

        [CommandMethod("GetStationLocation")]
        public void GetStationLocation()
        {
            tm = db.TransactionManager;
            ts = tm.StartTransaction();

            // User specifies the alignment name
            PromptStringOptions strOpt = new PromptStringOptions("Specify alignment name");
            PromptResult strRes = m_editor.GetString(strOpt);
            String AlignmentName = strRes.StringResult;
            if (AlignmentName == String.Empty)
            {
                AlignmentName = "align-1";
            }

            GetStationLocationPoint(AlignmentName);

            ts.Commit();
        }


        //Get other alignment general properties.
        private void GetAlignmentGeneralInfo(String alignmentName)
        {

            ObjectId alignOid = doc.GetSitelessAlignmentId(alignmentName);
            Alignment align = ts.GetObject(alignOid, OpenMode.ForRead) as Alignment;

            m_editor.WriteMessage("\n------User Story 2 :Alignment properties - OPM ----------------\n");
            m_editor.WriteMessage("{0, -50} {1} \n", "StationIndexIncrement :", align.StationIndexIncrement);
            m_editor.WriteMessage("{0, -50} {1} \n", "Style :", align.StyleName);
            m_editor.WriteMessage("\n\n");

            m_editor.WriteMessage("\n------User Story 3 :Alignment properties - NON-OPM----------------\n");

            m_editor.WriteMessage("{0, -50} {1} \n", "StartingStation :", align.StartingStation);
            m_editor.WriteMessage("{0, -50} {1} \n", "EndingStation :", align.EndingStation);
            m_editor.WriteMessage("{0, -50} {1} \n", "Length :", align.Length);
            m_editor.WriteMessage("{0, -50} {1} \n", "ReferencePoint :", align.ReferencePoint);
            m_editor.WriteMessage("{0, -50} {1} \n", "ReferencePointStation :", align.ReferencePointStation);
            m_editor.WriteMessage("{0, -50} {1} \n", "UseDesignSpeed :", align.UseDesignSpeed);
            m_editor.WriteMessage("{0, -50} {1} \n", "UseDesignCheckSet :", align.UseDesignCheckSet);
            m_editor.WriteMessage("{0, -50} {1} \n", "UseDesignCriteriaFile :", align.UseDesignCriteriaFile);
            m_editor.WriteMessage("{0, -50} {1} \n", "IsSiteless :", align.IsSiteless);
            try
            {
                m_editor.WriteMessage("{0, -50} {1} \n", "SiteName :", align.SiteName);
            }
            catch (System.InvalidOperationException e)
            {
                MessageBox.Show("Alignment.SiteName Calling Error!", e.Message);
            }
        }

        //Get complex alignment properties.
        private void GetAlignmentComplexInfo(String alignmentName)
        {

            ObjectId alignOid = doc.GetSitelessAlignmentId(alignmentName);
            Alignment align = ts.GetObject(alignOid, OpenMode.ForRead) as Alignment;

            m_editor.WriteMessage("\n----------------Alignment Complex properties - begin----------------\n");

            DesignSpeedCollection DesignSpeedColl = align.DesignSpeeds;
            m_editor.WriteMessage("{0, -50} {1} \n", "DesignSpeed Collection count :", DesignSpeedColl.Count);

            ObjectIdCollection LabelGroupColl = align.GetLabelGroupIds();
            m_editor.WriteMessage("{0, -50} {1} \n", "LabelGroup Collection count :", LabelGroupColl.Count);

            ObjectIdCollection LabelColl = align.GetLabelIds();
            m_editor.WriteMessage("{0, -50} {1} \n", "Label Collection count :", LabelColl.Count);

            ObjectIdCollection ProfileIdColl = align.GetProfileIds();
            m_editor.WriteMessage("{0, -50} {1} \n", "ProfileId Collection count :", ProfileIdColl.Count);

            ObjectIdCollection ProfileViewIdColl = align.GetProfileViewIds();
            m_editor.WriteMessage("{0, -50} {1} \n", "ProfileViewId Collection count :", ProfileViewIdColl.Count);

            ObjectIdCollection SampleLineGroupIdColl = align.GetSampleLineGroupIds();
            m_editor.WriteMessage("{0, -50} {1} \n", "SampleLineGroupId Collection count :", SampleLineGroupIdColl.Count);

            StationEquationCollection StationEquationColl = align.StationEquations;
            m_editor.WriteMessage("{0, -50} {1} \n", "StationEquation Collection count :", StationEquationColl.Count);

            SuperelevationCriticalStationCollection SuperEleData = align.SuperelevationCriticalStations;
            m_editor.WriteMessage("{0, -50} {1} \n", "Superelevation Critical Stations count :", SuperEleData.Count);

        }

        [CommandMethod("GetAlignInfo")]
        public void GetAlignmentProperties()
        {
            tm = db.TransactionManager;
            ts = tm.StartTransaction();

            // User specifies the alignment name
            PromptStringOptions strOpt = new PromptStringOptions("Specify alignment name");
            PromptResult strRes = m_editor.GetString(strOpt);
            String AlignmentName = strRes.StringResult;
            if (AlignmentName == String.Empty)
            {
                AlignmentName = "align-1";
            }

            GetAlignmentGeneralInfo(AlignmentName);
            GetAlignmentComplexInfo(AlignmentName);

            ts.Commit();
        }


        private void GetAlignmentInfoAfterChange(String alignmentName)
        {

            ObjectId alignOid = doc.GetSitelessAlignmentId(alignmentName);
            Alignment align = ts.GetObject(alignOid, OpenMode.ForRead) as Alignment;

            m_editor.WriteMessage("\n{0, -50} {1} \n", "UseDesignSpeed :", align.UseDesignSpeed);
            m_editor.WriteMessage("{0, -50} {1} \n", "UseDesignCheckSet :", align.UseDesignCheckSet);
            m_editor.WriteMessage("{0, -50} {1} \n", "UseDesignCriteriaFile :", align.UseDesignCriteriaFile);
            m_editor.WriteMessage("{0, -50} {1} \n", "ReferencePoint :", align.ReferencePoint);
            m_editor.WriteMessage("{0, -50} {1} \n", "ReferencePointStation :", align.ReferencePointStation);
            m_editor.WriteMessage("{0, -50} {1} \n", "StationIndexIncrement :", align.StationIndexIncrement);
            m_editor.WriteMessage("{0, -50} {1} \n", "StyleName :", align.StyleName);

            StationEquationCollection StationEquationColl = align.StationEquations;
            m_editor.WriteMessage("{0, -50} {1} \n", "StationEquation Collection count :", StationEquationColl.Count);

            SuperelevationCriticalStationCollection SuperEleData = align.SuperelevationCriticalStations;
            m_editor.WriteMessage("{0, -50} {1} \n", "Superelevation Critical Stations count :", SuperEleData.Count);

            DesignSpeedCollection DesignSpeedColl = align.DesignSpeeds;
            m_editor.WriteMessage("{0, -50} {1} \n", "DesignSpeed Collection count :", DesignSpeedColl.Count);

        }
        [CommandMethod("GetModifiedAlignmentInfo")]
        public void GetAlignmentInfoModified()
        {

            tm = db.TransactionManager;
            ts = tm.StartTransaction();

            // User specifies the alignment name
            PromptStringOptions strOpt = new PromptStringOptions("Specify alignment name");
            PromptResult strRes = m_editor.GetString(strOpt);
            String AlignmentName = strRes.StringResult;
            if (AlignmentName == String.Empty)
            {
                AlignmentName = "align-1";
            }

            GetAlignmentInfoAfterChange(AlignmentName);

            ts.Commit();

        }


        /*
        *
        UseDesignSpeed	        true
        UseDesignCheckSet	    true
        UseDesignCriteriaFile	true
        ReferencePoint	        (4451.1596, 3713.7837)
        ReferencePointStation	200
        StationIndexIncrement	50
        StyleName	            Plot Style
        */

        //Set some alignment properties.
        private void SetAlignmentInfo(String alignmentName)
        {

            ObjectId alignOid = doc.GetSitelessAlignmentId(alignmentName);
            Alignment align = ts.GetObject(alignOid, OpenMode.ForWrite) as Alignment;

            //align.Name = newAlignmentName;
            align.UseDesignCriteriaFile = true;
            align.UseDesignCheckSet = true;
            align.UseDesignSpeed = true;
            Point2d newReferencePoint = new Point2d(4451.1596, 3713.7837);
            if (DialogResult.OK == MessageBox.Show("Edit ReferencePoint or ReferencePointStation will affect DesignSpeeds,StationEquations and Superelevation\n Do you want to continue? ", "Warning", MessageBoxButtons.OKCancel))
            {
                align.ReferencePoint = newReferencePoint;
                //align.ReferencePointStation = 200;
            }

            try
            {
                align.StationIndexIncrement = 50;
            }
            catch (System.ArgumentOutOfRangeException e)
            {
                MessageBox.Show(e.Message);
            }

            try
            {
                String newStyleName = "Plot Style";
                align.StyleName = newStyleName;
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        [CommandMethod("SetAlignInfo")]
        public void SetAlignmentProperties()
        {
            tm = db.TransactionManager;
            ts = tm.StartTransaction();

            // User specifies the alignment name
            PromptStringOptions strOpt = new PromptStringOptions("Specify alignment name");
            PromptResult strRes = m_editor.GetString(strOpt);
            String AlignmentName = strRes.StringResult;
            if (AlignmentName == String.Empty)
            {
                AlignmentName = "align-1";
            }

            SetAlignmentInfo(AlignmentName);

            ts.Commit();

        }


        //EntityId + Constraint.
        private void GetAlignmentEntitiesInfo(String alignmentName)
        {

            ObjectId alignOid = doc.GetSitelessAlignmentId(alignmentName);
            Alignment align = ts.GetObject(alignOid, OpenMode.ForRead) as Alignment;

            m_editor.WriteMessage("\n---------User Story 4 :Alignment Entities properties- begin----------------\n");

            AlignmentEntityCollection alignColl = align.Entities;
            int count = alignColl.Count;
            int alignID = alignColl.FirstEntity;
            while (true)
            {
                AlignmentEntity alignEntity = alignColl.EntityAtId(alignID);
                Debug.Assert(alignID == alignEntity.EntityId);

                //m_editor.WriteMessage("Alignment EntityId : {0}  Constraint : {1}\n", alignEntity.EntityId, alignEntity.Constraint);
                if (alignID == alignColl.LastEntity)
                    break;
                alignID = alignEntity.EntityAfter;
            }
        }

        [CommandMethod("GetEntityInfo")]
        public void GetEntityInfo()
        {
            tm = db.TransactionManager;
            ts = tm.StartTransaction();

            // User specifies the alignment name
            PromptStringOptions strOpt = new PromptStringOptions("Specify alignment name");
            PromptResult strRes = m_editor.GetString(strOpt);
            String AlignmentName = strRes.StringResult;
            if (AlignmentName == String.Empty)
            {
                AlignmentName = "align-1";
            }

            GetAlignmentEntitiesInfo(AlignmentName);

            ts.Commit();
        }

    }
}