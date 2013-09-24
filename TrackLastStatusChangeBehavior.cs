using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

using HPMSdk;
using Hansoft.ObjectWrapper;
using Hansoft.ObjectWrapper.CustomColumnValues;

using Hansoft.Jean.Behavior;

namespace Hansoft.Jean.Behavior.TrackLastStatusChangeBehavior
{
    public class TrackLastStatusChangeBehavior : AbstractBehavior
    {
        string projectName;
        string trackingColumnName;
        string trackedColumnName;
        string viewName;
        EHPMReportViewType viewType;
        Project project;
        private ProjectView projectView;
        HPMProjectCustomColumnsColumn trackingColumn;
        HPMProjectCustomColumnsColumn trackedColumn;
        string title;

        public TrackLastStatusChangeBehavior(XmlElement configuration)
            : base(configuration)
        {
            projectName = GetParameter("HansoftProject");
            trackingColumnName = GetParameter("TrackingColumn");
            trackedColumnName = GetParameter("TrackedColumn");
            viewName = GetParameter("View");
            viewType = GetViewType(viewName);
            title = "TrackLastStatusChangeBehavior: " + configuration.InnerText;
        }

        public override void Initialize()
        {
            project = HPMUtilities.FindProject(projectName);
            if (project == null)
                throw new ArgumentException("Could not find project:" + projectName);
            if (viewType == EHPMReportViewType.AgileBacklog)
                projectView = project.ProductBacklog;
            else if (viewType == EHPMReportViewType.AllBugsInProject)
                projectView = project.BugTracker;
            else
                projectView = project.Schedule;
            trackedColumn = projectView.GetCustomColumn(trackedColumnName);
            if (trackedColumn == null)
                throw new ArgumentException("Could not find custom column in view " + viewName + " " + trackedColumnName);
            trackingColumn = projectView.GetCustomColumn(trackingColumnName);
            if (trackingColumn == null)
                throw new ArgumentException("Could not find custom column in view " + viewName + " " + trackingColumnName);
            DoUpdateFromHistory();
        }

        public override string Title
        {
            get { return title; }
        }

        // TODO: Subject to refactoting
        private EHPMReportViewType GetViewType(string viewType)
        {
            switch (viewType)
            {
                case ("Agile"):
                    return EHPMReportViewType.AgileMainProject;
                case ("Scheduled"):
                    return EHPMReportViewType.ScheduleMainProject;
                case ("Bugs"):
                    return EHPMReportViewType.AllBugsInProject;
                case ("Backlog"):
                    return EHPMReportViewType.AgileBacklog;
                default:
                    throw new ArgumentException("Unsupported View Type: " + viewType);
            }
        }

        private void DoUpdateFromHistory()
        {
            foreach (Task task in projectView.DeepLeaves)
                DoUpdateFromHistory(task);
        }

        private void DoUpdateFromHistory(Task task)
        {
            HPMDataHistoryGetHistoryParameters pars = new HPMDataHistoryGetHistoryParameters();
            pars.m_DataID = task.UniqueTaskID;
            pars.m_FieldID = EHPMStatisticsField.NoStatistics;
            pars.m_FieldData = 0;
            pars.m_DataIdent0 = EHPMStatisticsScope.NoStatisticsScope;
            pars.m_DataIdent1 = 0;
            HPMDataHistory history = SessionManager.Session.DataHistoryGetHistory(pars);
            if (history != null)
                DoUpdateFromHistory(task, history);
        }

        private void DoUpdateFromHistory(Task task, HPMDataHistory history)
        {
            DateTimeValue storedValue =  (DateTimeValue)task.GetCustomColumnValue(trackingColumn);
            // ToInt64 will return the value as microseconds since 1970 Jan 1
            ulong storedHpmTime = storedValue.ToHpmDateTime();
            if (history.m_Latests.m_Time > storedHpmTime)
            {
                foreach (HPMDataHistoryEntry entry in history.m_HistoryEntries)
                {
                    // Check if it is the status field
                    if (entry.m_FieldID == 15)
                    {
                        if (entry.m_Time > storedHpmTime)
                        {
                            storedHpmTime = entry.m_Time;
                            task.SetCustomColumnValue(trackingColumn, DateTimeValue.FromHpmDateTime(task, trackingColumn, storedHpmTime));
                        }
                    }
                }
            }
        }

        public override void OnTaskChangeCustomColumnData(TaskChangeCustomColumnDataEventArgs e)
        {
            if (e.Data.m_ColumnHash == trackedColumn.m_Hash)
            {
                Task task = Task.GetTask(e.Data.m_TaskID);
                if (task.MainProjectID.m_ID == project.UniqueID.m_ID && task.ProjectView.Equals(projectView))
                    task.SetCustomColumnValue(trackingColumn, DateTimeValue.FromHpmDateTime(task, trackingColumn, HPMUtilities.HPMNow()));
            }
        }

        public override void OnDataHistoryReceived(DataHistoryReceivedEventArgs e)
        {
            if (SessionManager.Session.UtilIsIDTask(e.Data.m_UniqueIdentifier) || SessionManager.Session.UtilIsIDTaskRef(e.Data.m_UniqueIdentifier))
            {
                Task task = Task.GetTask(e.Data.m_UniqueIdentifier);
                DoUpdateFromHistory(task);
            }
        }
    }
}

