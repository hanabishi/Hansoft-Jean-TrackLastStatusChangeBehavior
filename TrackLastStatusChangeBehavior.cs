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
        bool initializationOK = false;
        string trackingColumnName;
        string trackedColumnName;
        string viewName;
        EHPMReportViewType viewType;
        List<Project> projects;
        private List<ProjectView> projectViews;
        bool inverted = false;
        HPMProjectCustomColumnsColumn trackingColumn;
        HPMProjectCustomColumnsColumn trackedColumn;
        string title;

        public TrackLastStatusChangeBehavior(XmlElement configuration)
            : base(configuration)
        {
            projectName = GetParameter("HansoftProject");
            string invert = GetParameter("InvertedMatch");
            if (invert != null)
                inverted = invert.ToLower().Equals("yes");
            trackingColumnName = GetParameter("TrackingColumn");
            trackedColumnName = GetParameter("TrackedColumn");
            viewName = GetParameter("View");
            viewType = GetViewType(viewName);
            title = "TrackLastStatusChangeBehavior: " + configuration.InnerText;
        }

        public override void Initialize()
        {
            projects = new List<Project>();
            projectViews = new List<ProjectView>();
            initializationOK = false;
            projects = HPMUtilities.FindProjects(projectName, inverted);
            if (projects.Count == 0)
                throw new ArgumentException("Could not find any matching project:" + projectName);
            foreach (Project project in projects)
            {
                ProjectView projectView;
                if (viewType == EHPMReportViewType.AgileBacklog)
                    projectView = project.ProductBacklog;
                else if (viewType == EHPMReportViewType.AllBugsInProject)
                    projectView = project.BugTracker;
                else
                    projectView = project.Schedule;
                projectViews.Add(projectView);
            }

            trackedColumn = projectViews[0].GetCustomColumn(trackedColumnName);
            if (trackedColumn == null)
                throw new ArgumentException("Could not find custom column in view " + viewName + " " + trackedColumnName);
            trackingColumn = projectViews[0].GetCustomColumn(trackingColumnName);
            if (trackingColumn == null)
                throw new ArgumentException("Could not find custom column in view " + viewName + " " + trackingColumnName);
            initializationOK = true;
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
            foreach (ProjectView projectView in projectViews)
            {
                foreach (Task task in projectView.DeepLeaves)
                    DoUpdateFromHistory(task);
            }
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
            // Ensure that we get the custom column of the right project
            HPMProjectCustomColumnsColumn actualCustomColumn = task.ProjectView.GetCustomColumn(trackingColumn.m_Name);
            DateTimeValue storedValue = (DateTimeValue)task.GetCustomColumnValue(actualCustomColumn);
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
                            task.SetCustomColumnValue(trackingColumn, DateTimeValue.FromHpmDateTime(task, actualCustomColumn, storedHpmTime));
                        }
                    }
                }
            }
        }

        public override void OnTaskChangeCustomColumnData(TaskChangeCustomColumnDataEventArgs e)
        {
            if (initializationOK)
            {
                if (e.Data.m_ColumnHash == trackedColumn.m_Hash)
                {
                    Task task = Task.GetTask(e.Data.m_TaskID);
                    if (projects.Contains(task.Project) && projectViews.Contains(task.ProjectView))
                        task.SetCustomColumnValue(trackingColumn, DateTimeValue.FromHpmDateTime(task, trackingColumn, HPMUtilities.HPMNow()));
                }
            }
        }

        public override void OnDataHistoryReceived(DataHistoryReceivedEventArgs e)
        {
            if (initializationOK)
            {
                if (SessionManager.Session.UtilIsIDTask(e.Data.m_UniqueIdentifier) || SessionManager.Session.UtilIsIDTaskRef(e.Data.m_UniqueIdentifier))
                {
                    Task task = Task.GetTask(e.Data.m_UniqueIdentifier);
                    DoUpdateFromHistory(task);
                }
            }
        }
    }
}

