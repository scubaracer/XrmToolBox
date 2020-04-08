namespace Daniel.CrmExcel
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;
    using System.Windows.Forms;

    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;
    using Microsoft.Xrm.Sdk.Query;

    using XrmToolBox.Extensibility;
    using XrmToolBox.Extensibility.Args;
    using XrmToolBox.Extensibility.Interfaces;

    [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
    public partial class CrmExcelTool : PluginControlBase, ICodePlexPlugin, IHelpPlugin, IStatusBarMessenger
    {
        // IGitHubPlugin IPayPalPlugin
        public CrmExcelTool()
        {
            InitializeComponent();
            this.grpExcel.Enabled = false;
            this.cboSolution.Enabled = false;
            this.txtExcelFile.Enabled = false;
            this.btnSelectFile.Enabled = false;
            this.grpUpdate.Enabled = false;
        }

        public event EventHandler<StatusBarMessageEventArgs> SendMessageToStatusBar;

        public string CodePlexUrlName => "CodePlex";

        public string HelpUrl => "http://www.google.com";

        public void FixRelations()
        {
            WorkAsync(
                new WorkAsyncInfo
                    {
                        Message = "Fix relations",
                        Work = (w, e) =>
                            {
                                var selectedNodes = tvwEntities.Nodes["ROOT"].Nodes.Cast<TreeNode>().Where(a => a.Checked).Select(a => a.Text).ToList();
                                var excelToCrm = new ExcelToCrm(Service, w);

                                excelToCrm.FixRelations(selectedNodes);
                            },
                        ProgressChanged = e =>
                            {
                                // If progress has to be notified to user, use the following method:
                                // SetWorkingMessage("Message to display");

                                // If progress has to be notified to user, through the
                                // status bar, use the following method
                                if (SendMessageToStatusBar != null)
                                    SendMessageToStatusBar(this, new StatusBarMessageEventArgs(50, e.UserState.ToString()));
                            },
                        PostWorkCallBack = e =>
                            {
                                if (!e.Cancelled)
                                {
                                    var logger = new LogManager(typeof(CrmExcelTool), ConnectionDetail);

                                    MessageBox.Show("Sheet generated.");
                                }
                            },
                        AsyncArgument = null,
                        IsCancelable = true,
                        MessageWidth = 340,
                        MessageHeight = 150
                    });
        }

        public void GenerateSheet()
        {
            var solution = this.cboSolution.SelectedItem.ToString();
            var languageCode = Convert.ToInt32(cboLanguage.SelectedItem);

            WorkAsync(
                new WorkAsyncInfo
                    {
                        Message = "Generating Excel Sheet",
                        Work = (w, e) =>
                            {
                                var selectedNodes = tvwEntities.Nodes["ROOT"].Nodes.Cast<TreeNode>().Where(a => a.Checked).Select(a => a.Text).ToList();
                                var crmToExcel = new CrmToExcel(Service, w);

                                crmToExcel.CrmToExcelSheet(solution, this.txtExcelFile.Text, selectedNodes, chkUseSolutionXml.Checked, this.chkIncludeOwnerEtc.Checked, languageCode);
                            },
                        ProgressChanged = e =>
                            {
                                // If progress has to be notified to user, use the following method:
                                // SetWorkingMessage("Message to display");

                                // If progress has to be notified to user, through the
                                // status bar, use the following method
                                if (SendMessageToStatusBar != null)
                                    SendMessageToStatusBar(this, new StatusBarMessageEventArgs(50, e.UserState.ToString()));
                            },
                        PostWorkCallBack = e =>
                            {
                                if (!e.Cancelled)
                                {
                                    var logger = new LogManager(typeof(CrmExcelTool), ConnectionDetail);
                                    MessageBox.Show("Sheet generated.");
                                }
                            },
                        AsyncArgument = null,
                        IsCancelable = true,
                        MessageWidth = 340,
                        MessageHeight = 150
                    });
        }

        public void GetSolutions(EntityCollection solutions)
        {
            try
            {
                cboSolution.Enabled = true;

                // Check whether it already exists
                foreach (var solution in solutions.Entities)
                {
                    var isvis = (bool)solution.Attributes["isvisible"];
                    if (isvis)
                    {
                        cboSolution.Items.Add(solution.Attributes["uniquename"]);
                        cboSolutionUpdate.Items.Add(solution.Attributes["uniquename"]);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error getting solutions", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void InitializeData()
        {
            try
            {
                RetrieveAllEntitiesResponse retrieveAllEntitiesResponse = null;
                EntityCollection solutions = null;
                cmdRetieveCrmInformation.Enabled = false;

                WorkAsync(
                    new WorkAsyncInfo
                        {
                            Message = "Retrieving entities",
                            Work = (w, e) =>
                                {
                                    w.ReportProgress(0, "Retrieving Entities");
                                    var retrieveAllEntitiesRequest = new RetrieveAllEntitiesRequest();
                                    retrieveAllEntitiesRequest.EntityFilters = EntityFilters.Entity;
                                    retrieveAllEntitiesResponse = (RetrieveAllEntitiesResponse)Service.Execute(retrieveAllEntitiesRequest);
                                    w.ReportProgress(80, "Retrieving Solution");
                                    var queryCheckForSampleSolution = new QueryExpression { EntityName = "solution", ColumnSet = new ColumnSet(true) };

                                    // Create the solution if it does not already exist.
                                    solutions = Service.RetrieveMultiple(queryCheckForSampleSolution);
                                    e.Result = retrieveAllEntitiesResponse;
                                },
                            ProgressChanged = e =>
                                {
                                    // If progress has to be notified to user, use the following method:
                                    // SetWorkingMessage("Message to display");

                                    // If progress has to be notified to user, through the
                                    // status bar, use the following method
                                    if (SendMessageToStatusBar != null)
                                        SendMessageToStatusBar(this, new StatusBarMessageEventArgs(50, e.UserState.ToString()));
                                },
                            PostWorkCallBack = e =>
                                {
                                    this.RetrieveEntities(retrieveAllEntitiesResponse);
                                    this.GetSolutions(solutions);
                                    cmdRetieveCrmInformation.Enabled = true;
                                    if (!e.Cancelled)
                                    {
                                        var logger = new LogManager(typeof(CrmExcelTool), ConnectionDetail);

                                        MessageBox.Show("CRM Updated.");
                                    }
                                },
                            AsyncArgument = null,
                            IsCancelable = true,
                            MessageWidth = 340,
                            MessageHeight = 150
                        });
            }
            catch (Exception ex)
            {
                cmdRetieveCrmInformation.Enabled = true;
                MessageBox.Show(ex.Message, "Error getting entities", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ProcessWhoAmI()
        {
            WorkAsync(
                new WorkAsyncInfo
                    {
                        Message = "Who Im I",
                        Work = (w, e) =>
                            {
                                var request = new WhoAmIRequest();
                                var response = (WhoAmIResponse)Service.Execute(request);

                                e.Result = response.UserId;
                            },
                        ProgressChanged = e => { },
                        PostWorkCallBack = e => { MessageBox.Show(string.Format("You are {0}", (Guid)e.Result)); },
                        AsyncArgument = null,
                        IsCancelable = true,
                        MessageWidth = 340,
                        MessageHeight = 150
                    });
        }

        public void UpdateCrm()
        {
            var solution = this.cboSolutionUpdate.SelectedItem.ToString();
            var languageCode = Convert.ToInt32(cboLanguage.SelectedItem);
            WorkAsync(
                new WorkAsyncInfo
                    {
                        Message = "Generating Excel Sheet",
                        Work = (w, e) =>
                            {
                                w.ReportProgress(0, "Updateing CRM");
                                var excelToCrm = new ExcelToCrm(Service, w);

                                excelToCrm.Start(this.textFileNameUpdate.Text, solution, languageCode);
                            },
                        ProgressChanged = e =>
                            {
                                // If progress has to be notified to user, use the following method:
                                // SetWorkingMessage("Message to display");

                                // If progress has to be notified to user, through the
                                // status bar, use the following method
                                if (SendMessageToStatusBar != null)
                                    SendMessageToStatusBar(this, new StatusBarMessageEventArgs(50, e.UserState.ToString()));
                            },
                        PostWorkCallBack = e =>
                            {
                                if (!e.Cancelled)
                                {
                                    var logger = new LogManager(typeof(CrmExcelTool), ConnectionDetail);

                                    MessageBox.Show("CRM Updated.");
                                }
                            },
                        AsyncArgument = null,
                        IsCancelable = true,
                        MessageWidth = 340,
                        MessageHeight = 150
                    });
        }

        private void BtnAddEntitiesToInclude_Click(object sender, EventArgs e)
        {
            // add WinCare base entities
            foreach (TreeNode node in tvwEntities.Nodes["ROOT"].Nodes)
            {
                if (txtEntitiesToInclude.Text.ToLower().Contains(node.Text))
                {
                    node.Checked = true;
                }
            }
        }

        private void BtnCloseClick(object sender, EventArgs e)
        {
            CloseTool();
        }

        private void btnRefreshEntities_Click(object sender, EventArgs e)
        {
            var retrieveAllEntitiesRequest = new RetrieveAllEntitiesRequest();
            retrieveAllEntitiesRequest.EntityFilters = EntityFilters.Entity;
            var retrieveAllEntitiesResponse = (RetrieveAllEntitiesResponse)Service.Execute(retrieveAllEntitiesRequest);
            this.RetrieveEntities(retrieveAllEntitiesResponse);
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Multiselect = false;
            openFileDialog1.Filter = "Excel Files |*.xlsx";
            openFileDialog1.InitialDirectory = Application.StartupPath;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtExcelFile.Text = this.openFileDialog1.FileName;
            }
        }

        private void BtnWhoAmIClick(object sender, EventArgs e)
        {
            ExecuteMethod(ProcessWhoAmI);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                ExecuteMethod(FixRelations);
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                var logger = new LogManager(typeof(CrmExcelTool), ConnectionDetail);
                logger.LogError("Error creating excelfile", ex);
                MessageBox.Show(ex.Message, "Error creating excelfile", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // var selectedNodes = tvwEntities.Nodes["ROOT"].Nodes.Cast<TreeNode>().Where(a => a.Checked).Select(a => a.Text).ToList();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // add WinCare base entities
            foreach (TreeNode node in tvwEntities.Nodes["ROOT"].Nodes)
            {
                if (node.Text.ToLower().Contains("account") || node.Text.ToLower().Contains("contact") || node.Text.ToLower().Contains("systemuser") || node.Text.ToLower().Contains("appointment") || node.Text.ToLower().Contains("letter")
                    || node.Text.ToLower().Contains("task") || node.Text.ToLower().Contains("phonecall") || node.Text.ToLower().Contains("activitypointer") || node.Text.ToLower().Contains("activityparty")
                    || node.Text.ToLower().Contains("wv_"))
                {
                    node.Checked = true;
                }
            }
        }

        private void CboSolutionSelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            this.txtExcelFile.Enabled = true;
            this.btnSelectFile.Enabled = true;
            this.grpExcel.Enabled = true;
            this.grpUpdate.Enabled = true;
            this.txtExcelFile.Text = this.cboSolution.SelectedItem + ".xlsx";
            this.Cursor = Cursors.Default;
        }

        private void cboSolutionUpdate_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cboSolutionUpdate.SelectedItem.ToString()))
            {
                cmdUpdateCrm.Enabled = true;
            }
        }

        private void cmdGenerateExcelSheet_Click(object sender, EventArgs e)
        {
            try
            {
                if (cboLanguage.SelectedItem == null)
                {
                    MessageBox.Show("First select a language", "Error creating excelfile", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ExecuteMethod(GenerateSheet);
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                var logger = new LogManager(typeof(CrmExcelTool), ConnectionDetail);
                logger.LogError("Error creating excelfile", ex);
                MessageBox.Show(ex.Message, "Error creating excelfile", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmdRetieveCrmInformation_Click(object sender, EventArgs e)
        {
            this.grpExcel.Enabled = false;
            this.cboSolution.Enabled = false;
            this.txtExcelFile.Enabled = false;
            this.btnSelectFile.Enabled = false;
            this.grpUpdate.Enabled = false;
            ExecuteMethod(InitializeData);
            this.grpExcel.Enabled = true;
            this.grpUpdate.Enabled = true;
            this.txtExcelFile.Enabled = true;
            this.btnSelectFile.Enabled = true;
            this.cboSolution.Enabled = false;
            this.cmdUpdateCrm.Enabled = false;
        }

        private void cmdSelectExcelForUpdate_Click(object sender, EventArgs e)
        {
            this.openFileDialog2.FileName = string.Empty;
            this.openFileDialog2.Multiselect = false;
            this.openFileDialog2.Filter = "Excel Files |*.xlsx";
            this.openFileDialog2.InitialDirectory = Application.StartupPath;
            if (this.openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                this.textFileNameUpdate.Text = this.openFileDialog2.FileName;
            }
        }

        private void cmdSelectPrefix_Click(object sender, EventArgs e)
        {
            foreach (TreeNode node in tvwEntities.Nodes["ROOT"].Nodes)
            {
                if (node.Text.Contains(txtPrefix.Text))
                {
                    node.Checked = true;
                }
            }
        }

        private void cmdUpdateCrm_Click(object sender, EventArgs e)
        {
            try
            {
                if (cboLanguage.SelectedItem == null)
                {
                    MessageBox.Show("First select a language", "Error creating excelfile", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ExecuteMethod(this.UpdateCrm);
            }
            catch (Exception ex)
            {
                var logger = new LogManager(typeof(CrmExcelTool), ConnectionDetail);
                logger.LogError("Error creating excelfile", ex);
                MessageBox.Show(ex.Message, "Error updateing excelfile", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CrmExcelTool_Load(object sender, EventArgs e)
        {
        }

        private void RetrieveEntities(RetrieveAllEntitiesResponse retrieveAllEntitiesResponse)
        {
            try
            {
                tvwEntities.Sort();
                tvwEntities.Nodes.Clear();
                tvwEntities.Nodes.Add("ROOT", "Root");

                // Iterate through the retrieved entities
                foreach (var entity in retrieveAllEntitiesResponse.EntityMetadata)
                {
                    if (entity.IsIntersect != null)
                    {
                        if (entity.IsIntersect.Value != true)
                        {
                            tvwEntities.Nodes["ROOT"].Nodes.Add(entity.LogicalName, entity.LogicalName);
                        }
                        else
                        {
                            var s = entity.LogicalName;
                        }
                    }
                    else
                    {
                        tvwEntities.Nodes["ROOT"].Nodes.Add(entity.LogicalName, entity.LogicalName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error getting entities", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}