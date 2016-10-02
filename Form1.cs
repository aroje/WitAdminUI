using System;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Xml;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace WitAdminTool
{
    public partial class WitAdminToolForm : Form
    {
        TfsTeamProjectCollection tfsTPC;
        WorkItemStore workitemStore;

        #region Initialize
        public WitAdminToolForm(TfsTeamProjectCollection tfsTPCConnection)
        {
            InitializeComponent();
            try
            {
                tfsTPC = tfsTPCConnection;
                InitForm();
             }
            catch (Exception ex)
            {
                MessageBoxError(ex.Message);
            }
        }

 
        private void InitForm()
        {
            workitemStore = (WorkItemStore)tfsTPC.GetService(typeof(WorkItemStore));
            this.lblCollection.Text = tfsTPC.Uri.AbsoluteUri;

            InitProjects();
            InitWorkItemTypePage();
            
            InitFieldPage();
            InitLinkTypePage();
            InitGloballistPage();

            textBoxExportPath.Text = Path.GetTempPath();

            tabControl_ProjectLevelActions.SelectedIndex = 0;
        }

        private void InitCommand()
        {
            textBoxResult.Text = string.Empty;
            textBoxCommand.Text = string.Empty;
            textBoxImportFilePath.Text = string.Empty;
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                TeamProjectPicker tpp = new TeamProjectPicker();
                if (tpp.ShowDialog() == DialogResult.OK)
                {
                    tfsTPC = tpp.SelectedTeamProjectCollection;
                }
            }
            catch (Exception ex)
            {
                MessageBoxError(ex.Message);
            }

            InitForm();
        }
        private void tabControl_ProjectLevelActions_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (this.tabControl_ProjectLevelActions.SelectedIndex)
            {
                case 0:
                    this.InitWorkItemTypePage();
                    break;
                case 1:
                    this.InitWorkItemPage();
                    break;
             }
        }
        private void tabControl_ProjectLevelActions_Selected(object sender, TabControlEventArgs e)
        {
            InitCommand();
        }

        private void tabControl_CollectionLevelActions_Selected(object sender, TabControlEventArgs e)
        {
            InitCommand();
        }


        private void tabControlContainer_SelectedIndexChanged(object sender, EventArgs e)
        {
            InitCommand();
        }

        private void tabControl_CollectionLevelActions_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (this.tabControl_CollectionLevelActions.SelectedIndex)
            {
                case 0:
                    InitFieldPage();
                    break;
                case 1:
                    InitLinkTypePage();
                    break;
                case 2:
                    InitGloballistPage();
                    break;
                default:
                    break;
            }
        }

        private void InitProjects()
        {
            this.workitemStore.SyncToCache();
            this.listProject.Items.Clear();
            foreach (Project project in this.workitemStore.Projects)
            {
                this.listProject.Items.Add(project.Name);
            }
            listProject.SelectedIndex = 0;
        }
        private void InitWorkItemTypePage()
        {
            InitWorkItemList();
            InitCommand();
        }
        private void InitWorkItemList(bool clearXMLfield=true)
        {
            workitemStore.SyncToCache();
            listWit.Items.Clear();
            if (listProject.SelectedItem == null)
            {
                return;
            }
            string project = listProject.SelectedItem.ToString();
            foreach (WorkItemType wit in this.workitemStore.Projects[project].WorkItemTypes)
            {
                listWit.Items.Add(wit.Name);
            }
            if (clearXMLfield)
                txtWitXml.Text = string.Empty;
        }

        private void InitWorkItemPage()
        {
            this.txtWiID.Text = string.Empty;
            
        }
        DataTable dtFields = null;

        private void InitFieldPage()
        {
            this.workitemStore.SyncToCache();

            dtFields = new DataTable();
            dtFields.Columns.Add("Reference Name");
            dtFields.Columns.Add("Name");
            dtFields.Columns.Add("Type");
            dtFields.Columns.Add("Report Type");
            dtFields.Columns.Add("Indexed");

            foreach (FieldDefinition fd in this.workitemStore.FieldDefinitions)
            {
                dtFields.Rows.Add(fd.ReferenceName, fd.Name, fd.FieldType.ToString(),
                    fd.ReportingAttributes.Type.ToString().ToLower(), fd.IsIndexed.ToString());
            }

            this.gvFields.DataSource = dtFields;
            this.gvFields.Sort(gvFields.Columns["Reference Name"], ListSortDirection.Ascending);
        }

        private void InitLinkTypePage()
        {
            this.workitemStore.SyncToCache();

            WorkItemLinkTypeCollection ltc = this.workitemStore.WorkItemLinkTypes;

            DataTable dtLik = new DataTable();
            dtLik.Columns.Add("Ref Name");
            dtLik.Columns.Add("ForwordEnd Name");
            dtLik.Columns.Add("ReverseEnd Name");
            dtLik.Columns.Add("Topology");
            dtLik.Columns.Add("Activate");


            foreach (WorkItemLinkType lt in ltc)
            {
                string refName = lt.ReferenceName;
                string fn = lt.ForwardEnd.Name;
                string rn = lt.ReverseEnd.Name;
                string topology = lt.LinkTopology.ToString();
                string actYN = lt.IsActive.ToString();

                dtLik.Rows.Add(refName, fn, rn, topology, actYN);
            }

            gvLink.DataSource = dtLik;
            gvLink.Rows[0].Selected = true;

            txtLinkXml.Text = string.Empty;
        }

        private void InitGloballistPage()
        {
            this.workitemStore.SyncToCache();

            this.txtGLXml.Text = string.Empty;

            if (treeViewGlobalLists.Nodes.Count > 0)
            {
                treeViewGlobalLists.Nodes.RemoveAt(0);
            }

            Hourglass(true);

            XmlDocument globalListXmlDoc = this.workitemStore.ExportGlobalLists();

            TreeNode rootNode = new TreeNode("Global List");

            foreach (System.Xml.XmlNode list in globalListXmlDoc.DocumentElement.ChildNodes)
            {

                TreeNode listNode = new TreeNode(list.Attributes["name"].Value);
                listNode.Tag = "Server";

                foreach (System.Xml.XmlNode listItem in list.ChildNodes)
                {
                    TreeNode treeListItem = new TreeNode(listItem.Attributes["value"].Value);
                    listNode.Nodes.Add(treeListItem);
                }

                rootNode.Nodes.Add(listNode);

            }
            rootNode.ExpandAll();
            treeViewGlobalLists.Nodes.Add(rootNode);
            Hourglass(false);
        }
        #endregion

        #region Work Item Type
        private void listProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            InitWorkItemTypePage();
        }
        private void listWit_SelectedIndexChanged(object sender, EventArgs e)
        {
            InitCommand();
            txtWitXml.Text = string.Empty;
        }

        private void btnWit_Rename_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.RENAMEWITD);
                return;
            }

            if (listWit.SelectedItem == null)
            {
                MessageBoxWarning("Select Work Item Type");
                return;
            }
            if (this.txtNewWitName.Text == string.Empty)
            {
                MessageBoxWarning("Input New Work Item Type Name");
                return;
            }

            string command = GenerateComand(WitadminActions.RENAMEWITD, this.listWit.SelectedItem.ToString(), this.txtNewWitName.Text);

            if (ExcuteCommand(
                string.Format("Do you really want to RENAME the work item type '{0}' in project '{1}'?", 
                                listWit.SelectedItem.ToString(), listProject.SelectedItem.ToString()), 
                command))
                InitWorkItemList();
        }

        private void btnWit_Delete_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.DESTROYWITD);
                return;
            }

            if (listWit.SelectedItem == null)
            {
                MessageBoxWarning("Select Work Item Type");
                return;
            }

            string command = GenerateComand(WitadminActions.DESTROYWITD, this.listWit.SelectedItem.ToString());

            if (ExcuteCommand(
                string.Format("Do you really want to DELETE the work item type '{0}' in project '{1}'?",
                                listWit.SelectedItem.ToString(), listProject.SelectedItem.ToString()),
                command))
                InitWorkItemList();
        }

        private void btnWit_Export_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.EXPORTWITD);
                return;
            }

            if (listWit.SelectedItem == null)
            {
                MessageBoxWarning("Select Work Item Type");
                return;
            }

            if (textBoxExportPath.Text.Length == 0)
            {
                MessageBoxWarning("Select Export Path");
                return;
            }

            string witXmlFilePath = Path.Combine(textBoxExportPath.Text, 
                this.listWit.SelectedItem.ToString().Replace(" ", string.Empty) + ".xml");

            string command = GenerateComand(WitadminActions.EXPORTWITD, this.listWit.SelectedItem.ToString(), witXmlFilePath);
            txtWitXml.Text = string.Empty;
            if (ExcuteCommand(string.Empty, command))
                txtWitXml.Text = ReadFileToXmlControl(witXmlFilePath);
        }

        private void btnWit_ImportFromXml_Click(object sender, EventArgs e)
        {
            if (ImportFromXMLHelper(WitadminActions.IMPORTWITD, "Work Item Type", txtWitXml.Text))
            {
                InitWorkItemList(false);
            }
        }

        private void btn_WIT_ImportFromFile_Click(object sender, EventArgs e)
        {
            if (ImportFromFileHelper(WitadminActions.IMPORTWITD, "Work Item Type"))
            {
                InitWorkItemList();
            }
        }
                
        private void btn_WIT_ValidateFile_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.IMPORTWITD_V);
                return;
            }

            if (textBoxImportFilePath.Text.Length == 0)
            {
                MessageBoxWarning("File Name empty");
                return;
            }

            string command = GenerateComand(WitadminActions.IMPORTWITD_V, textBoxImportFilePath.Text);
            ExcuteCommand(string.Empty, command);
        }
        #endregion

        #region Destroy Work Item
        private void btnGetWi_Click(object sender, EventArgs e)
        {
            Hourglass(true);
            this.GetWi();
            Hourglass(false);
        }

        private void GetWi()
        {
            string query = "select [system.id], [System.Title], [System.TeamProject], [System.WorkItemType],[System.CreatedDate],[System.CreatedBy], [System.AssignedTo] from workitems";

            if(radioButtonSelectedProject.Checked)
            {
                if (listProject.SelectedItem == null )
                {
                    MessageBoxWarning("Select Project");
                    return;
                }
                query = string.Format("{0} WHERE [System.TeamProject] = '{1}'", query, listProject.SelectedItem.ToString());
            }

            this.gvWi.DataSource = null;
            WorkItemCollection wic = this.workitemStore.Query(query);

            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Type");
            dt.Columns.Add("Title");
            dt.Columns.Add("Project");
            dt.Columns.Add("CreatedBy");
            dt.Columns.Add("CreatedDate");
            dt.Columns.Add("AssignedTo");

            int counter = 0;
            foreach (WorkItem wi in wic)
            {
                dt.Rows.Add(wi.Id.ToString(), wi.Type.Name, wi.Title,wi.Project.Name, wi.CreatedBy, wi.CreatedDate.ToString("yyyy-MM-dd"), wi.Fields[CoreField.AssignedTo].Value.ToString());
                counter++;
            }

            this.lbl_WI_WICount.Text = counter.ToString();
            this.gvWi.DataSource = dt;
        }

        private void btnDestroyWis_SelectAllWis_Click(object sender, EventArgs e)
        {
            MarkListItems(gvWi.Rows.Count, true);
        }

         private void btnDestroyWis_UnselectAllWis_Click(object sender, EventArgs e)
        {
            MarkListItems(gvWi.Rows.Count, false);
        }
        private void MarkListItems(int count, bool selectedFlag)
        {
            int counter = 0;
            foreach (DataGridViewRow row in gvWi.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[0];
                chk.Value = selectedFlag;
                counter++;
                if (counter >= count)
                    break;
            }
        }

        private void btnDestroyWis_SelectNWis_Click(object sender, EventArgs e)
        {
            if (textBoxWICount.Text == null)
            {
                MessageBoxWarning("Input Number of Work Item to select");
                textBoxWICount.Focus();
                return;
            }

            MarkListItems(Int32.Parse(textBoxWICount.Text), true);
        }

        private void btnDestroyWis_DestroyWisByIds_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.DESTROYWI);
                return;
            }

            if (this.txtWiID.Text == string.Empty)
            {
                MessageBoxWarning("Input Work Item ID(s)");
                txtWiID.Focus();
                return;
            }

            string command = GenerateComand(WitadminActions.DESTROYWI, this.txtWiID.Text);
            if (ExcuteCommand(
                string.Format("Do you really want to DELETE selected work item(s) in project '{0}'?",
                                listProject.SelectedItem.ToString()),
                command))
                this.GetWi();
        }

        private void btnDestroyWis_DestroyWisBySelection_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.DESTROYWI);
                return;
            }

            string ids = string.Empty;

            for (int i = 0; i < this.gvWi.Rows.Count; i++)
            {
                if ((bool)this.gvWi.Rows[i].Cells[0].EditedFormattedValue)
                {
                    if (this.gvWi.Rows[i].Cells[1].Value == null || this.gvWi.Rows[i].Cells[1].Value.ToString() == string.Empty)
                    {
                        continue;
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(ids))
                        {
                            ids = this.gvWi.Rows[i].Cells[1].Value.ToString();
                        }
                        else
                        {
                            ids += "," + this.gvWi.Rows[i].Cells[1].Value.ToString();
                        }
                    }
                }
            }

            if (ids == string.Empty)
            {
                MessageBoxWarning("Check Work Item(s)");
                return;
            }

            string command = GenerateComand(WitadminActions.DESTROYWI, ids);
            if (ExcuteCommand(
                string.Format("Do you really want to DELETE selected work item(s) in project '{0}'?",
                                listProject.SelectedItem.ToString()),
                command))
                this.GetWi();
        }
        #endregion

        #region Fields

        private void btnFdSearchString_Click(object sender, EventArgs e)
        {
            // Set the search string:
            string filterString = txtFdSearchString.Text;

            DataView dvFiltered;
            string filter = string.Format("[Reference Name] LIKE '*{0}*' OR [Name] LIKE '*{0}*' ", filterString);
            dvFiltered = new DataView(dtFields, filter, "type Desc", DataViewRowState.CurrentRows);
            gvFields.DataSource = dvFiltered;
            txtFdSearchStringFoundCount.Text = dvFiltered.Count.ToString();
            gvFields.Sort(gvFields.Columns["Reference Name"], ListSortDirection.Ascending);
        }

        private void txtFdSearchString_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnFdSearchString_Click(this, new EventArgs());
            }
        }

        private void gvFields_SelectionChanged(object sender, EventArgs e)
        {

            this.txtFdName.Text = string.Empty;
            this.txtFdRefName.Text = string.Empty;
            this.txtFdType.Text = string.Empty;

            if (gvFields.SelectedRows.Count != 1)
                return;

            txtFdName.Text = gvFields.SelectedRows[0].Cells["Name"].Value.ToString();
            txtFdRefName.Text = gvFields.SelectedRows[0].Cells["Reference Name"].Value.ToString();
            txtFdType.Text = gvFields.SelectedRows[0].Cells["Type"].Value.ToString();
            cbbFdReportType.SelectedItem = gvFields.SelectedRows[0].Cells["Report Type"].Value.ToString().ToLower();
            if (gvFields.SelectedRows[0].Cells["Indexed"].Value.ToString().Equals("True"))
            {
                cbbIndexYN.SelectedItem = "Y";
                btnFields_Index.Text = "Index Off";
            }
            else
            {
                this.cbbIndexYN.SelectedItem = "N";
                this.btnFields_Index.Text = "Index On";
            }

        }

        private void btnFields_Delete_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.DELETEFIELD);
                return;
            }

            if (this.gvFields.SelectedRows.Count == 0)
            {
                MessageBoxWarning("Select Field");
                return;
            }

            string command = GenerateComand(WitadminActions.DELETEFIELD, txtFdRefName.Text);
            if (ExcuteCommand(
                string.Format("Do you really want to DELETE the field '{0}' from project collection?",
                                txtFdName.Text),
                command))
                    InitFieldPage();
        }

        private void btnFields_Change_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.CHANGEFIELD_1);
                return;
            }

            if (this.gvFields.SelectedRows.Count == 0)
            {
                MessageBoxWarning("Select Field");
                return;
            }

            string command = string.Empty;
            if (string.IsNullOrEmpty(this.cbbFdReportType.SelectedText))
            {
                command = GenerateComand(WitadminActions.CHANGEFIELD_1, txtFdRefName.Text, this.txtFdName.Text);
            }
            else
            {
                command = GenerateComand(WitadminActions.CHANGEFIELD_2, txtFdRefName.Text, this.txtFdName.Text, this.cbbFdReportType.SelectedText);
            }


            if (ExcuteCommand(
                string.Format("Do you really want to CHANGE the field '{0}' from project collection?",
                                txtFdName.Text),
                command))

                InitFieldPage();
        }
        private void btnFields_Index_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.INDEXFIELD);
                return;
            }

            string onOffSwitch= string.Empty;
            
            if (this.btnFields_Index.Text == "Index On")
                  onOffSwitch = "on";
            else
                  onOffSwitch = "of";
            
            string command = GenerateComand(WitadminActions.INDEXFIELD, txtFdRefName.Text, onOffSwitch);
            if (ExcuteCommand(
                string.Format("Do you really want to INDEX the field '{0}' from project collection?",
                                txtFdName.Text),
                command))
                InitFieldPage();
        }
        #endregion

        #region LinkType
        private bool IsLinkSelected()
        {
            if (this.gvLink.SelectedRows == null
                || this.gvLink.SelectedRows.Count == 0
                || this.gvLink.SelectedRows[0].Cells[0].Value == null
                || this.gvLink.SelectedRows[0].Cells[0].Value.ToString() == string.Empty
                )
            {
                MessageBoxWarning("Select LinkType Row");
                return false;
            }
            return true;
        }

        private void btnLinks_Delete_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.DELETELINKTYPE);
                return;
            }

            if (!IsLinkSelected())
                return;

            string linkType = this.gvLink.SelectedRows[0].Cells[0].Value.ToString();
            string command = GenerateComand(WitadminActions.DELETELINKTYPE, linkType);
            if (ExcuteCommand(
                string.Format("Do you really want to DELETE the Link Type '{0}' in project collection?",
                                linkType),
                command))
                this.InitLinkTypePage();
        }

        private void btnLinks_Export_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.EXPORTLINKTYPE);
                return;
            }

            if(!IsLinkSelected())
                return;

            if (textBoxExportPath.Text.Length == 0)
            {
                MessageBoxWarning("Select Export Path");
                return;
            }

            string linkFullName = gvLink.SelectedRows[0].Cells[0].Value.ToString();
            string linkTypeFilePath = Path.Combine(textBoxExportPath.Text, 
                linkFullName.Replace(" ", string.Empty).Replace(".", string.Empty) + ".xml");

            string command = GenerateComand(WitadminActions.EXPORTLINKTYPE, linkFullName, linkTypeFilePath);
            if (ExcuteCommand(string.Empty, command))
                txtLinkXml.Text = ReadFileToXmlControl(linkTypeFilePath);
        }

        private void btnLinks_ImportFromXml_Click(object sender, EventArgs e)
        {
            if (ImportFromXMLHelper(WitadminActions.IMPORTLINKTYPE, "LinkType", txtLinkXml.Text))
            {
                InitLinkTypePage();
            }
        }

        private void btnLinks_ImporFromFile_Click(object sender, EventArgs e)
        {
            if (ImportFromFileHelper(WitadminActions.IMPORTLINKTYPE, "LinkType"))
            {
                this.InitLinkTypePage();
            }
        }
        
        private void btnLinks_Deactivate_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.DEACTIVATELINKTYPE);
                return;
            }

            if (!IsLinkSelected())
                return;

            string linkType = this.gvLink.SelectedRows[0].Cells[0].Value.ToString();
            string command = GenerateComand(WitadminActions.DEACTIVATELINKTYPE,linkType );
            if (ExcuteCommand(
                string.Format("Do you really want to DEACTIVATE the Link Type '{0}' in project collection?",
                                linkType),
                command))

                InitLinkTypePage();
        }

        private void btnLinks_Reactivate_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.REACTIVATELINKTYPE);
                return;
            }

            if (!IsLinkSelected())
                return;

            string linkType = this.gvLink.SelectedRows[0].Cells[0].Value.ToString();
            string command = GenerateComand(WitadminActions.REACTIVATELINKTYPE, linkType);
            if (ExcuteCommand(
                string.Format("Do you really want to REACTIVATE the Link Type '{0}' in project collection?",
                                linkType),
                command))
                this.InitLinkTypePage();
        }
        #endregion

        #region GlobalList
        TreeNode mySelectedNode = new TreeNode();

        private void treeView1_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label != null)
            {
                if (e.Label.Length > 0)
                {
                    if (mySelectedNode.Parent == treeViewGlobalLists.Nodes[0] && mySelectedNode.Tag == (object)"Server")
                    {
                        e.CancelEdit = true;
                        MessageBoxError("Cannot rename lists that have been saved to the server.");
                    }
                    else if (e.Label.IndexOfAny(new char[] { '@', ',', '!', '\\' }) == -1)
                    {
                        // Stop editing without canceling the label change.
                        e.Node.EndEdit(false);
                    }
                    else
                    {
                        /* Cancel the label edit action, inform the user, and 
                           place the node in edit mode again. */
                        e.CancelEdit = true;
                        MessageBoxError("Invalid tree node label.\n" +
                           "The invalid characters are: '@', ',', '!', '\'",
                           "Node Label Edit");
                        e.Node.BeginEdit();
                    }
                }
                else
                {
                    /* Cancel the label edit action, inform the user, and 
                       place the node in edit mode again. */
                    e.CancelEdit = true;
                    MessageBoxError("Invalid tree node label.\nThe label cannot be blank",
                       "Node Label Edit");
                    e.Node.Parent.Expand();
                    e.Node.Expand();
                    e.Node.BeginEdit();

                }
                this.treeViewGlobalLists.LabelEdit = false;
            }
        }

        private void treeView1_MouseDown(object sender, MouseEventArgs e)
        {
            mySelectedNode = treeViewGlobalLists.GetNodeAt(e.X, e.Y);
        }
        private void NewItem()
        {
            TreeNode targetNode;

            if (mySelectedNode.Parent != null && mySelectedNode.Parent == treeViewGlobalLists.Nodes[0])
            {
                targetNode = mySelectedNode;
            }
            else if (mySelectedNode.Parent != null)
            {
                targetNode = mySelectedNode.Parent;
            }
            else if (mySelectedNode == treeViewGlobalLists.Nodes[0])
            {
                targetNode = mySelectedNode;
            }
            else
            {
                MessageBoxWarning("Select a target list");
                return;
            }

            TreeNode newNode = new TreeNode("New Item");

            mySelectedNode = newNode;
            targetNode.Nodes.Add(newNode);
            
            this.treeViewGlobalLists.LabelEdit = true;
            
            ExpandParents(newNode);
            newNode.BeginEdit();
        }

        private void ExpandParents(TreeNode node)
        {
            while (node.Parent != null)
            {
                node.Parent.Expand();
                node = node.Parent;
            }

        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (treeViewGlobalLists.Nodes[0] != mySelectedNode) //Can't delete root node
            {
                if (mySelectedNode.Level == 1)
                {
                    MessageBoxError("Cannon delete Root Node, Use the Delete Button");
                    return;
                }
                if (mySelectedNode.Parent == treeViewGlobalLists.Nodes[0] && mySelectedNode.Tag == (object)"server")
                {
                    MessageBoxError("Cannot delete lists that have been saved to the server");
                }
                else
                {
                    treeViewGlobalLists.Nodes.Remove(mySelectedNode);
                }
            }
            else
            {
                MessageBoxError("Cannot delete the root node");
            }
        }
        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewItem();
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (mySelectedNode != null && mySelectedNode.Parent != null)
            {
                treeViewGlobalLists.SelectedNode = mySelectedNode;
                treeViewGlobalLists.LabelEdit = true;
                if (!mySelectedNode.IsEditing)
                {
                    mySelectedNode.BeginEdit();
                }
            }
            else
            {
                MessageBoxError("No tree node selected or selected node is a root node\n" +
                   "Editing of root nodes is not allowed", "Invalid selection");
            }
        }

        private void expandAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            treeViewGlobalLists.ExpandAll();
            treeViewGlobalLists.SelectedNode = treeViewGlobalLists.Nodes[0];
        }

        private void collapseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            treeViewGlobalLists.CollapseAll();
            treeViewGlobalLists.Nodes[0].Expand();
        }

        private XmlDocument CreateXMLDocFromTree()
        {

            XmlDocument doc = new XmlDocument();
            XmlElement rootNode = doc.CreateElement("gl", "GLOBALLISTS", "http://schemas.microsoft.com/VisualStudio/2005/workitemtracking/globallists");


            foreach (TreeNode treeList in treeViewGlobalLists.Nodes[0].Nodes)
            {
                XmlElement listNode = doc.CreateElement("GLOBALLIST");
                XmlAttribute listNodeName = doc.CreateAttribute("name");
                listNodeName.Value = treeList.Text;
                listNode.Attributes.Append(listNodeName);
                foreach (TreeNode treeListItem in treeList.Nodes)
                {
                    XmlElement listItemNode = doc.CreateElement("LISTITEM");
                    XmlAttribute listItemNodeValue = doc.CreateAttribute("value");
                    listItemNodeValue.Value = treeListItem.Text;
                    listItemNode.Attributes.Append(listItemNodeValue);
                    listNode.AppendChild(listItemNode);
                } rootNode.AppendChild(listNode);

            }
            doc.AppendChild(rootNode);
            return doc;
        }

        private void btnGlobalListImportFromList_Click(object sender, EventArgs e)
        {
            textBoxCommand.Text = "This command uses workitemStore.ImportGlobalLists.";
            textBoxResult.Text = string.Empty;

            if (radioButtonMode_GetHelpOnCommand.Checked)
                return;

            Hourglass(true);

            if (treeViewGlobalLists.Nodes[0].Nodes.Count == 0)
            {
                textBoxResult.Text = "List of global lists is empty. Please add some global lists before importing.";
                return;
            }

            XmlDocument doc = CreateXMLDocFromTree();
            
            XmlWriterSettings ws = new XmlWriterSettings();
            ws.Indent = true;
            StringBuilder output = new StringBuilder();
            using (XmlWriter writer = XmlWriter.Create(output, ws))
            {
                doc.WriteContentTo(writer);
            }

            if (radioButtonMode_GenerateOnly.Checked)
                return;

            try
            {
                this.workitemStore.ImportGlobalLists(doc.DocumentElement);
                textBoxResult.Text= "Global List Imported";
            }
            catch (Exception exp)
            {
                textBoxResult.Text = "Could not save the list. Error:" + exp.ToString();
                return;
            }

            Hourglass(false);

            this.InitGloballistPage();

            txtGLXml.Text = output.ToString();
        }

        private void btnGLExport_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.EXPORTGLOBALLIST);
                return;
            }

            if (textBoxExportPath.Text.Length == 0)
            {
                MessageBoxWarning("Select Export Path");
                return;
            }

            string globlListFilePath = Path.Combine(textBoxExportPath.Text,
                workitemStore.TeamProjectCollection.Name.Replace("\\", "_") + ".xml");

            string command = GenerateComand(WitadminActions.EXPORTGLOBALLIST, globlListFilePath);
            if(ExcuteCommand(string.Empty, command))
                txtGLXml.Text = ReadFileToXmlControl(globlListFilePath);
        }
        private void btnGLImportFromXml_Click(object sender, EventArgs e)
        {
            if (ImportFromXMLHelper(WitadminActions.IMPORTGLOBALLIST, "Global List", txtGLXml.Text))
            {
                InitGloballistPage();
            }
        }
        private void btnGLImportFromFile_Click(object sender, EventArgs e)
        {
            if (ImportFromFileHelper(WitadminActions.IMPORTGLOBALLIST, "Global List"))
            {
                InitGloballistPage();
            }
        }
        private void btnGLDelete_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.DESTROYGLOBALLIST);
                return;
            }

            if (this.treeViewGlobalLists.SelectedNode == null)
            {
                MessageBoxWarning("Select list's root node");
                return;
            }
            if (this.treeViewGlobalLists.SelectedNode.Level != 1)
            {
                MessageBoxError("You can only can delete list's root node");
                return;
            }

            string globalList = this.treeViewGlobalLists.SelectedNode.Text;
            string command = GenerateComand(WitadminActions.DESTROYGLOBALLIST, globalList);
            if (ExcuteCommand(
                string.Format("Do you really want to DESTROY the Global List '{0}' in project collection?",
                globalList, listProject.SelectedItem.ToString()),
                command))
                this.InitGloballistPage();
        }

        #endregion

        #region Category
        private void btnCategoriesExport_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.EXPORTCATEGORIES);
                return;
            }

            if (textBoxExportPath.Text.Length == 0)
            {
                MessageBoxWarning("Select Export Path");
                return;
            }

            string project = listProject.SelectedItem.ToString();
            string cateXmlFilePath = Path.Combine(textBoxExportPath.Text, project.Replace(" ", string.Empty) + ".xml");
            string command = GenerateComand(WitadminActions.EXPORTCATEGORIES, cateXmlFilePath);
            if(ExcuteCommand(string.Empty, command))
                txtCategoryXml.Text = ReadFileToXmlControl(cateXmlFilePath);
        }

        private void btnCategoriesImportFromXml_Click(object sender, EventArgs e)
        {
            ImportFromXMLHelper(WitadminActions.IMPORTCATEGORIES, "Categories", txtCategoryXml.Text);
        }

        private void btnCategories_ImportFromFile_Click(object sender, EventArgs e)
        {
            ImportFromFileHelper(WitadminActions.IMPORTCATEGORIES, "Categories");
        }

        #endregion       

        #region Process Configuration

        private void btnProcessConfig_Export_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.EXPORTPROCESSCONFIG);
                return;
            }

            if (textBoxExportPath.Text.Length == 0)
            {
                MessageBoxWarning("Select Export Path");
                return;
            }

            string projectFileName = listProject.SelectedItem.ToString().Replace(" ", string.Empty) + ".xml";
            string xmlFilePath = Path.Combine(textBoxExportPath.Text, projectFileName);
            string command = GenerateComand(WitadminActions.EXPORTPROCESSCONFIG, xmlFilePath);
            if (ExcuteCommand(string.Empty, command))
                txtProcessConfigXml.Text = ReadFileToXmlControl(xmlFilePath);
        }

        private void btnProcessConfig_ImportFromXml_Click(object sender, EventArgs e)
        {
            ImportFromXMLHelper(WitadminActions.IMPORTPROCESSCONFIG, "Process Configuration", txtProcessConfigXml.Text);
        }

        private void btnProcessConfig_ImportFromFile_Click(object sender, EventArgs e)
        {
            ImportFromFileHelper(WitadminActions.IMPORTPROCESSCONFIG, "Process Configuration");
        }

        #endregion   

        #region Generate Command

        enum WitadminActions
        {
            IMPORTWITD,
            IMPORTWITD_V,
            EXPORTWITD,
            DESTROYWITD,
            RENAMEWITD,
            DESTROYWI,
            DELETEFIELD,
            CHANGEFIELD_1,
            CHANGEFIELD_2,
            INDEXFIELD,
            DELETELINKTYPE,
            EXPORTLINKTYPE,
            IMPORTLINKTYPE,
            DEACTIVATELINKTYPE,
            REACTIVATELINKTYPE,
            EXPORTGLOBALLIST,
            IMPORTGLOBALLIST,
            DESTROYGLOBALLIST,
            EXPORTCATEGORIES,
            IMPORTCATEGORIES,
            EXPORTPROCESSCONFIG,
            IMPORTPROCESSCONFIG
        };

        private void GenerateHelpOnComand(WitadminActions action)
        {
            if (!radioButtonMode_GetHelpOnCommand.Checked)
                throw new NotSupportedException();

            string command = string.Empty;

            string getHelpFormat = @"{0} /?";
            switch (action)
            {
                case WitadminActions.EXPORTWITD: command = string.Format(getHelpFormat, "exportwitd"); break;
                case WitadminActions.IMPORTWITD:
                case WitadminActions.IMPORTWITD_V: command = string.Format(getHelpFormat, "importwitd"); break;
                case WitadminActions.DESTROYWITD: command = string.Format(getHelpFormat, "destroywitd"); break;
                case WitadminActions.RENAMEWITD: command = string.Format(getHelpFormat, "renamewitd"); break;
                case WitadminActions.EXPORTCATEGORIES: command = string.Format(getHelpFormat, "exportcategories"); break;
                case WitadminActions.IMPORTCATEGORIES: command = string.Format(getHelpFormat, "importcategories"); break;
                case WitadminActions.DESTROYWI: command = string.Format(getHelpFormat, "destroywi"); break;
                case WitadminActions.DELETEFIELD: command = string.Format(getHelpFormat, "deletefield"); break;
                case WitadminActions.CHANGEFIELD_1:
                case WitadminActions.CHANGEFIELD_2: command = string.Format(getHelpFormat, "changefield"); break;
                case WitadminActions.INDEXFIELD: command = string.Format(getHelpFormat, "indexfield"); break;
                case WitadminActions.DELETELINKTYPE: command = string.Format(getHelpFormat, "deletelinktype"); break;
                case WitadminActions.EXPORTLINKTYPE: command = string.Format(getHelpFormat, "exportlinktype"); break;
                case WitadminActions.IMPORTLINKTYPE: command = string.Format(getHelpFormat, "importlinktype"); break;
                case WitadminActions.DEACTIVATELINKTYPE: command = string.Format(getHelpFormat, "deactivatelinktype"); break;
                case WitadminActions.REACTIVATELINKTYPE: command = string.Format(getHelpFormat, "reactivatelinktype"); break;
                case WitadminActions.EXPORTGLOBALLIST: command = string.Format(getHelpFormat, "exportgloballist"); break;
                case WitadminActions.IMPORTGLOBALLIST: command = string.Format(getHelpFormat, "importgloballist"); break;
                case WitadminActions.DESTROYGLOBALLIST: command = string.Format(getHelpFormat, "destroygloballist"); break;
                case WitadminActions.EXPORTPROCESSCONFIG: command = string.Format(getHelpFormat, "exportprocessconfig"); break;
                case WitadminActions.IMPORTPROCESSCONFIG: command = string.Format(getHelpFormat, "importprocessconfig "); break;
                default:
                    throw new NotSupportedException();
            }

            textBoxCommand.Text = string.Format("witadmin.exe {0}", command);
            ExcuteCommand(string.Empty, command);
        }

        private string GenerateComand(WitadminActions action, params object[] args)
        {
            textBoxResult.Text = string.Empty;
            textBoxCommand.Text = string.Empty;
            string command = string.Empty;

            switch (action)
            {
                case WitadminActions.EXPORTWITD:
                case WitadminActions.IMPORTWITD:
                case WitadminActions.IMPORTWITD_V:
                case WitadminActions.DESTROYWITD:
                case WitadminActions.RENAMEWITD:
                case WitadminActions.DESTROYWI:
                case WitadminActions.EXPORTCATEGORIES:
                case WitadminActions.IMPORTCATEGORIES:
                case WitadminActions.EXPORTPROCESSCONFIG:
                case WitadminActions.IMPORTPROCESSCONFIG:
                    if (this.listProject.SelectedItem == null)
                    {
                        MessageBoxWarning("Project not selected");
                        return string.Empty;
                    }
                    break;
            }

            string collection = this.lblCollection.Text;
            string project = this.listProject.SelectedItem.ToString();

            object[] commonParametersCollection = { collection };
            object[] commonParametersProjectCollection = { collection, project };
            object[] parametersProject = commonParametersCollection.Concat(args).ToArray();
            object[] parametersProjectCollection = commonParametersProjectCollection.Concat(args).ToArray();

            switch (action)
            {
                case WitadminActions.EXPORTWITD:
                    // exportwitd 	/collection: /p: /n: /f:
                    command = string.Format(@"exportwitd /collection:{0} /p:""{1}"" /n:""{2}"" /f:""{3}""", parametersProjectCollection);
                    break;
                case WitadminActions.IMPORTWITD:
                    // importwitd	/collection: /p: /f:
                    command = string.Format(@"importwitd /collection:{0} /p:""{1}"" /f:""{2}""", parametersProjectCollection);
                    break;
                case WitadminActions.IMPORTWITD_V:
                    command = string.Format(@"importwitd /collection:{0} /p:""{1}"" /f:""{2}"" /v", parametersProjectCollection);
                    break;
                case WitadminActions.DESTROYWITD:
                    // destroywitd 	/collection: /p: /n: /noprompt
                    command = string.Format(@"destroywitd /collection:{0} /p:""{1}"" /n:""{2}"" /noprompt", parametersProjectCollection);
                    break;
                case WitadminActions.RENAMEWITD:
                    command = string.Format(@"renamewitd /collection:{0} /p:""{1}"" /n:""{2}"" /new:""{3}"" /noprompt", parametersProjectCollection);
                    break;
                case WitadminActions.EXPORTCATEGORIES:
                    command = string.Format(@"exportcategories /collection:{0} /p:""{1}"" /f:""{2}""", parametersProjectCollection);
                    break;
                case WitadminActions.IMPORTCATEGORIES:
                    command = string.Format(@"importcategories /collection:{0} /p:""{1}"" /f:""{2}""", parametersProjectCollection);
                    break;
                case WitadminActions.EXPORTPROCESSCONFIG:
                    command = string.Format(@"exportprocessconfig /collection:{0} /p:""{1}"" /f:""{2}""", parametersProjectCollection);
                    break;
                case WitadminActions.IMPORTPROCESSCONFIG:
                    command = string.Format(@"importprocessconfig /collection:{0} /p:""{1}"" /f:""{2}""", parametersProjectCollection);
                    break;

                case WitadminActions.DESTROYWI:
                    command = string.Format(@"destroywi /collection:{0} /id:{1} /noprompt", parametersProject);
                    break;
                case WitadminActions.DELETEFIELD:
                    // deletefield	/collection: /n:(referncename) /noprompt
                    command = string.Format(@"deletefield /collection:{0} /n:{1} /noprompt", parametersProject);
                    break;
                case WitadminActions.CHANGEFIELD_1:
                    // changefield	/collection: /n:(referncename) /name:(newname) /syncnamechanges:(true/false) /reportingtype:(dimention,detail,measure,none) /reportingformula:(sum) /noprompt
                    command = string.Format(@"changefield /collection:{0} /n:{1} /name:""{2}"" /noprompt", parametersProject);
                    break;
                case WitadminActions.CHANGEFIELD_2:
                    command = string.Format(@"changefield /collection:{0} /n:{1} /name:""{2}"" /reportingtype:{3} /noprompt", parametersProject);
                    break;
                case WitadminActions.INDEXFIELD:
                    // indexfield	/collection: /n:(referncename) /index:(on/off)
                    command = string.Format(@"indexfield /collection:{0} /n:{1} /index:{2} ", parametersProject);
                    break;
                case WitadminActions.DELETELINKTYPE:
                    //deletelinktype /collection: /n:(linktypename,linkrefname) /noprompt
                    command = string.Format(@"deletelinktype /collection:{0} /n:""{1}"" /noprompt", parametersProject);
                    break;
                case WitadminActions.EXPORTLINKTYPE:
                    // exportlinktype	/collection: /n:(linktypename,linkrefname) /f:
                    command = string.Format(@"exportlinktype /collection:{0} /n:""{1}"" /f:""{2}""", parametersProject);
                    break;
                case WitadminActions.IMPORTLINKTYPE:
                    //importlinktype	/collection: /f: 
                    command = string.Format(@"importlinktype /collection:{0} /f:""{1}"" ", parametersProject);
                    break;
                case WitadminActions.DEACTIVATELINKTYPE:
                    // deactivatelinktype	/collection: /n:
                    command = string.Format(@"deactivatelinktype /collection:{0} /n:""{1}""", parametersProject);
                    break;
                case WitadminActions.REACTIVATELINKTYPE:
                    // reactivatelinktype	/collection: /n:
                    command = string.Format(@"reactivatelinktype /collection:{0} /n:""{1}""", parametersProject);
                    break;
                case WitadminActions.EXPORTGLOBALLIST:
                    command = string.Format(@"exportgloballist /collection:{0} /f:""{1}""", parametersProject);
                    break;
                case WitadminActions.IMPORTGLOBALLIST:
                    //importlinktype	/collection: /f: 
                    command = string.Format(@"importgloballist /collection:{0} /f:""{1}""", parametersProject);
                    break;
                case WitadminActions.DESTROYGLOBALLIST:
                    // destorygloballist	/collection: /n: /noprompt
                    command = string.Format(@"destroygloballist /collection:{0} /n:""{1}"" /noprompt ", parametersProject);
                    break;
                default:
                    throw new NotSupportedException();
            }

            textBoxCommand.Text = string.Format("witadmin.exe {0}", command);
            return command;
        }
        #endregion

        #region Excute Command
        private bool ExcuteCommand(string message, string ApplicationArguments)
        {
            if (radioButtonMode_GenerateOnly.Checked)
                return false;

            if (radioButtonMode_ExecuteCommand.Checked && message.Length > 0)
                if (MessageBox.Show(message, "Caution", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                    return false;

            try
            {
                textBoxResult.Text = string.Empty;

                string ApplicationPath = "witadmin.exe";

                Process ProcessObj = new Process();

                ProcessObj.StartInfo.FileName = ApplicationPath;
                ProcessObj.StartInfo.Arguments = ApplicationArguments;

                ProcessObj.StartInfo.UseShellExecute = false;
                ProcessObj.StartInfo.CreateNoWindow = true;

                ProcessObj.StartInfo.WindowStyle = ProcessWindowStyle.Normal;

                ProcessObj.StartInfo.RedirectStandardOutput = true;
                ProcessObj.StartInfo.RedirectStandardError = true;

                ProcessObj.Start();

                Cursor.Current = Cursors.WaitCursor;
                ProcessObj.WaitForExit();
                Cursor.Current = Cursors.Default;

                string Result = ProcessObj.StandardOutput.ReadToEnd();
                string Error = ProcessObj.StandardError.ReadToEnd();
                bool success = true;

                string resultMsg = "Output:" + Environment.NewLine + Result;
                if (Error.Length > 0)
                {
                    resultMsg = resultMsg + Environment.NewLine + "Error:" + Environment.NewLine + Error;
                    success = false;
                }
                textBoxResult.Text = resultMsg;
                return success;
            }
            catch (Exception ex)
            {
                MessageBoxError(ex.Message);
                return false;
            }
        }
        #endregion

        #region Export/Import
        private void btnExportImport_OpenFile_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Text Files (.xml)|*.xml|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;

            openFileDialog1.Multiselect = false;

            // Process input if the user clicked OK.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxImportFilePath.Text = openFileDialog1.FileName;
                textBoxCommand.Text = String.Empty;
                textBoxResult.Text = String.Empty;
            }
        }

        private void btnExportImport_SelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (textBoxExportPath.Text.Length > 0)
                folderBrowserDialog.SelectedPath = textBoxExportPath.Text;
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxExportPath.Text = folderBrowserDialog.SelectedPath;
            }
        }
        #endregion

        #region Helpers
        private void Hourglass(bool Show)
        {
            if (Show == true)
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            }
            else
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
            return;
        }

        private string ReadFileToXmlControl(string witXmlFilePath)
        {
            if (radioButtonMode_GenerateOnly.Checked)
                return string.Empty;

            // Read File
            StreamReader reader = File.OpenText(witXmlFilePath);
            string input = null;
            string fileContent = null;
            while ((input = reader.ReadLine()) != null)
            {
                fileContent += input + @"
";
            }
            reader.Close();
            return fileContent;
        }


        private bool ImportFromFileHelper(WitadminActions action, string importItem)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(action);
                return false;
            }

            if (textBoxImportFilePath.Text.Length == 0)
            {
                MessageBoxWarning("File Name empty");
                return false;
            }

            string command = GenerateComand(action, textBoxImportFilePath.Text);
            string message = string.Empty;
            switch (action)
            {
                case WitadminActions.IMPORTWITD:
                case WitadminActions.IMPORTCATEGORIES:
                    message = string.Format("Do you really want to IMPORT the {0} to project '{1}' from file '{2}'?",
                    importItem, listProject.SelectedItem.ToString(), textBoxImportFilePath.Text);
                    break;
                default:
                    message = string.Format("Do you really want to IMPORT the {0} from file '{1}'?",
                    importItem, textBoxImportFilePath.Text);
                    break;
            }
            return ExcuteCommand(message, command);
        }

        private bool ImportFromXMLHelper(WitadminActions action, string importItem, string xmlText)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(action);
                return false;
            }

            string witXmlFilePath = SaveXMLToFile(xmlText);
            if (witXmlFilePath.Length == 0)
                return false;

            string command = GenerateComand(action, witXmlFilePath);
            string message = string.Empty;
            switch (action)
            {
                case WitadminActions.IMPORTWITD:
                case WitadminActions.IMPORTCATEGORIES:
                    message = string.Format("Do you really want to IMPORT the {0} to project '{1}' from XML'?",
                    importItem, listProject.SelectedItem.ToString());
                    break;
                default:
                    message = string.Format("Do you really want to IMPORT the {0} from XML?",
                    importItem);
                    break;
            }
            return ExcuteCommand(message,command);
        }
        private string SaveXMLToFile(string xmlText)
        {
            if (xmlText.Length == 0)
            {
                MessageBoxError("XML is empty");
                return string.Empty;
            }

            XmlDocument doc = new XmlDocument();
            try
            {
                doc.LoadXml(xmlText);
            }
            catch (Exception ex)
            {
                MessageBoxError(ex.Message, "Error reading XML");
                return string.Empty;
            }

            string filePath = Path.GetTempFileName();
            filePath = Path.ChangeExtension(filePath, "xml");
            doc.Save(filePath);
            return filePath;
        }

        void MessageBoxWarning(string message)
        {
            MessageBox.Show(message, "Caution", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        void MessageBoxError(string message, string caption=null)
        {
            if (caption == null)
                caption = "Error";
            MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        #endregion
    }
}

