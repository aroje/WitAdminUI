using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using Microsoft.TeamFoundation.Client;
using System.Security;
using System.Diagnostics;
//using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.IO;
using System.Xml;
using System.Drawing.Drawing2D;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace WitAdminTool
{
    public partial class WitAdminToolForm : Form
    {
        TfsTeamProjectCollection tfsTPC;
        WorkItemStore workitemStore;

        #region Initialize
        public WitAdminToolForm()
        {
            InitializeComponent();
            try
            {
                TeamProjectPicker tpp = new TeamProjectPicker();
                if (tpp.ShowDialog() == DialogResult.OK)
                {
                    tfsTPC = tpp.SelectedTeamProjectCollection;

                    workitemStore = (WorkItemStore)tfsTPC.GetService(typeof(WorkItemStore));
                    this.lblCollection.Text = tfsTPC.Uri.AbsoluteUri;

                    InitForm();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                
            }
        }
 
 
        private void InitForm()
        {
            InitProjects();
            InitWorkItemTypePage();
            
            InitFieldPage();
            InitLinkTypePage();
            InitGloballistPage();

            textBoxExportPath.Text = Path.GetTempPath();
        }

        private void InitCommand()
        {
            textBoxResult.Text = string.Empty;
            textBoxCommand.Text = string.Empty;
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                TeamProjectPicker tpp = new TeamProjectPicker();
                if (tpp.ShowDialog() == DialogResult.OK)
                {
                    tfsTPC = tpp.SelectedTeamProjectCollection;

                    workitemStore = (WorkItemStore)tfsTPC.GetService(typeof(WorkItemStore));
                    this.lblCollection.Text = tfsTPC.Uri.AbsoluteUri;

                    this.tabControl_ProjectLevelActions.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
            workitemStore.SyncToCache();
            listWit.Items.Clear();
            txtWitXml.Text = string.Empty;
            
            if (listProject.SelectedItem == null)
            {
                return;
            }
            string project = listProject.SelectedItem.ToString();
            foreach (WorkItemType wit in this.workitemStore.Projects[project].WorkItemTypes)
            {
                listWit.Items.Add(wit.Name);
            }
        }
        private void InitWorkItemPage()
        {
            this.txtWiID.Text = string.Empty;
            
        }
        private void InitFieldPage()
        {
            this.workitemStore.SyncToCache();
            this.listFields.Items.Clear();
            foreach (FieldDefinition fd in this.workitemStore.FieldDefinitions)
            {
                this.listFields.Items.Add(fd.Name + "/" + fd.ReferenceName);
            }
            listFields.SelectedIndex = 0;
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

            this.gvLink.DataSource = dtLik;
            this.gvLink.Rows[0].Selected = true;

            this.txtLinkXml.Text = string.Empty;
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

        private void btnWit_Rename_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.RENAMEWITD);
                return;
            }

            if (listWit.SelectedItem == null)
            {
                MessageBox.Show("Select Work Item Type");
                return;
            }
            if (this.txtNewWitName.Text == string.Empty)
            {
                MessageBox.Show("Input New Work Item Type Name");
                return;
            }

            string command = GenerateComand(WitadminActions.RENAMEWITD, this.listWit.SelectedItem.ToString(), this.txtNewWitName.Text);

            if (ExcuteCommand("Really Rename Work Item Type?", command))
                this.InitWorkItemTypePage();
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
                MessageBox.Show("Select Work Item Type");
                return;
            }

            string command = GenerateComand(WitadminActions.DESTROYWITD, this.listWit.SelectedItem.ToString());

            if (ExcuteCommand("Really Delete Work Item Type?", command))
                this.InitWorkItemTypePage();
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
                MessageBox.Show("Select Work Item Type");
                return;
            }

            if (textBoxExportPath.Text.Length == 0)
            {
                MessageBox.Show("Select Export Path");
                return;
            }

            string witXmlFilePath = textBoxExportPath.Text + Path.DirectorySeparatorChar + this.listWit.SelectedItem.ToString().Replace(" ", string.Empty) + ".xml";
            string command = GenerateComand(WitadminActions.EXPORTWITD, this.listWit.SelectedItem.ToString(), witXmlFilePath);
            ExcuteCommand(string.Empty, command);

            this.txtWitXml.Text = ReadFileToXmlControl(witXmlFilePath);

        }

        private void btnWit_ImportFromXml_Click(object sender, EventArgs e)
        {
            if (ImportFromXMLHelper(WitadminActions.IMPORTWITD, "Work Item Type", txtWitXml.Text))
            {
                this.InitWorkItemTypePage();
            }
        }

        private void btn_WIT_ImportFromFile_Click(object sender, EventArgs e)
        {
            if (ImportFromFileHelper(WitadminActions.IMPORTWITD, "Work Item Type"))
            {
                this.InitWorkItemTypePage();
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
                MessageBox.Show("File Name empty", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    MessageBox.Show("Select Project");
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
                MessageBox.Show("Input Number of Work Item to select");
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
                MessageBox.Show("Input Work Item ID(s)");
                txtWiID.Focus();
                return;
            }

            string command = GenerateComand(WitadminActions.DESTROYWI, this.txtWiID.Text);
            if (ExcuteCommand("Really Delete Work Item(s)?", command))
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
                MessageBox.Show("Check Work Item(s)");
                return;
            }

            string command = GenerateComand(WitadminActions.DESTROYWI, ids);
            if (ExcuteCommand("Really Delete Work Item(s)?", command))
               this.GetWi();
        }
        #endregion

        #region Fields
        private void listFields_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.txtFdName.Text = string.Empty;
            this.txtFdRefName.Text = string.Empty;
            this.txtFdType.Text = string.Empty;

            string seletecItem = this.listFields.SelectedItem.ToString();
            if (!string.IsNullOrEmpty(seletecItem))
            {
                string refName = seletecItem.Split('/')[1];

                FieldDefinition fd = this.workitemStore.FieldDefinitions[refName];
                if (fd != null)
                {
                    txtFdName.Text = fd.Name;
                    txtFdRefName.Text = fd.ReferenceName;
                    txtFdType.Text = fd.FieldType.ToString();
                    this.cbbFdReportType.SelectedItem = fd.ReportingAttributes.Type.ToString().ToLower();
                    if (fd.IsIndexed)
                    {
                        this.cbbIndexYN.SelectedItem = "Y";
                        this.btnFields_Index.Text = "Index Off";
                    }
                    else
                    {
                        this.cbbIndexYN.SelectedItem = "N";
                        this.btnFields_Index.Text = "Index On";
                    }
                }
            }
        }

        private void btnFields_Delete_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.DELETEFIELD);
                return;
            }

            if (this.listFields.SelectedItem == null)
            {
                MessageBox.Show("Select Field");
                return;
            }

            string command = GenerateComand(WitadminActions.DELETEFIELD, txtFdRefName.Text);
            if (ExcuteCommand("Really Delete Field?", command))
                InitFieldPage();
        }

        private void btnFields_Change_Click(object sender, EventArgs e)
        {
            if (radioButtonMode_GetHelpOnCommand.Checked)
            {
                GenerateHelpOnComand(WitadminActions.CHANGEFIELD_1);
                return;
            }

            if (this.listFields.SelectedItem == null)
            {
                MessageBox.Show("Select Field");
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


            if (ExcuteCommand("Really Change Field?", command))
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
            if (ExcuteCommand("Really Index Field?", command))
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
                MessageBox.Show("Select LinkType Row");
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

            string command = GenerateComand(WitadminActions.DELETELINKTYPE, this.gvLink.SelectedRows[0].Cells[0].Value.ToString());
            if (ExcuteCommand("Really Change LinkType?", command))
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
                MessageBox.Show("Select Export Path");
                return;
            }

            string linkFullName = gvLink.SelectedRows[0].Cells[0].Value.ToString();
            string linkTypeFilePath = textBoxExportPath.Text + Path.DirectorySeparatorChar + linkFullName.Replace(" ", string.Empty).Replace(".", string.Empty) + ".xml";

            string command = GenerateComand(WitadminActions.EXPORTLINKTYPE, linkFullName, linkTypeFilePath);
            ExcuteCommand(string.Empty, command);

            this.txtLinkXml.Text = ReadFileToXmlControl(linkTypeFilePath);
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

            string command = GenerateComand(WitadminActions.DEACTIVATELINKTYPE, this.gvLink.SelectedRows[0].Cells[0].Value.ToString());
            if (ExcuteCommand("Really Deactivate LinkType?", command))
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

            string command = GenerateComand(WitadminActions.REACTIVATELINKTYPE, this.gvLink.SelectedRows[0].Cells[0].Value.ToString());
            if (ExcuteCommand("Really Reactivate LinkType?", command))
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
                        MessageBox.Show("Cannot rename lists that have been saved to the server.");
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
                        MessageBox.Show("Invalid tree node label.\n" +
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
                    MessageBox.Show("Invalid tree node label.\nThe label cannot be blank",
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
                MessageBox.Show("Select a target list");
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
                    MessageBox.Show("Cannon delete Root Node, Use the Delete Button");
                    return;
                }
                if (mySelectedNode.Parent == treeViewGlobalLists.Nodes[0] && mySelectedNode.Tag == (object)"server")
                {
                    MessageBox.Show("Cannot delete lists that have been saved to the server");
                }
                else
                {
                    treeViewGlobalLists.Nodes.Remove(mySelectedNode);
                }
            }
            else
            {
                MessageBox.Show("Cannot delete the root node");
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
                MessageBox.Show("No tree node selected or selected node is a root node\n" +
                   "Editing of root nodes is not allowed", "Invalid selection");
            }
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

            this.txtGLXml.Text = output.ToString();
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
                MessageBox.Show("Select Export Path");
                return;
            }

            string globlListFilePath = textBoxExportPath.Text + Path.DirectorySeparatorChar + 
                workitemStore.TeamProjectCollection.Name.Replace("\\", "_") + ".xml";

            string command = GenerateComand(WitadminActions.EXPORTGLOBALLIST, globlListFilePath);
            ExcuteCommand(string.Empty, command);

            this.txtGLXml.Text = ReadFileToXmlControl(globlListFilePath);
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
                MessageBox.Show("Select list's root node");
                return;
            }
            if (this.treeViewGlobalLists.SelectedNode.Level != 1)
            {
                MessageBox.Show("You can only can delete list's root node");
                return;
            }

            string command = GenerateComand(WitadminActions.DESTROYGLOBALLIST, this.treeViewGlobalLists.SelectedNode.Text);
            if (ExcuteCommand("Really Destroy Global List?", command))
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
                MessageBox.Show("Select Export Path");
                return;
            }

            string project = listProject.SelectedItem.ToString();
            string cateXmlFilePath = textBoxExportPath.Text + Path.DirectorySeparatorChar + project.Replace(" ", string.Empty) + ".xml";
            string command = GenerateComand(WitadminActions.EXPORTCATEGORIES, cateXmlFilePath);
            ExcuteCommand(string.Empty, command);

            this.txtCateXml.Text = ReadFileToXmlControl(cateXmlFilePath);
        }

        private void btnCategoriesImportFromXml_Click(object sender, EventArgs e)
        {
            ImportFromXMLHelper(WitadminActions.IMPORTCATEGORIES, "Global List", txtCateXml.Text);
        }

        private void btnCategories_ImportFromFile_Click(object sender, EventArgs e)
        {
            ImportFromFileHelper(WitadminActions.IMPORTCATEGORIES, "Categories");
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
            IMPORTCATEGORIES
        };

        private void GenerateHelpOnComand(WitadminActions action)
        {
            if (!radioButtonMode_GetHelpOnCommand.Checked)
                throw new NotSupportedException();

            string command = string.Empty;

            string getHelpFormat= @"{0} /?";
            switch (action)
            {
                case WitadminActions.EXPORTWITD:            command = string.Format(getHelpFormat, "exportwitd"); break;
                case WitadminActions.IMPORTWITD:            
                case WitadminActions.IMPORTWITD_V:          command = string.Format(getHelpFormat, "importwitd"); break;
                case WitadminActions.DESTROYWITD:           command = string.Format(getHelpFormat, "destroywitd");break;
                case WitadminActions.RENAMEWITD:            command = string.Format(getHelpFormat, "renamewitd");break;
                case WitadminActions.EXPORTCATEGORIES:      command = string.Format(getHelpFormat, "exportcategories");break;
                case WitadminActions.IMPORTCATEGORIES:      command = string.Format(getHelpFormat, "importcategories");break;
                case WitadminActions.DESTROYWI:             command = string.Format(getHelpFormat, "destroywi"); break;
                case WitadminActions.DELETEFIELD:           command = string.Format(getHelpFormat, "deletefield");break;
                case WitadminActions.CHANGEFIELD_1: 
                case WitadminActions.CHANGEFIELD_2:         command = string.Format(getHelpFormat, "changefield");break;
                case WitadminActions.INDEXFIELD:            command = string.Format(getHelpFormat, "indexfield"); break;
                case WitadminActions.DELETELINKTYPE:        command = string.Format(getHelpFormat, "deletelinktype"); break;
                case WitadminActions.EXPORTLINKTYPE:        command = string.Format(getHelpFormat, "exportlinktype"); break;
                case WitadminActions.IMPORTLINKTYPE:        command = string.Format(getHelpFormat, "importlinktype"); break;
                case WitadminActions.DEACTIVATELINKTYPE:    command = string.Format(getHelpFormat, "deactivatelinktype"); break;
                case WitadminActions.REACTIVATELINKTYPE:    command = string.Format(getHelpFormat, "reactivatelinktype"); break;
                case WitadminActions.EXPORTGLOBALLIST:      command = string.Format(getHelpFormat, "exportgloballist"); break;
                case WitadminActions.IMPORTGLOBALLIST:      command = string.Format(getHelpFormat, "importgloballist"); break;
                case WitadminActions.DESTROYGLOBALLIST:     command = string.Format(getHelpFormat, "destroygloballist"); break;
                default:
                    throw new NotSupportedException();
            }

            textBoxCommand.Text = string.Format("witadmin.exe {0}", command);
            ExcuteCommand(string.Empty, command);
        }

        private string GenerateComand(WitadminActions action, params object[] args)
        {
            InitCommand();
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
                    if (this.listProject.SelectedItem == null)
                    {
                        MessageBox.Show("Project not selected", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    command = string.Format(@"exportcategories /collection:{0} /p:""{1}"" /f:{2}", parametersProjectCollection);
                    break;
                case WitadminActions.IMPORTCATEGORIES:
                    command = string.Format(@"importcategories /collection:{0} /p:""{1}"" /f:{2}", parametersProjectCollection);
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

                string resultMsg = "Output:" + Environment.NewLine + Result;
                if (Error.Length > 0)
                {
                    resultMsg = resultMsg + Environment.NewLine + "Error:" + Environment.NewLine + Error;
                }
                textBoxResult.Text = resultMsg;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                MessageBox.Show("File Name empty", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            string command = GenerateComand(action, textBoxImportFilePath.Text);
            return ExcuteCommand(string.Format("Really Import {0}?", importItem), command);
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
            return ExcuteCommand(string.Format("Really Import {0}?", importItem), command);
        }
        private string SaveXMLToFile(string xmlText)
        {
            if (xmlText.Length == 0)
            {
                MessageBox.Show("XML is empty");
                return string.Empty;
            }

            XmlDocument doc = new XmlDocument();
            try
            {
                doc.LoadXml(xmlText);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error reading XML", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return string.Empty;
            }

            string filePath = Path.GetTempFileName();
            filePath = Path.ChangeExtension(filePath, "xml");
            doc.Save(filePath);
            return filePath;
        }
        #endregion
    }
}
        
