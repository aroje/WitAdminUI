**TFS WitAdminUI** is a simple wrapper around the	**_witadmin_** command-line tool offering various features.

[Witadmin](https://msdn.microsoft.com/en-us/library/dd236914.aspx) is a command-line tool which comes with Visual Studio. As described in [MSDN](https://msdn.microsoft.com/en-us/library/dd236914.aspx), “_By using the witadmin command-line tool, you can create, delete, import, and export objects such as categories, global lists, global workflow, types of links, and types of work items. You can also permanently delete types of work item types, and you can delete, list, or change the attributes of fields in work item._”

When working on customizing a process template, you need to perform several **_witadmin_** actions multiple times, and since it is a command line tool, it gets pretty bothersome composing the long commands over and over again with poor support for copy-paste actions. **TFS WitAdminUI** helps you by automatically generating commands based on the project collection and team project you are connected to. It allows you to preview generated commands with parameters for several **_witadmin_** actions, get help on those actions and execute them against the selected project collection and team project.

### Supported Versions

**TFS WitAdminUI** currently supports TFS versions:
- 2017 and all updates
- 2015
- 2013

**TFS WitAdminUI** currently supports VSTS in the following way:

"_With witadmin, you can modify XML definition files to support the On-premises XML process model. For Hosted XML and Inheritance process models, you can only use witadmin commands to list information._" [MSDN](https://msdn.microsoft.com/en-us/library/dd236914.aspx)

### Documentation

Documentation is available at [https://aroje.github.io/](https://aroje.github.io/).
