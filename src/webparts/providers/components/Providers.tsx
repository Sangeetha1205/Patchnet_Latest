import { IconButton } from "@fluentui/react/lib/Button";
import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/security/web";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-users/web";
import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";
import "@pnp/sp/webs";
import "alertifyjs";
import { Announced } from "office-ui-fabric-react/lib/Announced";
import {
  DefaultButton,
  PrimaryButton,
} from "office-ui-fabric-react/lib/Button";
import {
  ChoiceGroup,
  IChoiceGroupOption,
  IChoiceGroupStyles,
} from "office-ui-fabric-react/lib/ChoiceGroup";
import {
  CommandBar,
  ICommandBarStyles,
} from "office-ui-fabric-react/lib/CommandBar";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsListStyles,
  Selection,
} from "office-ui-fabric-react/lib/DetailsList";
import {
  Dialog,
  DialogFooter,
  DialogType,
  IDialogStyles,
} from "office-ui-fabric-react/lib/Dialog";
import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { Link } from "office-ui-fabric-react/lib/Link";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import {
  IStackProps,
  IStackStyles,
  IStackTokens,
  Stack,
} from "office-ui-fabric-react/lib/Stack";
import {
  mergeStyles,
  mergeStyleSets,
} from "office-ui-fabric-react/lib/Styling";
import {
  ITextFieldStyles,
  TextField,
} from "office-ui-fabric-react/lib/TextField";
import { getId } from "office-ui-fabric-react/lib/Utilities";
import * as React from "react";
import "../../../ExternalRef/CSS/alertify.min.css";
import "../../../ExternalRef/CSS/style.css";
import { IProvidersProp } from "./IProvidersProps";
import styles from "./Providers.module.scss";

// import * as Excel from "exceljs/dist/exceljs.min.js";
//import * as Excel from "../../../../node_modules/exceljs/dist/exceljs.min.js";
import * as XLSX from 'xlsx';
import * as FileSaver from "file-saver";

var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

const exampleChildClass = mergeStyles({
  display: "block",
  marginBottom: "0",
});

const textFieldStyles: Partial<ITextFieldStyles> = {
  root: { maxWidth: "100%", fontFamily: "Poppins, sans-serif" },
};

const currentYear = new Date().getFullYear();

const fileId = getId("anInput");

export interface ILocalProvidersProp {
  Id: number;
  Title: string;
  LegalName: string;
  ProviderID: string;
  TemplateType: string;
  ContractId: string;
  Users: string;
  IsDeleted: boolean;
  Logs: string;
}

export interface IDetailsListBasicExampleState {
  items: ILocalProvidersProp[];
  allItems: ILocalProvidersProp[];
  selectionDetails: string;
  hideDialog: boolean;
  hideDeleteDialog: boolean;

  editUsers: string;
  providerName: string;
  folders: IDropdownOption[];
  subFolders: IDropdownOption[];
  cols: [];
  rows: [];
  AllUsers: any[];
  formData: {
    ProviderID: string;
    Title: string;
    LegalName: string;
    ContractId: string;
    TemplateType: string;
    Users: string;
    Id: number;
    IsDeleted: boolean;
    Logs: string;
  };
  fileName: "";
}

var listUrl = "";

export default class Providers extends React.Component<
  IProvidersProp,
  IDetailsListBasicExampleState
> {



  private _selection: Selection;
  private _columns: IColumn[];


  newAddedUsers = [];
  deletedUsers = [];
  allUploadFolders = [];

  selUsers = [];
  allUsers = [];
  fileObj = null;
  rootFolder = "FY ";
  to = '';


  templateTypes = [
    {
      key: "Contract Providers",
      text: "Contract Provider",
    },
    {
      key: "Agreement Providers",
      text: "Agreement Provider",
    },
  ];


  contributePermission = null;
  readPermission = null;
  currentUser = null;
  userDetails = [];
  azureGuestUsers = [];

  loginNamePrefix = "i:0#.f|membership|";
  loginNameSuffix = "#ext#@browardbehavioralhc.onmicrosoft.com";
  visitorsGroupName = "BBHC Provider SharePoint Viewers";
  configurationList = "Configuration";
  processYear = currentYear;



  constructor(props: IProvidersProp) {
    super(props);
    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });

    alertify.set("notifier", "position", "top-right");

    listUrl = this.props.currentContext.pageContext.web.absoluteUrl;
    var siteindex = listUrl.toLocaleLowerCase().indexOf("sites");
    listUrl = listUrl.substr(siteindex - 1) + "/Lists/";

    var that = this;

    sp.web
      .getList(listUrl + this.configurationList)
      .items.select("Title")
      .filter("Title eq '" + currentYear + "'")
      .get()
      .then((res) => {
        if (res.length == 0) {
          that.processYear = currentYear - 1;
        } else {
          that.processYear = currentYear;
        }
      });

    sp.web.roleDefinitions
      .getByName("Read")
      .get()
      .then(function (res) {
        that.readPermission = res.Id;
      });

    sp.web.roleDefinitions
      .getByName("Read-Only-Upload")
      .get()
      .then(function (res) {
        that.contributePermission = res.Id;
      });

    sp.web
      .getList(listUrl + "DocumentType")
      .items.select("Title", "TemplateType")
      .get()
      .then((res) => {
        that.allUploadFolders = [];
        for (let u = 0; u < res.length; u++) {
          that.allUploadFolders.push({
            Title: res[u].Title,
            TemplateType: res[u].TemplateType,
          });
        }
      });
    this.currentUser = sp.web.currentUser();

    this._selection = new Selection({
      onSelectionChanged: () =>
        this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    this._columns = [
      // { key: 'column1', name: 'Id', fieldName: 'Id', minWidth: 100, maxWidth: 200, isResizable: true },
      {
        key: "column1",
        name: "Provider Name",
        fieldName: "Title",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
      },
      {
        key: "column2",
        name: "Legal Name",
        fieldName: "LegalName",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
      },
      {
        key: "column3",
        name: "Provider ID",
        fieldName: "ProviderID",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column4",
        name: "Template Type",
        fieldName: "TemplateType",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column5",
        name: "Contract Id",
        fieldName: "ContractId",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
    ];

    that.state = {
      items: [],
      allItems: [],
      selectionDetails: "",
      hideDialog: true,
      hideDeleteDialog: true,
      editUsers: "",

      providerName: "",
      folders: [],
      subFolders: [],
      cols: [],
      rows: [],
      AllUsers: [""],
      formData: {
        ProviderID: "",
        Title: "",
        LegalName: "",
        ContractId: "",
        TemplateType: "Contract Providers",
        Users: "",
        Id: 0,
        IsDeleted: false,
        Logs: "",
      },
      fileName: "",
    };
    this.loadTableData();
  }

  loadTableData() {
    var that = this;
    sp.web
      .getList(listUrl + "ProviderDetails")
      .items.orderBy("Id", false)
      .select(
        "Title",
        "LegalName",
        "ProviderID",
        "TemplateType",
        "ContractId",
        "Id",
        "Users",
        "IsDeleted",
        "Logs"
      )
      .get()
      .then(function (data) {
        var allItems = that.state.items;
        allItems = [];
        for (let index = 0; index < data.length; index++) {
          const element = data[index];
          if (!element.IsDeleted) {
            allItems.push({
              Id: element.Id,
              Title: element.Title,
              LegalName: element.LegalName,
              ProviderID: element.ProviderID,
              TemplateType: element.TemplateType,
              ContractId: element.ContractId,
              Users: element.Users,
              Logs: element.Logs,
              IsDeleted: element.IsDeleted,
            });
          }
        }
        that.setState({
          items: allItems,
          allItems: allItems,
          selectionDetails: that._getSelectionDetails(),
        });
      });
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return "No items selected";
      case 1:
        return "1 item selected";
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onFilter = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ): void => {
    this.setState({
      items: text
        ? this.state.allItems.filter(
          (i) => i.Title.toLowerCase().indexOf(text) > -1
        )
        : this.state.allItems,
    });
  };

  private _onItemInvoked = (item: ILocalProvidersProp): void => {
    alert(`Item invoked: ${item.Title}`);
  };

  private editItem(element) {
    var id = element.target.id;
    var data = this.state.allItems.filter((c) => c.Id == id);
    var formData = this.state.formData;
    var allUsers = this.state.AllUsers;
    formData = {
      Id: parseInt(id),
      Title: data[0].Title,
      ProviderID: data[0].ProviderID,
      LegalName: data[0].LegalName,
      ContractId: data[0].ContractId,
      TemplateType: data[0].TemplateType,
      Users: "",
      IsDeleted: false,
      Logs: data[0].Logs,
    };
    allUsers = [];
    var susers = data[0].Users.split(";");
    for (let index = 0; index < susers.length; index++) {
      if (susers[index]) {
        allUsers.push(susers[index]);
      }
    }
    this.setState({
      editUsers: data[0].Users,
      AllUsers: allUsers,
      formData: formData,
      hideDialog: false,
    });
  }

  hideDialog() {
    this.setState({ hideDialog: true });
  }

  providerNameChange = (event) => {
    this.setState({ providerName: event.target.value });
    this.inputChangeHandler(event);
  };

  inputChangeHandler(e) {
    let formData = this.state.formData;
    formData[e.target.name] = e.target.value;
    this.setState({
      formData,
    });
  }

  processInputProvider = () => {
    var formData = this.state.formData;
    if (!formData.ProviderID) {
      alertify.error("Provider ID is required");
      return;
    }
    if (!formData.Title) {
      alertify.error("Provider name is required");
      return;
    }
    if (!formData.ContractId) {
      alertify.error("Contract Id is required");
      return;
    }
    if (!formData.LegalName) {
      alertify.error("Legal name is required");
      return;
    }

    for (let index = 0; index < this.state.AllUsers.length; index++) {
      const user = this.state.AllUsers[index];
      if (!user) {
        alertify.error(user + " is not a valid user");
        return;
      }
    }

    this.checkInvitation(0, true);
  };

  checkInvitation = async (index, valid) => {
    var email = this.state.AllUsers[index];
    var userExists;
    await sp.web.siteGroups
      .getByName(this.visitorsGroupName)
      .users.get()
      .then((results) => {
        userExists = results.filter((result) => {
          return result.Email.toLowerCase() == email.toLowerCase();
        });
        if (userExists.length == 0) {
          valid = false;
        }
      });

    if (!valid) {
      alertify.error(email + " is not part of BBHC yet"); // Alertify size fix
      return;
    }
    index = index + 1;
    if (index < this.state.AllUsers.length) {
      this.checkInvitation(index, valid);
    } else {
      this.otherProcess();
    }
  };

  otherProcess = () => {
    var newlyAdded = [];
    if (this.state.editUsers) {
      var existingUsers = this.state.editUsers.split(";");
      var newUsers = this.state.AllUsers;
      for (let index = 0; index < newUsers.length; index++) {
        if (newUsers[index]) {
          var exist = existingUsers.filter((c) => c == newUsers[index]);
          if (exist.length == 0) {
            newlyAdded.push(newUsers[index]);
          }
        }
      }

      var valid = true;

      if (valid) {
        if (newlyAdded.length > 0) {
          this.sendMailToUsers(newlyAdded);
        }
        this.mainProcess();
      }
    } else {
      this.sendMailToUsers(this.state.AllUsers);
      this.mainProcess();
    }
  };

  sendMailToUsers = (to) => {
    var that = this;
    var filepath = that.props.currentContext.pageContext.web.absoluteUrl;
    const emailProps: IEmailProperties = {
      To: to,
      Subject: "BBHC Provider Library has been shared with you",
      Body: this.props.providerAssignedHTML,
      AdditionalHeaders: {
        "content-type": "text/html",
      },
    };
    sp.utility.sendEmail(emailProps);
  };

  mainProcess = () => {
    var formData = this.state.formData;
    var that = this;
    that.userDetails = [];
    this.newAddedUsers = [];
    this.deletedUsers = [];

    for (let index = 0; index < this.state.AllUsers.length; index++) {
      const user = this.state.AllUsers[index];
      if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(user)) {
        formData.Users = formData.Users + user + ";";
        sp.web.siteUsers
          .getByEmail(user)
          .get()
          .then(function (data) {
            that.userDetails.push({
              Id: data.Id,
              Email: user,
            });
          });
      } else {
        alertify.error("User " + (index + 1) + " not valid");
        return;
      }
    }

    if (this.state.editUsers) {
      var existingUsers = this.state.editUsers.split(";");
      var newUsers = this.state.formData.Users.split(";");

      for (let index = 0; index < newUsers.length; index++) {
        if (newUsers[index]) {
          var exist = existingUsers.filter((c) => c == newUsers[index]);
          if (exist.length == 0) {
            this.newAddedUsers.push(newUsers[index]);
            that.setpermissionfornewuser(
              "TemplateLibrary/" + that.state.formData.TemplateType,
              newUsers[index],
              true
            );
            that.setpermissionfornewuser(
              "TemplateLibrary/" + that.state.formData.TemplateType,
              newUsers[index],
              true
            );
            // that.setpermissionformaintemplate("TemplateLibrary/" + this.state.formData.TemplateType, newUsers[index]);
          }
        }
      }

      var providerData = {
        Title: formData.Title,
        ContractId: formData.ContractId,
        Users: "",
      };

      for (let j = 0; j < existingUsers.length; j++) {
        if (existingUsers[j]) {
          var removeuser = newUsers.filter((c) => c == existingUsers[j]);
          if (removeuser.length == 0) {
            this.deletedUsers.push(existingUsers[j]);
            providerData.Users = providerData.Users + existingUsers[j] + ";";
            that.setpermissionfornewuser(
              "TemplateLibrary/" + that.state.formData.TemplateType,
              existingUsers[j],
              false
            );
          }
        }
      }
      if (providerData.Users) {
        this.removerMainFolderUserPermission(providerData);
      }
    }

    this.setState({ formData: formData });
    this.addToList(currentYear, this.state.formData);
  };

  setpermissionfornewuser(folderPath, user, addpermission) {
    var reacthandler = this;
    sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.setpermission(0, data, user, addpermission);
        }
      });
  }

  setpermission(index, data, user, addpermission) {
    var reacthandler = this;
    var clonedUrl = data[index].ServerRelativeUrl;
    var url = clonedUrl.replace(
      this.props.currentContext.pageContext.web.serverRelativeUrl + "/",
      ""
    );
    const spHttpClient: SPHttpClient = this.props.currentContext.spHttpClient;

    // var contract = this.state.formData.ContractId.substr(
    //   this.state.formData.ContractId.length - 2,
    //   2
    // );

    // var nextyear = parseInt(contract) + 1;
    // var currentyearprefix = currentYear.toString().substr(0, 2);
    var yearfolder =
      reacthandler.processYear +
      reacthandler.to +
      (reacthandler.processYear + 1) +
      "/" +
      this.state.formData.Title;

    var providerFolder = reacthandler.rootFolder + yearfolder;
    var mainTemplateFolder =
      "TemplateLibrary/" + reacthandler.state.formData.TemplateType;
    url = url.replace(mainTemplateFolder, providerFolder);

    url = url.replace(" - Upload", "");

    var queryUrl =
      this.props.currentContext.pageContext.web.absoluteUrl +
      "/_api/web/GetFolderByServerRelativeUrl(" +
      "'" +
      url +
      "'" +
      ")/ListItemAllFields/breakroleinheritance(false)";
    const spOpts: ISPHttpClientOptions = {};

    sp.web.siteUsers
      .getByEmail(user)
      .get()
      .then(function (userdata) {
        spHttpClient
          .post(queryUrl, SPHttpClient.configurations.v1, spOpts)
          .then((response: SPHttpClientResponse) => {
            if (response.ok) {
              var permission = reacthandler.readPermission;
              var sdata = clonedUrl.split("/");

              var folderFilter = reacthandler.allUploadFolders.filter(
                (c) =>
                  c.TemplateType == reacthandler.state.formData.TemplateType
              );
              var found = false;
              for (let l = 0; l < folderFilter.length; l++) {
                const fold = folderFilter[l].Title.split(" - ");
                if (fold[fold.length - 1] == sdata[sdata.length - 1]) {
                  found = true;
                  break;
                }
              }

              if (found) {
                permission = reacthandler.contributePermission;
              }

              var postUrl =
                reacthandler.props.currentContext.pageContext.web.absoluteUrl +
                "/_api/web/GetFolderByServerRelativeUrl(" +
                "'" +
                url +
                "'" +
                ")/ListItemAllFields/roleassignments/removeroleassignment(principalid=" +
                userdata.Id +
                ",roledefid=" +
                permission +
                ")";
              if (addpermission) {
                postUrl =
                  reacthandler.props.currentContext.pageContext.web
                    .absoluteUrl +
                  "/_api/web/GetFolderByServerRelativeUrl(" +
                  "'" +
                  url +
                  "'" +
                  ")/ListItemAllFields/roleassignments/addroleassignment(principalid=" +
                  userdata.Id +
                  ",roledefid=" +
                  permission +
                  ")";
              }
              spHttpClient
                .post(postUrl, SPHttpClient.configurations.v1, spOpts)
                .then((response: SPHttpClientResponse) => {
                  if (response.ok) {
                  }
                });
            }
          });
      });
    reacthandler.setpermissionfornewuser(
      data[index].ServerRelativeUrl,
      user,
      addpermission
    );
    index = index + 1;
    if (index < data.length) {
      reacthandler.setpermission(index, data, user, addpermission);
    }
  }

  // setpermissionformaintemplate(url, useremail) {
  //   var that = this;
  //   sp.web
  //     .getFolderByServerRelativePath(url)
  //     .folders.get()
  //     .then(function (data) {
  //       for (let index = 0; index < data.length; index++) {
  //         that.setmaintemplatepermission(data[index].ServerRelativeUrl, useremail);
  //         that.setpermissionformaintemplate(data[index].ServerRelativeUrl, useremail);
  //       }
  //     });
  // }

  // setmaintemplatepermission(mainurl, usermail) {
  //   var reacthandler = this;
  //   var url = mainurl.replace(this.props.currentContext.pageContext.web.serverRelativeUrl + '/', '');
  //   const spHttpClient: SPHttpClient = this.props.currentContext.spHttpClient;
  //   var queryUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativeUrl(" + "'" + url + "'" + ")/ListItemAllFields/breakroleinheritance(false)";
  //   const spOpts: ISPHttpClientOptions = {};
  //   spHttpClient.post(queryUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
  //     if (response.ok) {
  //       var permission = reacthandler.readPermission;
  //       sp.web.siteUsers.getByEmail(usermail).get().then(function (userdata) {
  //         var postUrl = reacthandler.props.currentContext.pageContext.web.absoluteUrl + '/_api/web/GetFolderByServerRelativeUrl(' + "'" + url + "'" + ')/ListItemAllFields/roleassignments/addroleassignment(principalid=' + userdata.Id + ',roledefid=' + permission + ')';
  //         spHttpClient.post(postUrl, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
  //           if (response.ok) {
  //           }
  //         });
  //       });
  //     }
  //   })
  // }

  addToList(year, formData) {
    var that = this;
    if (formData.Id > 0) {
      formData.Logs =
        formData.Logs +
        "\n\nUpdated on : " +
        new Date() +
        "\nUpdated by : " +
        this.props.currentContext.pageContext.user.displayName;

      var newUsersAdded = null;
      if (this.newAddedUsers.length > 0) {
        newUsersAdded = "\nNew user(s) added : ";
        for (let u = 0; u < this.newAddedUsers.length; u++) {
          const user = this.newAddedUsers[u];
          newUsersAdded = newUsersAdded + user + "; ";
        }
      }

      var deletededUser = null;
      if (this.deletedUsers.length > 0) {
        deletededUser = "\nUser(s) deleted : ";
        for (let d = 0; d < this.deletedUsers.length; d++) {
          const user = this.deletedUsers[d];
          deletededUser = deletededUser + user + "; ";
        }
      }

      if (newUsersAdded) {
        formData.Logs = formData.Logs + newUsersAdded;
      }
      if (deletededUser) {
        formData.Logs = formData.Logs + deletededUser;
      }

      sp.web
        .getList(listUrl + "ProviderDetails")
        .items.getById(formData.Id)
        .update(formData)
        .then((res) => {
          alertify.success("Provider updated");
          that.setrootfolderpermission(
            "TemplateLibrary/" + formData.TemplateType,
            formData.Title,
            year
          );
          that.loadTableData();
          that.setState({ hideDialog: true });
        });
    } else {
      // var currentMonth = new Date().getMonth() + 1;
      // var stryear = currentYear.toString().substr(2, 2);
      // if (currentMonth < 7) {
      //   stryear = (currentYear - 1).toString().substr(2, 2);
      // }

      var stryear = this.processYear.toString().substr(2, 2);

      formData.ContractId = formData.ContractId + "-" + stryear;
      formData.Logs =
        "Added on : " +
        new Date() +
        "\nAdded by : " +
        this.props.currentContext.pageContext.user.displayName;

      sp.web
        .getList(listUrl + "ProviderDetails")
        .items.add(formData)
        .then((res) => {
          that.createProvider(formData.Title, year, formData);
          that.loadTableData();
        });
    }
  }

  createProvider = (providerName, year, formData) => {
    var reacthandler = this;
    // var currentMonth = new Date().getMonth() + 1;
    // var stryear = year + reacthandler.to + (year + 1);
    // if (currentMonth < 7) {
    //   stryear = year - 1 + reacthandler.to + year;
    // }

    var stryear = reacthandler.processYear +
      reacthandler.to +
      (reacthandler.processYear + 1);

    var folderName = reacthandler.rootFolder + stryear;

    sp.web.folders.getByName(folderName).folders.add(providerName).then((data) => {
      reacthandler.getFolder(
        "TemplateLibrary/" + formData.TemplateType,
        providerName,
        year,
        formData
      );
      setTimeout(() => {
        reacthandler.setrootfolderpermission(
          "TemplateLibrary/" + formData.TemplateType,
          providerName,
          year
        );
        reacthandler.setpermissionsforfolders(
          "TemplateLibrary/" + formData.TemplateType,
          providerName,
          year,
          formData
        );
      }, 10000);
    });

    // sp.web.folders.add(folderName + "/" + providerName).then(function (data) {
    //   reacthandler.getFolder(
    //     "TemplateLibrary/" + formData.TemplateType,
    //     providerName,
    //     year,
    //     formData
    //   );
    //   setTimeout(() => {
    //     reacthandler.setrootfolderpermission(
    //       "TemplateLibrary/" + formData.TemplateType,
    //       providerName,
    //       year
    //     );
    //     reacthandler.setpermissionsforfolders(
    //       "TemplateLibrary/" + formData.TemplateType,
    //       providerName,
    //       year,
    //       formData
    //     );
    //   }, 2000);
    // });


    alertify.success("Provider is created");
    reacthandler.setState({ hideDialog: true });
  };

  setrootfolderpermission = (templateLibrary, providerName, year) => {
    var reacthandler = this;

    // var currentMonth = new Date().getMonth() + 1;
    // var stryear = year + reacthandler.to + (year + 1);
    // if (currentMonth < 7) {
    //   stryear = year - 1 + reacthandler.to + year;
    // }

    var stryear = reacthandler.processYear +
      reacthandler.to +
      (reacthandler.processYear + 1);

    var folderPath =
      reacthandler.rootFolder + stryear + "/" + providerName;

    const spHttpClient: SPHttpClient = this.props.currentContext.spHttpClient;
    var queryUrl =
      this.props.currentContext.pageContext.web.absoluteUrl +
      "/_api/web/GetFolderByServerRelativeUrl(" +
      "'" +
      folderPath +
      "'" +
      ")/ListItemAllFields/breakroleinheritance(false)";
    const spOpts: ISPHttpClientOptions = {};
    spHttpClient
      .post(queryUrl, SPHttpClient.configurations.v1, spOpts)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          var permission = reacthandler.readPermission;
          sp.web
            .getFolderByServerRelativeUrl(templateLibrary)
            .expand(
              "ListItemAllFields/RoleAssignments/Member",
              "ListItemAllFields/RoleAssignments/RoleDefinitionBindings",
              "ListItemAllFields/RoleAssignments/Member/Users"
            )
            .get()
            .then((resdata) => {
              var roleAssignments =
                resdata["ListItemAllFields"].RoleAssignments;
              for (let i = 0; i < roleAssignments.length; i++) {
                const role = roleAssignments[i];
                if (
                  role.Member.LoginName != "BBHC Provider SharePoint Viewers"
                ) {
                  for (let j = 0; j < role.RoleDefinitionBindings.length; j++) {
                    const definition = role.RoleDefinitionBindings[j];
                    var bbhcpostUrl =
                      this.props.currentContext.pageContext.web.absoluteUrl +
                      "/_api/web/GetFolderByServerRelativeUrl(" +
                      "'" +
                      folderPath +
                      "'" +
                      ")/ListItemAllFields/roleassignments/addroleassignment(principalid=" +
                      role.Member.Id +
                      ",roledefid=" +
                      definition.Id +
                      ")";
                    spHttpClient
                      .post(bbhcpostUrl, SPHttpClient.configurations.v1, spOpts)
                      .then((response: SPHttpClientResponse) => { });
                  }
                }
              }
            });
          for (let s = 0; s < reacthandler.userDetails.length; s++) {
            const userData = reacthandler.userDetails[s];
            
            var postUrl =
              this.props.currentContext.pageContext.web.absoluteUrl +
              "/_api/web/GetFolderByServerRelativeUrl(" +
              "'" +
              folderPath +
              "'" +
              ")/ListItemAllFields/roleassignments/addroleassignment(principalid=" +
              userData.Id +
              ",roledefid=" +
              permission +
              ")";
              
            spHttpClient
              .post(postUrl, SPHttpClient.configurations.v1, spOpts)
              .then((response: SPHttpClientResponse) => { var folder= folderPath.split('/');
              console.log('%c'+  'Read Permission is created for ' +userData.Email+ ' for the folder ' + folder[folder.length-1] ,'color:#0000ff')   });
          }
        }
      });
  };

  getFolder = (folderPath, providerName, year, formData) => {
    var reacthandler = this;
    sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.processFolder(0, data, providerName, year, formData);
        }
      });
  };

  processFolder(index, data, providerName, year, formData) {
    var reacthandler = this;
    // var currentMonth = new Date().getMonth() + 1;
    // var stryear = year + reacthandler.to + (year + 1);
    // if (currentMonth < 7) {
    //   stryear = year - 1 + reacthandler.to + year;
    // }

    var stryear = reacthandler.processYear +
      reacthandler.to +
      (reacthandler.processYear + 1);

    var folderName = reacthandler.rootFolder + stryear;
    var clonedUrl = data[index].ServerRelativeUrl.replace(
      "TemplateLibrary/" + formData.TemplateType,
      folderName + "/" + providerName
    );
    var fullurl = clonedUrl;
    
    clonedUrl = clonedUrl.replace(" - Upload", "");
    sp.web.folders.add(clonedUrl).then((res) => {
      var clonedfolder= clonedUrl.split('/') 
      console.log('%c'+clonedfolder[clonedfolder.length-1] + ' is created','color:#00ff00')      
      reacthandler.getFolder(
        data[index].ServerRelativeUrl,
        providerName,
        year,
        formData
      );
      index = index + 1;
      if (index < data.length) {
        reacthandler.processFolder(index, data, providerName, year, formData);
      }
    });
  }

  setpermissionsforfolders = (folderPath, providerName, year, formData) => {
    var reacthandler = this;
    sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.addfolderpermission(
            0,
            data,
            providerName,
            year,
            formData
          );
        }
      });
  };

  addfolderpermission = (index, data, providerName, year, formData) => {
    var reacthandler = this;
    // var currentMonth = new Date().getMonth() + 1;
    // var stryear = year + reacthandler.to + (year + 1);
    // if (currentMonth < 7) {
    //   stryear = year - 1 + reacthandler.to + year;
    // }

    var stryear = reacthandler.processYear +
      reacthandler.to +
      (reacthandler.processYear + 1);

    var serverRelativeUrl = data[index].ServerRelativeUrl;
    var folderName = reacthandler.rootFolder + stryear;
    var clonedUrl = serverRelativeUrl.replace(
      "TemplateLibrary/" + formData.TemplateType,
      folderName + "/" + providerName
    );

    clonedUrl = clonedUrl.replace(" - Upload", "");

    var url = clonedUrl.replace(
      this.props.currentContext.pageContext.web.serverRelativeUrl + "/",
      ""
    );
    const spHttpClient: SPHttpClient = this.props.currentContext.spHttpClient;
    var queryUrl =
      this.props.currentContext.pageContext.web.absoluteUrl +
      "/_api/web/GetFolderByServerRelativeUrl(" +
      "'" +
      url +
      "'" +
      ")/ListItemAllFields/breakroleinheritance(false)";
    const spOpts: ISPHttpClientOptions = {};
    spHttpClient
      .post(queryUrl, SPHttpClient.configurations.v1, spOpts)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          var permission = reacthandler.readPermission;
          var sdata = serverRelativeUrl.split("/");
          var permissionName ="Read";

          var folderFilter = reacthandler.allUploadFolders.filter(
            (c) => c.TemplateType == formData.TemplateType
          );
          var found = false;
          for (let l = 0; l < folderFilter.length; l++) {
            const fold = folderFilter[l].Title.split(" - ");
            if (fold[fold.length - 1] == sdata[sdata.length - 1]) {
              found = true;
              break;
            }
          }

          if (found) {
            permission = reacthandler.contributePermission;
            permissionName = "Upload"
          }

          sp.web
            .getFolderByServerRelativeUrl(serverRelativeUrl)
            .expand(
              "ListItemAllFields/RoleAssignments/Member",
              "ListItemAllFields/RoleAssignments/RoleDefinitionBindings",
              "ListItemAllFields/RoleAssignments/Member/Users"
            )
            .get()
            .then((resdata) => {
              var roleAssignments =
                resdata["ListItemAllFields"].RoleAssignments;
              for (let i = 0; i < roleAssignments.length; i++) {
                const role = roleAssignments[i];
                if (
                  role.Member.LoginName != "BBHC Provider SharePoint Viewers"
                ) {
                  for (let j = 0; j < role.RoleDefinitionBindings.length; j++) {
                    const definition = role.RoleDefinitionBindings[j];
                    var bbhcpostUrl =
                      this.props.currentContext.pageContext.web.absoluteUrl +
                      "/_api/web/GetFolderByServerRelativeUrl(" +
                      "'" +
                      url +
                      "'" +
                      ")/ListItemAllFields/roleassignments/addroleassignment(principalid=" +
                      role.Member.Id +
                      ",roledefid=" +
                      definition.Id +
                      ")";
                    spHttpClient
                      .post(bbhcpostUrl, SPHttpClient.configurations.v1, spOpts)
                      .then((response: SPHttpClientResponse) => { });
                  }
                }
              }
            });
          for (let s = 0; s < reacthandler.userDetails.length; s++) {
            const userData = reacthandler.userDetails[s];
            var postUrl =
              this.props.currentContext.pageContext.web.absoluteUrl +
              "/_api/web/GetFolderByServerRelativeUrl(" +
              "'" +
              url +
              "'" +
              ")/ListItemAllFields/roleassignments/addroleassignment(principalid=" +
              userData.Id +
              ",roledefid=" +
              permission +
              ")";
            spHttpClient
              .post(postUrl, SPHttpClient.configurations.v1, spOpts)
              .then((response: SPHttpClientResponse) => { var folder= url.split('/');
              console.log('%c'+ permissionName+ ' Permission is created for ' +userData.Email+ ' for the folder ' + folder[folder.length-1] ,'color:#0000ff')   });
          }

          reacthandler.setpermissionsforfolders(
            data[index].ServerRelativeUrl,
            providerName,
            year,
            formData
          );
          index = index + 1;
          if (index < data.length) {
            reacthandler.addfolderpermission(
              index,
              data,
              providerName,
              year,
              formData
            );
          }
        }
      });
  };

  createFolder = async (folderPath) => {
    await sp.web.folders.add(folderPath);
  };

  uploadFile = async (event) => {
    var reacthandler = this;
    if (event.target.files && event.target.files.length > 0) {
      reacthandler.fileObj = event.target.files[0];
      reacthandler.setState({ fileName: event.target.files[0].name });
    } else {
      reacthandler.fileObj = null;
    }
  };

  userchange(event) {
    var allusers = this.state.AllUsers;
    allusers[parseInt(event.target.id)] = event.target.value;
    this.setState({ AllUsers: allusers });
  }

  removeuser(index) {
    var allusers = this.state.AllUsers;
    allusers.splice(index, 1);
    this.setState({ AllUsers: allusers });
  }

  newuser() {
    var allusers = this.state.AllUsers;
    allusers.push("");
    this.setState({ AllUsers: allusers });
  }

  templateChange(
    ev: React.FormEvent<HTMLInputElement>,
    option: IChoiceGroupOption
  ): void {
    var formData = this.state.formData;
    formData.TemplateType = option.key;
    this.setState({ formData: formData });
  }

  _onAddRow() {
    var formData = this.state.formData;
    var allUsers = this.state.AllUsers;
    formData = {
      Id: 0,
      Title: "",
      ProviderID: "",
      LegalName: "",
      ContractId: "",
      TemplateType: "Contract Providers",
      Users: "",
      IsDeleted: false,
      Logs: "",
    };
    allUsers = [""];
    this.setState({
      editUsers: "",
      AllUsers: allUsers,
      formData: formData,
      hideDialog: false,
    });
  }

  _onDeleteRow() {
    this.setState({ hideDeleteDialog: false });
  }

  public exportToExcel() {



    // const workbook = new Excel.Workbook();
    // const worksheet = workbook.addWorksheet('My Sheet');
    // var dobCol = worksheet.getRow(1); // You can define a row like 2 , 3
    // worksheet.columns = [
    //   { header: "Provider Name", key: "Title", width: 25 },
    //   { header: "Emails", key: "Users", width: 25 }
    // ];
    var data = [["Provider Name", "Users"]];
    this.state.allItems.forEach(function (item, index) {
      var arrData = item.Users.split(';');
      for (let i = 0; i < arrData.length; i++) {
        if (arrData[i]) {
          if (i == 0) {
            // worksheet.addRow({
            //   Title: item.Title,
            //   Users: arrData[i]
            // });
            data.push([item.Title, arrData[i]])
          } else {
            // worksheet.addRow({
            //   Title: '',
            //   Users: arrData[i]
            // });
            data.push(['', arrData[i]])
          }
        }
      }
    });
    // ['A1', 'B1'].map(key => {
    //   worksheet.getCell(key).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    // });
    // worksheet.eachRow({ includeEmpty: true }, function (cell, index) {
    //   cell._cells.map((key, index) => {
    //     worksheet.getCell(key._address).border = {
    //       top: { style: 'thin' },
    //       left: { style: 'thin' },
    //       bottom: { style: 'thin' },
    //       right: { style: 'thin' }
    //     };
    //   });
    // });

    var worksheetNew = XLSX.utils.aoa_to_sheet(data);
    var new_workbook = XLSX.utils.book_new();


    XLSX.utils.book_append_sheet(new_workbook, worksheetNew, "Provider Details");

    var files = XLSX.write(new_workbook, { type: 'base64' })

    // new_workbook.writeBuffer().then(buffer => FileSaver.saveAs(new Blob([buffer]), 'Provider Details.xlsx'))
    //   .catch(err => console.log('Error writing excel export', err));


    var link = document.createElement('a');
    link.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + encodeURIComponent(files);
    link.setAttribute('download', 'Provider Details.xlsx');

    link.style.display = 'none';
    document.body.appendChild(link);

    link.click();

    document.body.removeChild(link);
  }


  hideDelete() {
    this.setState({ hideDeleteDialog: true });
  }

  deleteItems() {
    var selItems = this._selection.getSelection();
    if (selItems.length > 0) {
      this.updateDeleteTag(0, selItems);
    }
  }

  updateDeleteTag(index, items) {
    this.removeProviderFolderPermission(items);

    var that = this;
    var formData = items[index];
    formData.IsDeleted = true;
    formData.Logs =
      formData.Logs +
      "\n\nDeleted on : " +
      new Date() +
      "\nDeleted by : " +
      that.props.currentContext.pageContext.user.displayName;
    sp.web
      .getList(listUrl + "ProviderDetails")
      .items.getById(formData.Id)
      .update(formData)
      .then((res) => {
        var users = formData.Users.split(";");
        for (let j = 0; j < users.length; j++) {
          const user = users[j];
          if (user) {
            that.setpermissionfornewuser(
              "TemplateLibrary/" + formData.TemplateType,
              user,
              false
            );
          }
        }

        index = index + 1;
        if (index < items.length) {
          that.updateDeleteTag(index, items);
        } else {
          alertify.success("Provider deleted successfully");
          that.loadTableData();
          that.setState({ hideDeleteDialog: true });
        }
      });
  }

  removeProviderFolderPermission(items) {
    for (let index = 0; index < items.length; index++) {
      const provider = items[index];
      this.removerMainFolderUserPermission(provider);
    }
  }

  removerMainFolderUserPermission(provider) {
    var reacthandler = this;
    var users = provider.Users.split(";");
    for (let j = 0; j < users.length; j++) {
      const user = users[j];
      if (user) {
        sp.web.siteUsers
          .getByEmail(user)
          .get()
          .then(function (userdata) {
            // var contract = provider.ContractId.substr(
            //   provider.ContractId.length - 2,
            //   2
            // );
            // var nextyear = parseInt(contract) + 1;
            // var currentyearprefix = currentYear.toString().substr(0, 2);
            // var yearfolder =
            //   (currentyearprefix + contract) +
            //   reacthandler.to +
            //   (currentyearprefix + nextyear);

            var yearfolder = reacthandler.processYear +
              reacthandler.to +
              (reacthandler.processYear + 1);

            var providerFolder =
              reacthandler.rootFolder + "/" + yearfolder + "/" + provider.Title;

            var postUrl =
              reacthandler.props.currentContext.pageContext.web.absoluteUrl +
              "/_api/web/GetFolderByServerRelativeUrl(" +
              "'" +
              providerFolder +
              "'" +
              ")/ListItemAllFields/roleassignments/removeroleassignment(principalid=" +
              userdata.Id +
              ",roledefid=" +
              reacthandler.readPermission +
              ")";

            const spHttpClient: SPHttpClient =
              reacthandler.props.currentContext.spHttpClient;
            const spOpts: ISPHttpClientOptions = {};

            spHttpClient
              .post(postUrl, SPHttpClient.configurations.v1, spOpts)
              .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                }
              });
          });
      }
    }
  }

  _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const newColumns: IColumn[] = this._columns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    var items = this.state.items;
    const newItems = this._copyAndSort(
      items,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    this.setState({
      items: newItems,
    });
  };

  _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    const key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
  }

  public render(): React.ReactElement<IProvidersProp> {
    const _renderItemColumn = (item, index: number, column: IColumn) => {
      const fieldContent = item[column.fieldName] as string;
      switch (column.fieldName) {
        case "Title":
          return (
            <Link id={item["Id"] + ""} onClick={this.editItem.bind(this)}>
              {fieldContent}
            </Link>
          );

        default:
          return <span>{fieldContent}</span>;
      }
    };

    const modelProps = {
      isBlocking: true,
      topOffsetFixed: false,
    };

    const stackTokens: IStackTokens = {
      childrenGap: 4,
    };
    const stackStyles: Partial<IStackStyles> = {
      root: {
        // width: 600,
      },
    };
    const gridStyles: Partial<IDetailsListStyles> = {
      headerWrapper: [
        {
          selectors: {
            ".ms-DetailsHeader-cell": {
              backgroundColor: "rgb(243, 242, 241)",
              fontFamily: "Poppins, sans-serif",
              marginTop: 3,
            },
            ".ms-DetailsHeader-cellName": {
              fontWeight: 500,
              fontSize: "14px",
            },
          },
        },
      ],
      contentWrapper: [
        {
          selectors: {
            ".ms-DetailsRow": {
              fontFamily: "Poppins, sans-serif",
            },
          },
        },
      ],
    };
    const columnstyle: Partial<IStackProps> = {
      tokens: {
        childrenGap: 5,
      },
      styles: {
        root: {
          width: "100%",
          // paddingTop: 10,
        },
      },
    };

    const iconcolumnstyle: Partial<IStackProps> = {
      tokens: {
        childrenGap: 5,
      },
      styles: {
        root: {
          width: 30,
          paddingTop: 27,
        },
      },
    };

    const btnstackTokens: IStackTokens = {
      childrenGap: 4,
    };
    const btnstackStyles: Partial<IStackStyles> = {
      root: {
        width: 600,
      },
    };

    const btncolumnstyle: Partial<IStackProps> = {
      tokens: {
        childrenGap: 4,
      },
      styles: {
        root: {
          width: 100,
          paddingTop: 10,
        },
      },
    };

    const classes = mergeStyleSets({
      cell: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        margin: "80px",
        float: "left",
        height: "50px",
        width: "50px",
        fontSize: "200px",
      },
      icon: {
        fontSize: "50px",
      },
      code: {
        background: "#f2f2f2",
        borderRadius: "4px",
        padding: "4px",
      },
      navigationText: {
        width: 100,
        margin: "0 5px",
      },
    });

    const commandBarStyles: Partial<ICommandBarStyles> = {
      root: {
        marginBottom: 15,
        padding: 0,
        selectors: {
          ".ms-Button": {
            borderWidth: 0,
            marginRight: 5,
            marginLeft: 5,
            fontFamily: "Poppins, sans-serif",
          },
        },
      },
    };
    const choisestyle: Partial<IChoiceGroupStyles> = {
      label: {
        fontFamily: "Poppins, sans-serif",
      },
      flexContainer: [
        {
          marginTop: -5,
          marginBottom: 5,
          selectors: {
            ".ms-ChoiceField": {
              fontFamily: "Poppins, sans-serif",
              marginTop: 3,
              fontSize: "13px",
            },
          },
        },
      ],
    };
    const dialogStyles: Partial<IDialogStyles> = {
      main: [
        {
          fontFamily: "Poppins, sans-serif",
          selectors: {
            ".ms-Dialog-title": {
              fontFamily: "Poppins, sans-serif",
            },
            ".ms-Dialog-subText": {
              fontFamily: "Poppins, sans-serif",
            },
          },
        },
      ],
    };
    // const labelId: string = useId('dialogLabel');
    // const subTextId: string = useId('subTextLabel');
    // const dialogStyles = { main: { maxWidth: 450 } };
    // const [isDraggable, { toggle: toggleIsDraggable }] = useBoolean(false);

    // const modalProps = React.useMemo(
    //   () => ({
    //     titleAriaId: labelId,
    //     subtitleAriaId: subTextId,
    //     isBlocking: false,
    //     styles: dialogStyles,
    //     dragOptions: undefined,
    //   }),
    //   [isDraggable],
    // );

    const dialogContentProps = {
      type: DialogType.normal,
      title: "Delete",
      closeButtonAriaLabel: "Close",
      subText: "Do you want to delete the selected Provider(s)?",
    };

    return (
      <div>
        <CommandBar
          styles={commandBarStyles}
          items={[
            {
              key: "addRow",
              text: "Add provider",
              iconProps: { iconName: "Add" },
              onClick: this._onAddRow.bind(this),
            },
            {
              key: "deleteRow",
              text: "Delete row",
              iconProps: { iconName: "Delete" },
              onClick: this._onDeleteRow.bind(this),
            },
            {
              key: "exportToExcel",
              text: "Export to excel",
              iconProps: { iconName: "Download" },
              onClick: this.exportToExcel.bind(this)
            }
          ]}
        />

        <Fabric>
          <div className={styles.announcement}>
            <div className={exampleChildClass}>
              {this.state.selectionDetails}
            </div>
            <Announced message={this.state.selectionDetails} />

            <TextField
              prefix="Filter by provider name:"
              onChange={this._onFilter.bind(this)}
              styles={textFieldStyles}
              className={styles.searchTextbox}
            />
            <Announced
              message={`Number of items after filter applied: ${this.state.items.length}.`}
            />
          </div>
          <div className={styles.tableContainer}>
            <MarqueeSelection selection={this._selection}>
              <DetailsList
                items={this.state.items}
                columns={this._columns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selection={this._selection}
                selectionPreservedOnEmptyClick={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                checkButtonAriaLabel="Row checkbox"
                onItemInvoked={this._onItemInvoked}
                onRenderItemColumn={_renderItemColumn}
                styles={gridStyles}
              />
            </MarqueeSelection>
          </div>
        </Fabric>

        <Dialog
          hidden={this.state.hideDialog}
          modalProps={modelProps}
          minWidth="400px"
          styles={dialogStyles}
        >
          <Stack {...columnstyle}>
            {this.state.formData.Id == 0 ? (
              <div>
                <ChoiceGroup
                  defaultSelectedKey={this.state.formData.TemplateType}
                  options={this.templateTypes}
                  onChange={this.templateChange.bind(this)}
                  label="Provider Type"
                  styles={choisestyle}
                  className={styles.heading_01}
                />

                <TextField
                  label="Provider ID"
                  onChange={(e) => this.inputChangeHandler.call(this, e)}
                  name="ProviderID"
                  value={this.state.formData.ProviderID}
                  required
                  className={styles.input_field}
                ></TextField>

                <TextField
                  label="Provider Name"
                  onChange={this.providerNameChange}
                  name="Title"
                  value={this.state.formData.Title}
                  required
                  className={styles.input_field}
                ></TextField>

                <TextField
                  label="Contract ID"
                  onChange={(e) => this.inputChangeHandler.call(this, e)}
                  value={this.state.formData.ContractId}
                  name="ContractId"
                  required
                  className={styles.input_field}
                ></TextField>

                <TextField
                  label="Legal Name"
                  onChange={(e) => this.inputChangeHandler.call(this, e)}
                  value={this.state.formData.LegalName}
                  name="LegalName"
                  required
                  className={styles.input_field}
                ></TextField>
              </div>
            ) : (
              ""
            )}

            {this.state.AllUsers.map((user, index) => {
              if (this.state.AllUsers.length == 1) {
                return (
                  <div>
                    <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                      <Stack {...columnstyle}>
                        <TextField
                          label="User"
                          id={index + ""}
                          onChange={(e) => this.userchange.call(this, e)}
                          value={user}
                          name="userName"
                          required
                          className={styles.input_field}
                        ></TextField>
                      </Stack>

                      <Stack {...iconcolumnstyle}>
                        <IconButton
                          iconProps={{ iconName: "Add" }}
                          title="Add User"
                          ariaLabel="Add"
                          onClick={this.newuser.bind(this)}
                          className={styles.primary_button}
                        />
                      </Stack>
                    </Stack>
                  </div>
                );
              } else {
                return (
                  <div>
                    <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                      <Stack {...columnstyle}>
                        <TextField
                          label="User"
                          id={index + ""}
                          onChange={(e) => this.userchange.call(this, e)}
                          value={user}
                          name="userName"
                          required
                          className={styles.input_field}
                        ></TextField>
                      </Stack>

                      <Stack {...iconcolumnstyle}>
                        <IconButton
                          iconProps={{ iconName: "Add" }}
                          title="Add User"
                          ariaLabel="Add"
                          onClick={this.newuser.bind(this)}
                          className={styles.primary_button}
                        />
                      </Stack>

                      <Stack {...iconcolumnstyle}>
                        <IconButton
                          iconProps={{ iconName: "Cancel" }}
                          title="Remove User"
                          ariaLabel="Cancel"
                          onClick={this.removeuser.bind(this, index)}
                          className={styles.secondary_button}
                        />
                      </Stack>
                    </Stack>
                  </div>
                );
              }
            })}
          </Stack>

          <DialogFooter>
            <PrimaryButton
              onClick={this.processInputProvider}
              className={styles.button_primary}
            >
              {this.state.formData.Id == 0 ? "Add New Provider" : "Submit"}
            </PrimaryButton>
            <DefaultButton onClick={this.hideDialog.bind(this)} text="Close" />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={this.state.hideDeleteDialog}
          dialogContentProps={dialogContentProps}
          styles={dialogStyles}
        >
          <DialogFooter>
            <PrimaryButton
              onClick={this.deleteItems.bind(this)}
              text="Yes"
              className={styles.button_primary}
            />
            <DefaultButton onClick={this.hideDelete.bind(this)} text="No" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
