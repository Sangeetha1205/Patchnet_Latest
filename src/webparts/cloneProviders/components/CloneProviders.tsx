import * as React from "react";
import { escape } from "@microsoft/sp-lodash-subset";
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
import { IIconProps, Modal } from "office-ui-fabric-react";
import {
  mergeStyles,
  mergeStyleSets,
} from "office-ui-fabric-react/lib/Styling";
import {
  ITextFieldStyles,
  TextField,
} from "office-ui-fabric-react/lib/TextField";
import { getId } from "office-ui-fabric-react/lib/Utilities";
import "../../../ExternalRef/CSS/alertify.min.css";
import "../../../ExternalRef/CSS/style.css";
import * as XLSX from "xlsx";
import * as FileSaver from "file-saver";

var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

import styles from "./CloneProviders.module.scss";
import { ICloneProvidersProps } from "./ICloneProvidersProps";

const cancelIcon: IIconProps = { iconName: "Cancel" };
export interface ICloneState {
  providerDetails: any[];
  ModalSwitch: any;
  isOpenModal: boolean;
}

var listUrl = "";
var currentYear = new Date().getFullYear();

// var currentMonth = new Date().getMonth() + 1;
// var stryear = currentYear + " to " + (currentYear + 1);
// if (currentMonth < 7) {
//   stryear = currentYear - 1 + " to " + currentYear;
// }

// var newProvidersLibrary = 'Providers Library ' + currentYear + '-' + (currentYear + 1);
//Sangeetha
/*var stryear = currentYear + " to " + (currentYear + 1);
var newProvidersLibrary = "Providers Library " + stryear;*/

//Sangeetha
var stryear = currentYear + "-" + (currentYear + 1);
var newProvidersLibrary = "FY " + stryear;
var strnewyear = currentYear + "" + (currentYear + 1);
var newProviderLibrary = "FY " + strnewyear;

export default class CloneProviders extends React.Component<
  ICloneProvidersProps,
  ICloneState
> {
  configurationList = "Configuration";
  currentUser = null;
  contributePermission = null;
  readPermission = null;
  allUploadFolders = [];

  constructor(props: ICloneProvidersProps) {
    super(props);
    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });

    listUrl = this.props.currentContext.pageContext.web.absoluteUrl;
    var siteindex = listUrl.toLocaleLowerCase().indexOf("sites");
    listUrl = listUrl.substr(siteindex - 1) + "/Lists/";

    var that = this;

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

    alertify.set("notifier", "position", "top-right");

    this.currentUser = sp.web.currentUser();

    this.state = {
      providerDetails: [],
      ModalSwitch: [],
      isOpenModal: false,
    };
  }

  cloneProviders = () => {
    sp.web
      .getList(listUrl + this.configurationList)
      .items.select("Title")
      .filter("Title eq '" + currentYear + "'")
      .get()
      .then((res) => {
        if (res.length == 0) {
          sp.web
            .getList(listUrl + this.configurationList)
            .items.add({
              Title: currentYear + "",
            })
            .then((res) => {
              this.createProviderLibrary();
              this.setState({ isOpenModal: false });
            });
        } else {
          alertify.error("Clone already proccessed");
          this.setState({ isOpenModal: false });
        }
      });
  };

  createProviderLibrary = () => {
    sp.web.lists
      .add(newProvidersLibrary, newProvidersLibrary, 101, true, {
        OnQuickLaunch: true,
      })
      .then((res) => {
        this.loadProviderDetails();
      });
  };

  loadProviderDetails = () => {
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
      .then((data) => {
        this.setState({ providerDetails: data });
        this.updateContractId();
      });
  };

  updateContractId = () => {
    var reacthandler = this;
    for (let index = 0; index < this.state.providerDetails.length; index++) {
      const provider = this.state.providerDetails[index];
      provider.ContractId = provider.ContractId.substring(
        0,
        provider.ContractId.length - 2
      );
      provider.ContractId =
        provider.ContractId + currentYear.toString().substring(2);
      sp.web
        .getList(listUrl + "ProviderDetails")
        .items.getById(provider.Id)
        .update(provider)
        .then((res) => {});

      // var urls = reacthandler.props.siteUrl.split('/');
      // var url = '';
      // for (let j = 3; j < urls.length; j++) {
      //   url = url + urls[j] + '/';
      // }
      // url = url + newProvidersLibrary + '/' + provider.Title;
      // sp.web.folders.add(url).then((res) => {
      //   this.getFolder("TemplateLibrary/" + provider.TemplateType, currentYear, provider);
      // });

      sp.web.folders
        .getByName(newProviderLibrary)
        .folders.add(provider.Title)
        .then((data) => {
          this.getFolder(
            "TemplateLibrary/" + provider.TemplateType,
            currentYear,
            provider
          );
        });

      // this.getFolder("TemplateLibrary/" + provider.TemplateType, currentYear, provider);

      setTimeout(() => {
        reacthandler.setrootfolderpermission(
          "TemplateLibrary/" + provider.TemplateType,
          provider
        );
        reacthandler.setpermissionsforfolders(
          "TemplateLibrary/" + provider.TemplateType,
          provider
        );
      }, 2000);
    }

    alertify.success("Cloned successfully");
  };

  setrootfolderpermission = (templateLibrary, provider) => {
    var reacthandler = this;
    var folderPath = newProviderLibrary + "/" + provider.Title;

    const spHttpClient: SPHttpClient = this.props.currentContext.spHttpClient;
    var queryUrl =
      reacthandler.props.currentContext.pageContext.web.absoluteUrl +
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
                      reacthandler.props.currentContext.pageContext.web
                        .absoluteUrl +
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
                      .then((response: SPHttpClientResponse) => {});
                  }
                }
              }
            });

          var userDetails = provider.Users.split(";");
          for (let s = 0; s < userDetails.length; s++) {
            var user = userDetails[s];
            if (user) {
              sp.web.siteUsers
                .getByEmail(user)
                .get()
                .then(function (data) {
                  var postUrl =
                    reacthandler.props.currentContext.pageContext.web
                      .absoluteUrl +
                    "/_api/web/GetFolderByServerRelativeUrl(" +
                    "'" +
                    folderPath +
                    "'" +
                    ")/ListItemAllFields/roleassignments/addroleassignment(principalid=" +
                    data.Id +
                    ",roledefid=" +
                    permission +
                    ")";
                  spHttpClient
                    .post(postUrl, SPHttpClient.configurations.v1, spOpts)
                    .then((response: SPHttpClientResponse) => {});
                });
            }
          }
        }
      });
  };

  setpermissionsforfolders = (folderPath, provider) => {
    var reacthandler = this;
    sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.addfolderpermission(0, data, provider);
        }
      });
  };

  addfolderpermission = (index, data, provider) => {
    var reacthandler = this;
    var serverRelativeUrl = data[index].ServerRelativeUrl;
    var clonedUrl = serverRelativeUrl.replace(
      "TemplateLibrary/" + provider.TemplateType,
      newProviderLibrary
    );

    clonedUrl = clonedUrl.replace(" - Upload", "");

    var url = clonedUrl.replace(
      reacthandler.props.currentContext.pageContext.web.serverRelativeUrl + "/",
      ""
    );
    const spHttpClient: SPHttpClient =
      reacthandler.props.currentContext.spHttpClient;
    var queryUrl =
      reacthandler.props.currentContext.pageContext.web.absoluteUrl +
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

          var folderFilter = reacthandler.allUploadFolders.filter(
            (c) => c.TemplateType == provider.TemplateType
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
                      reacthandler.props.currentContext.pageContext.web
                        .absoluteUrl +
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
                      .then((response: SPHttpClientResponse) => {});
                  }
                }
              }
            });

          var userDetails = provider.Users.split(";");
          for (let s = 0; s < userDetails.length; s++) {
            const user = userDetails[s];
            if (user) {
              sp.web.siteUsers
                .getByEmail(user)
                .get()
                .then(function (data) {
                  var postUrl =
                    reacthandler.props.currentContext.pageContext.web
                      .absoluteUrl +
                    "/_api/web/GetFolderByServerRelativeUrl(" +
                    "'" +
                    url +
                    "'" +
                    ")/ListItemAllFields/roleassignments/addroleassignment(principalid=" +
                    data.Id +
                    ",roledefid=" +
                    permission +
                    ")";
                  spHttpClient
                    .post(postUrl, SPHttpClient.configurations.v1, spOpts)
                    .then((response: SPHttpClientResponse) => {});
                });
            }
          }

          reacthandler.setpermissionsforfolders(
            data[index].ServerRelativeUrl,
            provider
          );
          index = index + 1;
          if (index < data.length) {
            reacthandler.addfolderpermission(index, data, provider);
          }
        }
      });
  };

  getFolder = (folderPath, year, provider) => {
    var reacthandler = this;
    sp.web
      .getFolderByServerRelativePath(folderPath)
      .folders.get()
      .then(function (data) {
        if (data.length > 0) {
          reacthandler.processFolder(0, data, year, provider);
        }
      });
  };

  processFolder = (index, data, year, provider) => {
    var reacthandler = this;
    var currentMonth = new Date().getMonth() + 1;
    var stryear = year + "-" + (year + 1);
    if (currentMonth < 7) {
      stryear = year - 1 + "-" + year;
    }
    var folderName = newProviderLibrary;
    var clonedUrl = data[index].ServerRelativeUrl.replace(
      "TemplateLibrary/" + provider.TemplateType,
      folderName + "/" + provider.Title
    );
    clonedUrl = clonedUrl.replace(" - Upload", "");
    sp.web.folders.add(clonedUrl).then((res) => {
      reacthandler.getFolder(data[index].ServerRelativeUrl, year, provider);
      index = index + 1;
      if (index < data.length) {
        reacthandler.processFolder(index, data, year, provider);
      }
    });
  };

  public render(): React.ReactElement<ICloneProvidersProps> {
    return (
      <div>
        
        Using this tool will allow the cloning of a new Provider Library based on the most up-to-date <b>Template Library</b> and the existing list of providers in the <b>Provider Details list</b>. 
        <p style={{marginTop: 20}}>        
        <b>Some considerations:</b><br></br>
        <div style={{marginLeft: 30}}>
        <b>1.</b> Access to folders will be cloned based on the assigned permissions to the existing provider library in use. <br></br>
        <b>2.</b> Folder structure will be cloned based on the most updated Template Library. If any folder has been manually added to a provider, it will not be cloned until it is updated in the template library.<br></br> 
        <b>3.</b> Changes in the Template Library only affect new providers folder created. Any existing provider folder will have to be manually updated. <br></br>
        <b>4.</b> Cloning is irreversible. Please make sure you have reviewed all the required changes before cloning the library.<br></br> 
        <b>5.</b> Allow up to 5 days after the cloning to make sure all features are working properly before requesting any submission to the new created library. <br></br>
        </div>
        </p>
        <div className="clone-btn-sec">    
          <PrimaryButton
            className={styles.button_primary}
            onClick={() => {
              this.setState({ isOpenModal: true });
            }}
            text="Clone Library"
          />
        </div>

        <Modal
          className="Patch-Modal"
          isOpen={this.state.isOpenModal}
          // onDismiss={hideModal}
          isBlocking={false}
        >
          <div className="header-btn"></div>

          <div className="PatchModalBody">Are you sure you want to proceed?</div>
          <div className="Modal-footer-btn-section">
            <DefaultButton
              text="Cancel"
              className={styles.button_default}
              onClick={() => this.setState({ isOpenModal: false })}
            />
            <PrimaryButton
              text="Proceed"
              className={styles.button_primary}
              onClick={this.cloneProviders}
            />
          </div>
        </Modal>
      </div>
    );
  }
}
