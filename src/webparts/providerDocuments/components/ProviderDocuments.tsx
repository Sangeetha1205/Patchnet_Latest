import * as React from "react";
import { escape } from "@microsoft/sp-lodash-subset";
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import "@pnp/sp/files";


import "@pnp/sp/site-groups";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";


import "@pnp/sp/site-groups";

import { IItemAddResult } from "@pnp/sp/items";

import { PrimaryButton } from "@fluentui/react";
import { Label } from "office-ui-fabric-react/lib/Label";
import { Link } from "office-ui-fabric-react/lib/Link";

import { getId } from "office-ui-fabric-react/lib/Utilities";

import "alertifyjs";
import styles from "./ProviderDocuments.module.scss";

import "../../../ExternalRef/CSS/style.css";
import "../../../ExternalRef/CSS/alertify.min.css";
var alertify: any = require("../../../ExternalRef/JS/alertify.min.js");

import { Image, IImageProps } from "office-ui-fabric-react/lib/Image";
import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";

import {
  TextField,
  MaskedTextField,
} from "office-ui-fabric-react/lib/TextField";

import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
} from "office-ui-fabric-react/lib/Dropdown";

const currentYear = new Date().getFullYear();

const fileId = getId("anInput");
import { IProviderDocumentsProps } from "./IProviderDocumentsProps";

export interface IBbhcState {
  folders: any[];
  destinationPath: any[];
  file: any;
  selectedPath: string;
  selectedProvider: string;
  notes: string;
  fileName: "";
  previousyeardata: any[];
  allProviders: any[];
  allData: any[];
  groupUsers: any[]
}
const dropDownStyles: Partial<IDropdownStyles> = {
  root: {
    selectors: {
      ".ms-Label": {
        fontFamily: "Poppins, sans-serif",
      },
    },
  },
};

const dropDown2Styles: Partial<IDropdownStyles> = {
  root: {
    selectors: {
      ".ms-Label": {
        fontFamily: "Poppins, sans-serif",
      },
    },
    marginLeft: "0",
  },
};

var listUrl = '';



export default class ProviderDocuments extends React.Component<
  IProviderDocumentsProps,
  IBbhcState
> {
  currentYear = new Date().getFullYear();

  rootFolder = "FY ";
  to = '';
  configurationList = "Configuration";
  processYear = currentYear;


  templateLibrary = "TemplateLibrary";
  generalSubmission = "general submission";
  generalSubmissionChanged = false;
  staffGroupName = "Staff";
  constructor(props) {
    super(props);
    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });

    alertify.set("notifier", "position", "top-right");

    listUrl = this.props.currentContext.pageContext.web.absoluteUrl;
    var siteindex = listUrl.toLocaleLowerCase().indexOf('sites');
    listUrl = listUrl.substr(siteindex - 1) + '/Lists/';
    this.state = {
      folders: [],
      destinationPath: [],
      file: null,
      selectedPath: "",
      selectedProvider: "",
      fileName: "",
      notes: "",
      previousyeardata: [],
      allProviders: [],
      allData: [],
      groupUsers: []
    };
    this.getProviderMetaData();
    this.getGroupMembers();

    sp.web
      .getList(listUrl + this.configurationList)
      .items.select("Title")
      .filter("Title eq '" + currentYear + "'")
      .get()
      .then((res) => {
        if (res.length == 0) {
          this.processYear = currentYear - 1;
        } else {
          this.processYear = currentYear;
        }
      });


  }

  async getGroupMembers() {
     const groupName = "ProvSPInvoiceSubmissionNotificationTest";//for dev
    // const groupName = "ProvSP Invoice Submission Notification";//for Live

    var grp = await sp.web.siteGroups.getByName(groupName).users.get().then((members) => {
      var emailArray = [];
      members.forEach((item) => {
        emailArray.push(item.Email);
      })
      this.setState({ groupUsers: emailArray });
      console.log(this.state.groupUsers);
    });
  }

  getProviderMetaData() {
    var that = this;
    sp.web.getList(listUrl + "ProviderDetails")
      .items.top(5000).select("Title", "ContractId", "TemplateType", "Users","IsDeleted").filter("IsDeleted eq false")
      // .filter(
      //   "substringof('" +
      //   this.props.currentContext.pageContext.user.email.toLowerCase() +
      //   "',Users) and IsDeleted eq 0"
      // )
      .get()

      .then((res) => {
        if (res.length > 0) {
          var currentEmail = this.props.currentContext.pageContext.user.email.toLowerCase();
          // var currentMonth = new Date().getMonth() + 1;
          var stryear = that.processYear;
          // if (currentMonth < 7) {
          //   stryear = that.currentYear - 1;
          // }
          var previousyeardata = that.state.previousyeardata;
          var allProviders = that.state.allProviders;
          var allData = that.state.allData;
          var dataLoaded = false;
          for (let j = 0; j < res.length; j++) {
            if (res[j].Users.toLowerCase().indexOf(currentEmail) >= 0) {
              const providerData = res[j];
              allData.push(providerData);
              var contract = providerData.ContractId.substr(
                providerData.ContractId.length - 2,
                2
              );
              if (contract != stryear.toString().substr(2, 2)) {
                var nextyear = parseInt(contract) + 1;
                var currentyearprefix = that.processYear.toString().substr(0, 2);
                previousyeardata.push({
                  Title:
                    "FY " +
                    (currentyearprefix + contract) +
                    "-" +
                    (currentyearprefix + nextyear),
                  URL:
                    that.props.siteUrl +
                    "/" +
                    this.rootFolder +
                    (currentyearprefix + contract) +
                    this.to +
                    (currentyearprefix + nextyear) +
                    "/" +
                    providerData.Title,
                });
              } else {
                if (!dataLoaded) {
                  dataLoaded = true;
                }
                allProviders.push({
                  key: providerData.Title,
                  text: providerData.Title,
                });
              }
            }
          }
          that.setState({
            previousyeardata: previousyeardata,
            allProviders: allProviders,
            allData: allData,
          });
        } else {
          that.setState({ folders: [] });
        }
      });
  }

  loadUploadFolders(templateType) {
    var that = this;
    sp.web.getList(listUrl + "DocumentType")
      .items.select("Title", "TemplateType")
      .filter("TemplateType eq '" + templateType + "'")
      .get()
      .then((res) => {
        var allFolders = that.state.folders;
        allFolders = [];
        var generalSub = null;
        for (let index = 0; index < res.length; index++) {
          var cleartext = res[index].Title.replace(" - Upload", "");
          var url = cleartext.replace(" - ", "/");
          if (cleartext != "General Submissions") {
            allFolders.push({
              key: url,
              text: cleartext,
            });
          } else {
            generalSub = {
              key: url,
              text: cleartext,
            };
          }
        }
        if (generalSub) {
          allFolders.push(generalSub);
        }
        that.setState({ folders: allFolders });
      });
  }

  // getFolders(folderName, templateType, displayName) {
  //   var url = this.templateLibrary + "/" + templateType;
  //   if (folderName) {
  //     url = url + "/" + folderName;
  //   }
  //   var that = this;
  //   var allFolders = that.state.folders;
  //   sp.web
  //     .getFolderByServerRelativePath(url)
  //     .folders.get()
  //     .then(function (data) {
  //       if (data.length > 0) {
  //         for (let index = 0; index < data.length; index++) {
  //           var text = '';
  //           var cleartext = data[index].Name.replace(' - Upload', '')
  //           if (displayName) {
  //             text = displayName + ' - ' + cleartext;
  //           } else {
  //             text = cleartext;
  //           }
  //           if (data[index].Name.toLocaleLowerCase().indexOf('upload') > 0) {
  //             allFolders.push({
  //               key: folderName + '/' + cleartext,
  //               text: text,
  //             });
  //           }
  //           that.setState({ folders: allFolders });
  //           that.getFolders(folderName + '/' + data[index].Name, templateType, text);
  //         }
  //       }
  //     });
  // }

  fileUpload(e) {
    var files = e.target.files;
    if (files && files.length > 0) {
      this.setState({ file: files[0], fileName: files[0].name });
    } else {
      this.setState({ file: null });
    }
  }

  // uploadFile() {
  //   if (this.state.file) {
  //     var destinationPaths = this.state.destinationPath;
  //     if (destinationPaths.length > 0) {

  //       if (destinationPaths.length != this.state.folders.length) {
  //         alertify.error('Fill all dropdown values');
  //         return;
  //       }

  //       var folderPath = this.sharedDocument + '/' + this.currentYear + '/' + this.userName + '/';
  //       for (let index = 0; index < destinationPaths.length; index++) {
  //         folderPath = folderPath + destinationPaths[index].value + '/';
  //       }
  //       var that = this;
  //       sp.web.getFolderByServerRelativeUrl(folderPath).files.add(that.state.file.name, that.state.file, true)
  //         .then(function (result) {
  //           alertify.success('File uploaded successfully');
  //         });
  //     } else {
  //       alertify.error('Select any folder');
  //     }
  //   } else {
  //     alertify.error('Select any file');
  //   }
  // }

  uploadFile() {
    if (this.state.file) {
      if (!this.state.selectedProvider) {
        alertify.error("Select any provider");
        return;
      }
      var selectedPath = this.state.selectedPath;
      if (selectedPath) {
        if (
          selectedPath.toLocaleLowerCase().indexOf(this.generalSubmission) >= 0
        ) {
          if (!this.state.notes) {
            alertify.error("Notes is required");
            return;
          }
        }
        // var currentMonth = new Date().getMonth() + 1;
        var stryear = this.processYear + this.to + (this.processYear + 1);
        // if (currentMonth < 7) {
        //   stryear = this.currentYear - 1 + this.to + this.currentYear;
        // }
        var folderName = stryear + "/" + this.state.selectedProvider;

        var folderPath =
          this.rootFolder + folderName + "/" + selectedPath;
        var that = this;
        sp.web
          .getFolderByServerRelativeUrl(folderPath)
          .files.add(that.state.file.name, that.state.file, true)
          .then(function (result) {

            var to = that.state.allData.filter(c => c.Title == that.state.selectedProvider)[0].Users;
            to = to.substr(0, to.length - 1);
            to = to.split(';');
            that.notifyUsers(to, that, result);

            if (
              selectedPath
                .toLocaleLowerCase()
                .indexOf(that.generalSubmission) >= 0
            ) {
              result.file.listItemAllFields.get().then(function (fileData) {
                sp.web.getList(listUrl + that.rootFolder)
                  .items.getById(fileData.Id)
                  .update({ FileNotes: that.state.notes })
                  .then(function () {
                    sp.web.lists
                      .getByTitle("EmailConfig")
                      .items.get()
                      .then((res) => {
                        var filepath =
                          that.props.currentContext.pageContext.web
                            .absoluteUrl +
                          "/" +
                          folderPath +
                          "/" +
                          that.state.file.name;
                        var to = res[0].To.split(";");
                        var cc = [];
                        if (res[0].CC) {
                          cc = res[0].CC.split(";");
                        }
                        var bcc = [];
                        if (res[0].BCC) {
                          bcc = res[0].BCC.split(";");
                        }
                        const emailProps: IEmailProperties = {
                          To: to,
                          CC: cc,
                          BCC: bcc,
                          Subject: res[0].Subject,
                          Body:
                            "New file is uploaded in the general submission folder for the <a href='" +
                            filepath +
                            "'>" +
                            that.state.selectedProvider +
                            "</a> provider.\n\nNotes : " +
                            that.state.notes,
                          AdditionalHeaders: {
                            "content-type": "text/html",
                          },
                        };
                        sp.utility.sendEmail(emailProps);
                        alertify.success("File uploaded successfully");
                        var invoice = that.state.selectedPath.split('/')[that.state.selectedPath.split('/').length - 1];
                        if (invoice.toLocaleLowerCase() == 'invoice') {
                          that.getStaffs(that, result);
                        } else {
                          location.reload();
                        }
                      });
                  });
              });
            } else {
              alertify.success("File uploaded successfully");
              var invoice = that.state.selectedPath.split('/')[that.state.selectedPath.split('/').length - 1];
              if (invoice.toLocaleLowerCase() == 'invoice') {
                that.getStaffs(that, result);
              } else {
                location.reload();
              }
            }
          });
      } else {
        alertify.error("Select any folder");
      }
    } else {
      alertify.error("Select any file");
    }
  }

  inputChangeHandler(e) {
    this.setState({
      notes: e.target.value,
    });
  }
  getStaffs = (that, fileResult) => {
    // sp.web.siteGroups	
    //   .getByName(that.staffGroupName)	
    //   .users.get()	
    //   .then((result) => {	
    //     var to = [];	
    //     for (let index = 0; index < result.length; index++) {	
    //       const element = result[index];	
    //       to.push(element.Email);	
    //     }	
    //     if (to.length > 0) {	
    //       this.notifyStaffs(to, that, fileResult);	
    //     }	
    //     location.reload();	
    //   });	
    this.notifyStaffs(this.state.groupUsers, that, fileResult);
  }
  notifyStaffs = (to, that, result) => {
    var filePath = that.props.currentContext.pageContext.web.absoluteUrl.replace(that.props.currentContext.pageContext.web.serverRelativeUrl, '') + result.data.ServerRelativeUrl;
    var body = '<span>Hi,</span><br><br>';
    body = body + '<span>' + that.state.selectedProvider + ' has uploaded a new invoice for reviewing.</span><br><br>';
    body = body + "<span>     - Click here to see the invoice: <a href='" + filePath + "'>" + result.data.Name + "</a></span><br><br>";
    body = body + "<span style='color:red;font-size:12px;'>Note:You must use a supported browser to access the above document</span>";

    const emailProps: IEmailProperties = {
      To: to,
      Subject: 'New invoice submitted by ' + that.state.selectedProvider,
      Body: body,
      AdditionalHeaders: {
        "content-type": "text/html",
      },
    };
    sp.utility.sendEmail(emailProps);
  }
  notifyUsers = (to, that, result) => {

    var SendmailID = that.props.currentContext.pageContext.user.email;
    var currentUser = that.props.currentContext.pageContext.user.displayName

    if (SendmailID.indexOf("#EXT#") >= 0) {
      SendmailID = SendmailID.split("#")[0].replace("_", "@");
    }
    var filePath = that.props.currentContext.pageContext.web.absoluteUrl.replace(that.props.currentContext.pageContext.web.serverRelativeUrl, '') + result.data.ServerRelativeUrl;
    var body = "<span>Hi <span>" + currentUser + ",<br><br>";
    body = body + "<span>The file below has been successfully uploaded to the " + that.state.selectedPath + " folder.</span><br>";
    body = body + "<span>Thank you for your submission.</span><br><br>"
    body = body + "<span>Click here to see the file: <a href='" + filePath + "'>" + result.data.Name + "</a><br><br>";
    body = body + "<span style='color:red;font-size:12px;'>Note:You must use a supported browser to access the above document</span>";

    const emailProps: IEmailProperties = {
      To: [SendmailID],
      Subject: 'New document successfully uploaded to BBHC Provider SharePoint',
      Body: body,
      AdditionalHeaders: {
        "content-type": "text/html",
      },
    };
    sp.utility.sendEmail(emailProps);
    location.reload();
  }

  public render(): React.ReactElement<IProviderDocumentsProps> {
    // const dropdownChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    //   if (item) {
    //     var dropDownIndex = parseInt(event.target["id"]);
    //     var destinationPath = this.state.destinationPath;
    //     var destinationFound = false;
    //     for (let d = 0; d < destinationPath.length; d++) {
    //       if (destinationPath[d].index == dropDownIndex) {
    //         destinationPath[d].value = item.text;
    //         destinationFound = true;
    //         break;
    //       }
    //     }
    //     if (!destinationFound) {
    //       destinationPath.push({
    //         index: dropDownIndex,
    //         value: item.text
    //       });
    //     }
    //     var stateFolder = this.state.folders;
    //     var removeIndexes = [];
    //     for (let index = dropDownIndex + 1; index < this.state.folders.length; index++) {
    //       removeIndexes.push(index);
    //     }
    //     for (var i = removeIndexes.length - 1; i >= 0; i--) {
    //       stateFolder.splice(removeIndexes[i], 1);
    //       for (let d = 0; d < destinationPath.length; d++) {
    //         if (destinationPath[d].index == removeIndexes[i]) {
    //           destinationPath.splice(removeIndexes[i], 1);
    //         }
    //       }
    //     }
    //     this.setState({ folders: stateFolder, destinationPath: destinationPath });
    //     this.getFolders(item.key + '/' + item.text);
    //   }
    // };

    const dropdownChange = (
      event: React.FormEvent<HTMLDivElement>,
      item: IDropdownOption
    ): void => {
      // if (!this.generalSubmissionChanged) {
      //   this.generalSubmissionChanged = true;
      //   var folders = this.state.folders;
      //   var gindex = -1;
      //   for (let index = 0; index < folders.length; index++) {
      //     const folder = folders[index];
      //     if (folder.text == "General Submissions") {
      //       gindex = index;
      //       break;
      //     }
      //   }
      //   if (gindex >= 0) {
      //     var data = folders[gindex];
      //     folders.splice(gindex, 1);
      //     folders.splice(folders.length, 0, data);
      //     this.setState({ folders: folders });
      //   }
      // }
      this.setState({ selectedPath: item.key.toString() });
    };

    const providerChange = (
      event: React.FormEvent<HTMLDivElement>,
      item: IDropdownOption
    ): void => {
      var providerName = item.key.toString();
      var data = this.state.allData.filter((c) => c.Title == providerName);
      if (data.length > 0) {
        this.loadUploadFolders(data[0].TemplateType);
      }
      this.setState({ selectedProvider: providerName });
    };

    return (
      <div>
        <h2 style={{ fontFamily: "Poppins, sans-serif" }}>Add File</h2>
        <div className={styles.d_flex}>
          <div>
            {
              <Dropdown
                placeholder="Select an provider"
                label="Providers"
                options={this.state.allProviders}
                onChange={providerChange}
                style={{ width: "300px" }}
                styles={dropDownStyles}
                className={styles.input_field}
              />
            }
            {
              <Dropdown
                placeholder="Select an option"
                label="Submission Types"
                options={this.state.folders}
                onChange={dropdownChange}
                style={{ width: "300px" }}
                styles={dropDown2Styles}
                className={styles.input_field}
              />
            }
            <input
              type="file"
              name="UploadedFile"
              id={fileId}
              onChange={(e) => this.fileUpload.call(this, e)}
              style={{ display: "none" }}
              className={styles.input_field}
            />
            <Label htmlFor={fileId} style={{ width: "300px" }}>
              <Label
                style={{
                  padding: "5px",
                  fontFamily: "Poppins, sans-serif",
                  width: "150px",
                }}
              >
                Attach File
              </Label>
              <div className={styles.files_upload}>
                <Image
                  styles={{ image: { padding: "5px" } }}
                  src={require("./Attach.png")}
                ></Image>
                <Label
                  style={{
                    padding: "5px",
                    fontFamily: "Poppins, sans-serif",
                  }}
                >
                  {this.state.fileName}
                </Label>
              </div>
            </Label>

            {/*<div>
              {this.state.previousyeardata.map((provider) => {
                return (
                  <div>
                    <Link href={provider.URL} target="_blank">
                      {provider.Title}
                    </Link>
                    <br></br>
                  </div>
                );
              })}
            </div>*/}

            <PrimaryButton
              text="Upload"
              onClick={this.uploadFile.bind(this)}
              className={styles.primary_button}
            />
          </div>
          <div style={{ marginLeft: "40px" }}>
            {this.state.selectedPath
              .toLocaleLowerCase()
              .indexOf(this.generalSubmission) >= 0 ? (
              <TextField
                required
                label="Notes"
                multiline
                rows={3}
                onChange={(e) => this.inputChangeHandler.call(this, e)}
                value={this.state.notes}
                name="notes"
                className={styles.notesinput_field}
              ></TextField>
            ) : (
              ""
            )}
            </div>
        </div>
      </div>
    );
  }
}
