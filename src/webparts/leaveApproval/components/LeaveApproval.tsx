import * as React from "react";
import { useId, useBoolean } from "@fluentui/react-hooks";
import styles from "./LeaveApproval.module.scss";
import { ILeaveApprovalProps } from "./ILeaveApprovalProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  IDetailsListProps,
  SelectionMode,
} from "@fluentui/react/lib/DetailsList";
import { MarqueeSelection } from "@fluentui/react/lib/MarqueeSelection";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import { Text } from "@fluentui/react/lib/Text";
import { TextField, ITextFieldStyles } from "@fluentui/react/lib/TextField";
import { Announced } from "@fluentui/react/lib/Announced";
import {
  Modal,
  IDragOptions,
  mergeStyleSets,
  getTheme,
  FontWeights,
  IIconProps,
  DatePicker,
  defaultDatePickerStrings,
  PrimaryButton,
  IconButton,
  IButtonStyles,
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption,
  IDropdownStyles,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import { Image, IImageProps, ImageFit } from "@fluentui/react/lib/Image";
import {
  IBasePickerSuggestionsProps,
  NormalPeoplePicker,
  ValidationState,
} from "@fluentui/react/lib/Pickers";
import { people, mru } from "@fluentui/example-data";

import * as $ from "jquery";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { BasePeoplePicker } from "office-ui-fabric-react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

//adding funtional workflow req imports
import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { Web } from "@pnp/sp/webs";
import { sp } from "@pnp/sp";
import { IItemAddResult } from "@pnp/sp/items";

import "@pnp/sp/lists";
import "@pnp/sp/items";

import CommOff from "./CompOff";

SPComponentLoader.loadCss(
  "https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css"
);
SPComponentLoader.loadCss(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"
);

require("bootstrap");

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

const dropdownControlledExampleOptions = [
  { key: "fw_to_HA", text: "Forward to Higher Authorities" },
  { key: "fw_to_HR", text: "Forward to HR" },
  { key: "accepted", text: "Accepted" },
  { key: "rejected", text: "Rejected" },
];

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "Suggested People",
  mostRecentlyUsedHeaderText: "Suggested Contacts",
  noResultsFoundText: "No results found",
  loadingText: "Loading",
  showRemoveButtons: true,
  suggestionsAvailableAlertText: "People Picker Suggestions available",
  suggestionsContainerAriaLabel: "Suggested contacts",
};

const imageProps: Partial<IImageProps> = {
  imageFit: ImageFit.centerContain,
  width: 150,
  height: 150,
  src: "http://via.placeholder.com/250x150",
  // Show a border around the image (just for demonstration purposes)
  styles: (props) => ({
    root: { border: "1px solid " + props.theme.palette.neutralSecondary },
  }),
};

const exampleChildClass = mergeStyles({
  display: "block",
  marginBottom: "10px",
});

const textFieldStyles: Partial<ITextFieldStyles> = {
  root: { maxWidth: "300px" },
};

// modal and inside styles
const cancelIcon: IIconProps = { iconName: "Cancel" };
const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: "flex",
    flexFlow: "column nowrap",
    alignItems: "stretch",
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: "1 1 auto",
      borderTop: `4px solid ${theme.palette.themePrimary}`,
      // color: theme.palette.neutralPrimary,
      color: theme.palette.black,
      display: "flex",
      alignItems: "center",
      fontWeight: FontWeights.semibold,
      padding: "12px 12px 14px 24px",
    },
  ],
  body: {
    flex: "4 4 auto",
    padding: "0 24px 24px 24px",
    overflowY: "hidden",
    selectors: {
      p: { margin: "14px 0" },
      "p:first-child": { marginTop: 0 },
      "p:last-child": { marginBottom: 0 },
    },
  },
});
const iconButtonStyles: Partial<IButtonStyles> = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: "auto",
    marginTop: "4px",
    marginRight: "2px",
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};

export interface IDetailsListBasicExampleItem {
  // employee data --> coming from AD
  empName: string;
  empDesignation: string;
  empDepartment: string;
  empmobile: string;
  empEmail: string;

  // leave req data --> coming from leave req list
  Leave_From: String;
  Leave_Till: String;
  Total_Days: number;
  return_date: string;
  leave_type_id: Number;
  leave_type_text: string; // have to be fetched
  Purpose: string;
  commOffDate: string;
  commOffOccasion: string;
  leaveReqItem_Id: number;
}

export interface BalLeftBlueprintObj {
  CL: any;
  SL: any;
  EL: any;
  Comp_Off: any;
  Leave_Without_pay: any;
  ML: any;
  PL: any;
  Total: any;
}

export default class LeaveApproval extends React.Component<
  ILeaveApprovalProps,
  {
    items: IDetailsListBasicExampleItem[];
    isModalOpen: boolean;
    selectedItem: any;

    EmpName: any;
    EmpDepartment: any;
    EmpDesignation: any;
    EmpEmail: any;
    EmpMobile: any;

    ApproverEmail: any;
    ApproverId: any;
    reliever: any;
    selectedAction: any;
    remarks: any;
    relieverEmail: any;
    relieverId: any;

    leaveStartDate: any;
    leaveEndDate: any;
    returnDate: any;
    leavePurpose: any;
    leaveAppliedForDays: any;
    leaveType_Id: any;
    leaveType: any;
    CommpOffDate: any;
    CommpOffOccasion: any;
    CommpOffVisible: any;

    LeaveReqItem_Id: any;
    LeaveMasterItem_Id: any;

    // balance leaves
    balLeavesObj: BalLeftBlueprintObj;

    isIncomplete: any;
  }
> {
  w = Web(this.props.webUrl + "/sites/Maitri");
  url = location.search;
  params = new URLSearchParams(this.url);
  id = this.params.get("spid");

  // code to initalize jquery
  private GetIPAddress(): void {
    var call = $.ajax({
      url: "https://api.ipify.org/?format=json",
      method: "GET",
      async: false,
      dataType: "json",
      success: (data) => {
        console.log("IP Address : " + data.ip);
        // ipaddress = data.ip;
      },
      error: (textStatus, errorThrown) => {
        console.log(
          "Ip Address fetch failed : " + textStatus + "--" + errorThrown
        );
      },
    }).responseJSON;
  }

  private _selection: Selection;
  private _allItems: IDetailsListBasicExampleItem[];
  private _columns: IColumn[];
  private _selMode: IDetailsListProps;
  private _temp: BalLeftBlueprintObj = {
    CL: 0,
    SL: 0,
    EL: 0,
    Comp_Off: 0,
    Leave_Without_pay: 0,
    ML: 0,
    PL: 0,
    Total: 0,
  };

  constructor(props: ILeaveApprovalProps, state: any) {
    super(props);

    // Populate with items for demos.
    this._allItems = [];

    // for (let i = 0; i < 200; i++) {
    //   this._allItems.push({
    //     Name: "Item " + i,
    //     Designation: "Intern",
    //     Department: "Development",
    //     "Leave From": "26/7/2021",
    //     "Leave_Till": "28/7/2021",
    //     Total_Days: "2",
    //     "Leave Type": "CL",
    //   });
    // }

    // Hardcoding 2 list items
    // this._allItems.push({
    //   Name: "Test Employee 1",
    //   Designation: "Nurse",
    //   Department: "Emergency",
    //   Leave_From: "27/8/2021",
    //   Leave_Till: "27/8/2021",
    //   Total_Days: "1",
    //   "Leave Type": "CL",
    //   return_date: "28/8/2021",
    //   Purpose: "Test Purpose 1",
    //   leave_type_id: 1,
    //   ExtID: 1,
    // });

    // this._allItems.push({
    //   Name: "Test Empployee 2",
    //   Designation: "Wardboy",
    //   Department: "OPD",
    //   Leave_From: "27/8/2021",
    //   Leave_Till: "28/8/2021",
    //   Total_Days: "2",
    //   "Leave Type": "EL",
    //   return_date: "29/8/2021",
    //   Purpose: "Test Purpose 2",
    //   leave_type_id: 2,
    //   ExtID: 2,
    // });

    this._columns = [
      {
        key: "column1",
        name: "Name",
        fieldName: "empName",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column2",
        name: "Designation",
        fieldName: "empDesignation",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column3",
        name: "Department",
        fieldName: "empDepartment",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column4",
        name: "Leave From",
        fieldName: "Leave_From",
        minWidth: 75,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column5",
        name: "Leave Till",
        fieldName: "Leave_Till",
        minWidth: 70,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column6",
        name: "Total Days",
        fieldName: "Total_Days",
        minWidth: 70,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column7",
        name: "Leave Type",
        fieldName: "leave_type_text",
        minWidth: 70,
        maxWidth: 200,
        isResizable: true,
        isMultiline: true,
      },
    ];

    this.state = {
      items: this._allItems,
      isModalOpen: false,
      selectedItem: undefined,

      EmpName: undefined,
      EmpDepartment: undefined,
      EmpDesignation: undefined,
      EmpEmail: undefined,
      EmpMobile: undefined,

      ApproverEmail: undefined,
      ApproverId: undefined,
      reliever: undefined,
      selectedAction: undefined,
      remarks: undefined,
      relieverEmail: undefined,
      relieverId: undefined,

      leaveStartDate: undefined,
      leaveEndDate: undefined,
      returnDate: undefined,
      leavePurpose: undefined,
      leaveAppliedForDays: undefined,
      leaveType_Id: undefined,
      leaveType: undefined,
      CommpOffDate: undefined,
      CommpOffOccasion: undefined,
      CommpOffVisible: false,

      LeaveReqItem_Id: undefined,
      balLeavesObj: this._temp,
      LeaveMasterItem_Id: undefined,

      isIncomplete: false,
    };
  }

  private hideModal = () => {
    this.setState({
      isModalOpen: false,
    });
  };

  private handleDropdownChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    this.setState(
      {
        selectedItem: item.key,
        selectedAction: item.text,
      },
      () => {
        console.log(this.state.selectedItem);
      }
    );
  };

  private handelRemarksChange = (event) => {
    this.setState({ remarks: event.target.value });
  };

  private handleErrorMessage = () => {
    this.setState({ isIncomplete: false });
  };

  public render(): React.ReactElement<ILeaveApprovalProps> {
    return (
      <div>
        {/* <div className={exampleChildClass}>{this.state.selectionDetails}</div> */}
        {/* <TextField
          className={exampleChildClass}
          label="Filter by name:"
          onChange={this._onFilter}
          styles={textFieldStyles}
        /> */}
        {/* <Announced
          message={`Number of items after filter applied: ${this.state.items.length}.`}
        /> */}
        {/* <MarqueeSelection selection={this._selection}>
          
        </MarqueeSelection> */}
        <DetailsList
          items={this.state.items}
          columns={this._columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          selection={this._selection}
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          checkButtonAriaLabel="select row"
          onItemInvoked={this._onItemInvoked}
        />
        <Modal
          // titleAriaId={titleId}
          isOpen={this.state.isModalOpen}
          onDismiss={this.hideModal}
          isBlocking={false}
          containerClassName={contentStyles.container}
          dragOptions={undefined}
        >
          <div className={contentStyles.header}>
            <IconButton
              styles={iconButtonStyles}
              iconProps={cancelIcon}
              ariaLabel="Close popup modal"
              onClick={this.hideModal}
            />
          </div>
          <div className={contentStyles.body}>
            <div className="panel panel-default">
              <div className="panel-body">
                {/* Name, Department, Desgignation, ApproverEmail */}
                <div className="row top-buffer">
                  <div className="col-sm-4">
                    <div className="form-group">
                      <TextField
                        label="Name"
                        readOnly
                        // defaultValue="Test Employee 1"
                        value={this.state.EmpName}
                      />
                      <TextField
                        label="Designation"
                        readOnly
                        // defaultValue="Nurse"
                        value={this.state.EmpDesignation}
                      />
                    </div>
                  </div>
                  <div className="col-sm-4">
                    <div className="form-group">
                      <TextField
                        label="Department"
                        readOnly
                        // defaultValue="OPD"
                        value={this.state.EmpDepartment}
                      />

                      <TextField
                        label="Email"
                        readOnly
                        // defaultValue="TestUser1@healthPoint.com"
                        value={this.state.EmpEmail}
                      />
                    </div>
                  </div>
                  <div className="col-sm-4">
                    <div className="form-group">
                      <Image
                        {...imageProps}
                        alt="Example with no image fit value and no height or width is specified."
                      />
                    </div>
                  </div>
                </div>

                {/* current shift, manager in-charge, mobile no- TextField */}
                <div className="row top-buffer">
                  <div className="col-lg-6">
                    <div className="form-group">
                      <TextField
                        label="Current Shift"
                        readOnly
                        defaultValue="Day"
                      />
                      <TextField
                        label="Manager In-Charge"
                        readOnly
                        defaultValue="Test Employee 2"
                      />
                    </div>
                  </div>
                  <div className="col-lg-6">
                    <TextField
                      label="Mobile No"
                      readOnly
                      // defaultValue="1234567890"
                      value={this.state.EmpMobile}
                    />
                  </div>
                </div>

                {/* Leave Start and End date read-only DatePicker col-6 */}
                <div className="row top-buffer">
                  <div className="col-lg-6">
                    <div className="form-group">
                      <DatePicker
                        label="Leave Start Date"
                        placeholder="Select a date..."
                        ariaLabel="Start Date"
                        // DatePicker uses English strings by default. For localized apps, you must override this prop.
                        strings={defaultDatePickerStrings}
                        value={new Date(this.state.leaveStartDate)}
                        disabled
                      />
                    </div>
                  </div>
                  <div className="col-lg-6">
                    <div className="form-group">
                      <DatePicker
                        label="Leave End Date"
                        placeholder="Select a date..."
                        ariaLabel="End Date"
                        // DatePicker uses English strings by default. For localized apps, you must override this prop.
                        strings={defaultDatePickerStrings}
                        value={new Date(this.state.leaveEndDate)}
                        disabled
                      />
                    </div>
                  </div>
                </div>

                {/* Returning on Datepicker readonly col-6 */}
                <div className="row top-buffer">
                  <div className="col-lg-6">
                    <div className="form-group">
                      <DatePicker
                        label="Returning On"
                        placeholder="Select a date..."
                        ariaLabel="Return Date"
                        // DatePicker uses English strings by default. For localized apps, you must override this prop.
                        strings={defaultDatePickerStrings}
                        value={new Date(this.state.returnDate)}
                        disabled
                      />
                    </div>
                  </div>
                </div>

                {/* Leave applied for days,  Applied leave type */}
                <div className="row top-buffer">
                  <div className="col-lg-6">
                    <div className="form-group">
                      <TextField
                        readOnly={true}
                        label="Leave applied for Days"
                        value={this.state.leaveAppliedForDays}
                      />
                    </div>
                  </div>
                  <div className="col-lg-6">
                    <div className="form-group">
                      <TextField
                        readOnly={true}
                        label="Applied Leave Type"
                        value={this.state.leaveType}
                      />
                    </div>
                  </div>
                </div>

                {/* If comm. off then display this component */}
                {this.state.CommpOffVisible == true ? (
                  <CommOff
                    CommpOffDate={this.state.CommpOffDate}
                    CommpOffOccasion={this.state.CommpOffOccasion}
                  />
                ) : null}

                {/* Employee Leave Purpose */}
                <div className="row top-buffer">
                  <div className="col-lg-12">
                    <div className="form-group">
                      <TextField
                        readOnly
                        label={"Purpose of Leave"}
                        multiline
                        autoAdjustHeight
                        value={this.state.leavePurpose}
                        style={{ minWidth: 500 }}
                      />
                    </div>
                  </div>
                </div>

                {/* Total leaves left and type of leave in table form */}
                <div className="row top-buffer">
                  <div className="col-lg-6">
                    <div className="form-group">
                      <table className="table table-borderless table-hover">
                        <thead>
                          <tr>
                            <th scope="col">Leave Type</th>
                            <th scope="col">Number</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <th scope="row">CL</th>
                            <td>{this.state.balLeavesObj.CL}</td>
                          </tr>
                          <tr>
                            <th scope="row">SL</th>
                            <td>{this.state.balLeavesObj.SL}</td>
                          </tr>
                          <tr>
                            <th scope="row">EL</th>
                            <td>{this.state.balLeavesObj.EL}</td>
                          </tr>
                          <tr>
                            <th scope="row">Commp. Off</th>
                            <td>{this.state.balLeavesObj.Comp_Off}</td>
                          </tr>
                          <tr>
                            <th scope="row">Leave Without Pay</th>
                            <td>{this.state.balLeavesObj.Leave_Without_pay}</td>
                          </tr>
                          <tr>
                            <th scope="row">ML</th>
                            <td>{this.state.balLeavesObj.ML}</td>
                          </tr>
                          <tr>
                            <th scope="row">PL</th>
                            <td>{this.state.balLeavesObj.PL}</td>
                          </tr>
                          <tr className="table-primary">
                            <th scope="row">Total</th>
                            <td className="table-dark">
                              {this.state.balLeavesObj.Total}
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>

                {/* will be releived by: people picker  */}
                <div className="row top-buffer">
                  <div className="col-lg-6">
                    <div className="form-group">
                      <PeoplePicker
                        context={this.props.context}
                        titleText="Reliever"
                        personSelectionLimit={1}
                        groupName={""} // Leave this blank in case you want to filter from all users
                        showtooltip={true}
                        required={true}
                        // disabled={true}
                        onChange={this._getPeoplePickerItems}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                      />
                    </div>
                  </div>
                </div>

                {/* Application Status: Dropdown  */}
                <div className="row top-buffer">
                  <div className="col-lg-6">
                    <div className="form-group">
                      <label>Application Status : </label>
                      <Dropdown
                        // label="Leave Type"
                        selectedKey={this.state.selectedItem}
                        // eslint-disable-next-line react/jsx-no-bind
                        onChange={this.handleDropdownChange}
                        placeholder="Select an option"
                        options={dropdownControlledExampleOptions}
                        styles={dropdownStyles}
                        required
                      />
                    </div>
                  </div>
                </div>

                {/* Remarks */}
                <div className="row top-buffer">
                  <div className="col-lg-12">
                    <div className="form-group">
                      <TextField
                        label={"Remarks"}
                        multiline
                        onChange={this.handelRemarksChange}
                        required
                      />
                    </div>
                  </div>
                </div>

                {/* Submit button */}
                <div className="row top-buffer">
                  <div className="col-lg-12 text-center">
                    <br />
                    <PrimaryButton
                      text="Submit"
                      onClick={this._onSubmitInvoked}
                      allowDisabledFocus
                      // disabled={disabled}
                      // checked={checked}
                    />
                  </div>
                </div>

                {/* Error component due to incomplete action items */}
                <div className="row top-buffer">
                  <div className="col-lg-12 text-center">
                    <br />
                    {this.state.isIncomplete ? (
                      <MessageBar
                        messageBarType={MessageBarType.error}
                        isMultiline={false}
                        onDismiss={this.handleErrorMessage}
                        dismissButtonAriaLabel="Close"
                      >
                        Fill all the required feilds before submission.
                      </MessageBar>
                    ) : null}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </Modal>
      </div>
    );
  }

  public componentDidMount() {
    this.getApproverData();
  }

  private getEmpData = (obj: any) => {
    console.log(`${obj.Employee_Email}`);

    // make graph api call and store emp data in cols array
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api(`/users/${obj.Employee_Email}`)
          .select("displayName,jobTitle,department,mobilePhone,mail")
          .get()
          .then((res) => {
            console.log(
              `${res.displayName}, ${res.department}, ${res.mail}, ${res.mobilePhone}, ${res.jobTitle}\n`
            );

            // get the leave type text from leave_types table using its id
            this.w.lists
              .getByTitle("Leave_Types")
              .items.getById(obj.Leave_TypeId)
              .get()
              .then((item: any) => {
                this.initializeColData(obj, res, item.Title);
                // console.log(item.Leave_Type_Full);
              });
          })
          .catch((err) => {
            console.log("🔥 There was an error 🧯 ", err);
          });
      });
  };

  private getListItems = () => {
    this.w.lists
      .getByTitle("Leave_Requests")
      .items.get()
      .then((items: any[]) => {
        items.map((el) => {
          console.log("Inside getListItems: ", el.Status);
          if (
            el.Assigned_To_PersonId === this.state.ApproverId &&
            el.Status != "Accepted"
          ) {
            this.getEmpData(el);
          }
        });
      });
  };

  private _onSubmitInvoked = (): void => {
    const list = this.w.lists.getByTitle("Leave_Requests");
    const list2 = this.w.lists.getByTitle("Leave_Master");
    const leave_type = this.state.leaveType;
    console.log("when submit: ", leave_type);

    if (
      this.state.relieverId === undefined ||
      this.state.selectedAction === undefined ||
      this.state.remarks === undefined
    )
      this.setState({ isIncomplete: true });
    else {
      if (this.state.selectedAction === "Accepted") {
        // get a specific item by id
        this.w.lists
          .getByTitle("Leave_Master")
          .items.getById(this.state.LeaveMasterItem_Id)
          .get()
          .then((item: any) => {
            // modifying
            item[leave_type] =
              this.state.balLeavesObj[leave_type] -
              this.state.leaveAppliedForDays;

            // now updatating
            list2.items
              .getById(this.state.LeaveMasterItem_Id)
              .update(item)
              .then((i) => {
                console.log(i);
              });
            console.log(item);
          });
      }
      list.items
        .getById(this.state.LeaveReqItem_Id)
        .update({
          RelieverId: this.state.relieverId,
          Approver_Remarks: this.state.remarks,
          Status: this.state.selectedAction,
        })
        .then((i) => {
          console.log(i);
          if (this.state.selectedAction != "Accepted") {
            alert(
              `Leave Request has been successfully ${this.state.selectedAction}`
            );
          } else {
            alert("Approved leaves, deducted from leave master.");
            window.location.reload();
          }
          this.setState({ isModalOpen: false });
        });
    }
  };

  private _checkIfCommpOff = (): void => {
    if (this.state.leaveType == "Compensatory Off")
      this.setState({ CommpOffVisible: true });
    else this.setState({ CommpOffVisible: false });
  };

  private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
    // alert(`Item invoked: ${item.Name}`);
    this.setState(
      {
        isModalOpen: true,

        EmpName: item.empName,
        EmpDepartment: item.empDepartment,
        EmpDesignation: item.empDesignation,
        EmpEmail: item.empEmail,
        EmpMobile: item.empmobile,

        leaveStartDate: item.Leave_From,
        leaveEndDate: item.Leave_Till,
        returnDate: item.return_date,
        leaveAppliedForDays: item.Total_Days,
        leaveType: item.leave_type_text,
        leavePurpose: item.Purpose,
        CommpOffDate: item.commOffDate,
        CommpOffOccasion: item.commOffOccasion,

        LeaveReqItem_Id: item.leaveReqItem_Id,
      },

      () => {
        console.log(
          "in onInvoke: ",
          this.state.CommpOffDate,
          " ",
          this.state.CommpOffOccasion,
          " ",
          this.state.leaveType
        );
        this._checkIfCommpOff();
        this.getBalLeaveData();
      }
    );

    // console.log(item);
    // this.getDataInsideModal(item.ExtID.valueOf());
  };

  private getBalLeaveData = (): void => {
    /*
    fetch the leave master table, search the logged in employeeId.
    and store the items in leaveBalanceLeft array state.

    should be called only after empId state is set
    */

    this.w.lists
      .getByTitle("Leave_Master")
      .items.get()
      .then((items: any[]) => {
        for (let i = 0; i < items.length; i++) {
          console.log("Inside leave_master ", items[i]);
          if (items[i].Title == this.state.EmpEmail) {
            const temp: BalLeftBlueprintObj = {
              CL: items[i].CL,
              SL: items[i].SL,
              EL: items[i].EL,
              Comp_Off: items[i].Comp_Off,
              Leave_Without_pay: items[i].Leave_Without_Pay,
              ML: items[i].ML,
              PL: items[i].PL,
              Total:
                items[i].CL +
                items[i].SL +
                items[i].EL +
                items[i].Comp_Off +
                items[i].Leave_Without_Pay +
                items[i].ML +
                items[i].PL,
            };
            this.setState(
              { balLeavesObj: temp, LeaveMasterItem_Id: items[i].Id },
              () => {
                console.log(
                  "Found! ",
                  items[i].Employee_ID,
                  " ",
                  this.state.balLeavesObj
                );
              }
            );
            break;
          }
        }
      });
  };

  private _getPeoplePickerItems = (items: any[]) => {
    console.log("Reliever ExtID:", items[0].secondaryText);
    const email = items[0].secondaryText;
    const obj = this.GetUserId(items[0].secondaryText);

    this.setState({ relieverEmail: email, relieverId: obj.d.Id }, () => {
      console.info(this.state.relieverEmail, " ", this.state.relieverId);
    });
  };

  // Pass logged in user's emailID to this function to get his userID
  // which will be pushed to the EmployeeId list col
  private GetUserId(userName) {
    var siteUrl = this.props.webUrl + "/sites/Maitri";

    // console.log("siteUrl", siteUrl);

    var enclogin = encodeURIComponent(userName);

    var call = $.ajax({
      // url:
      //   siteUrl +
      //   "/_api/web/siteusers/getbyloginname(@v)?@v=%27" +
      //   enclogin +
      //   "%27",

      url:
        siteUrl +
        "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|" +
        userName +
        "%27",

      method: "GET",

      headers: { Accept: "application/json; odata=verbose" },

      async: false,

      dataType: "json",
    }).responseJSON;

    // console.log("Call : " + JSON.stringify(call));

    return call;
  }

  private getApproverData = (): void => {
    // Makes a graph api call to fetch logged in user's data from Azure AD

    // preventDefault();
    // console.log("webpart context is: ", this.props.context);

    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("/me")
          .select("mail")
          .get()
          .then((res) => {
            // console.log(
            //   `${res.displayName}, ${res.department}, ${res.mail}, ${res.mobilePhone}, ${res.employeeId}`
            // );
            this.setState(
              {
                ApproverEmail: res.mail,
              },
              () => {
                const obj = this.GetUserId(this.state.ApproverEmail);
                this.setState({ ApproverId: obj.d.Id }, () => {
                  console.log("EmployeeId: ", this.state.ApproverId);
                  this.getListItems();
                });
              }
            );
          })
          .catch((err) => {
            console.log("🔥 There was an error 🧯 ", err);
          });
      });
  };

  private GetUserDetails(userId) {
    //userName format = i:0#.w|bidev\sp_admin
    var siteUrl = this.props.webUrl + "/sites/Maitri";
    //console.log("Site URL : " + siteUrl + "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|"+userName+"%27");
    var call = $.ajax({
      url: siteUrl + "/_api/web/getuserbyid(" + parseInt(userId) + ")",
      method: "GET",
      headers: { Accept: "application/json; odata=verbose" },
      async: false,
      dataType: "json",
    }).responseJSON;
    // console.log("Call : " + JSON.stringify(call));
    return call;
  }

  private initializeColData = (obj: any, res: any, LeaveTypeTxt: any) => {
    this.setState({
      items: [
        ...this.state.items,
        {
          empName: res.displayName,
          empDesignation: res.jobTitle,
          empDepartment: res.department,
          empEmail: res.mail,
          empmobile: res.mobilePhone,

          Leave_From: new Date(obj.Leave_From).toLocaleDateString("en-US"),
          Leave_Till: new Date(obj.Leave_To).toLocaleDateString("en-US"),
          Total_Days: Math.floor(obj.No_of_days),
          return_date: obj.Return_On,
          leave_type_id: obj.Leave_TypeId,
          Purpose: obj.Purpose,
          leave_type_text: LeaveTypeTxt,
          commOffDate: obj.Compoff_against_date,
          commOffOccasion: obj.Compoff_occasion,
          leaveReqItem_Id: obj.Id,
        },
      ],
    });
  };
}
