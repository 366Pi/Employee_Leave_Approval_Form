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

export interface IDetailsListBasicExampleItem {
  Name: string;
  Designation: string;
  Department: string;
  "Leave From": string;
  "Leave Till": string;
  "Total Days": string;
  "Leave Type": string;
}

export default class LeaveApproval extends React.Component<
  ILeaveApprovalProps,
  {
    items: IDetailsListBasicExampleItem[];
    selectionDetails: string;
    isModalOpen: boolean;
    selectedItem;
  }
> {
  public handleDropdownChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    this.setState({
      selectedItem: item.key,
    });
  };

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

  // const [isModalOpen, { setTrue: showModal, setFalse: hideModal }] = useBoolean(false);
  // const [isDraggable, { toggle: toggleIsDraggable }] = useBoolean(false);
  // const [keepInBounds, { toggle: toggleKeepInBounds }] = useBoolean(false);
  // // Normally the drag options would be in a constant, but here the toggle can modify keepInBounds
  // const dragOptions = React.useMemo(
  //   (): IDragOptions => ({
  //     moveMenuItemText: 'Move',
  //     closeMenuItemText: 'Close',
  //     menu: ContextualMenu,
  //     keepInBounds,
  //   }),
  //   [keepInBounds],
  // );

  private _selection: Selection;
  private _allItems: IDetailsListBasicExampleItem[];
  private _columns: IColumn[];
  private _selMode: IDetailsListProps;

  private cancelIcon: IIconProps = { iconName: "Cancel" };
  private theme = getTheme();
  private contentStyles = mergeStyleSets({
    container: {
      display: "flex",
      flexFlow: "column nowrap",
      alignItems: "stretch",
    },
    header: [
      // eslint-disable-next-line deprecation/deprecation
      // theme.fonts.xLargePlus,
      {
        flex: "1 1 auto",
        borderTop: `4px solid ${this.theme.palette.themePrimary}`,
        color: this.theme.palette.neutralPrimary,
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

  constructor(props: ILeaveApprovalProps, state: any) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () =>
        this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    // Populate with items for demos.
    this._allItems = [];

    // for (let i = 0; i < 200; i++) {
    //   this._allItems.push({
    //     Name: "Item " + i,
    //     Designation: "Intern",
    //     Department: "Development",
    //     "Leave From": "26/7/2021",
    //     "Leave Till": "28/7/2021",
    //     "Total Days": "2",
    //     "Leave Type": "CL",
    //   });
    // }

    // Hardcoding 2 list items
    this._allItems.push({
      Name: "Test Employee 1",
      Designation: "Nurse",
      Department: "Emergency",
      "Leave From": "26/7/2021",
      "Leave Till": "26/7/2021",
      "Total Days": "1",
      "Leave Type": "CL",
    });

    this._allItems.push({
      Name: "Test Empployee 2",
      Designation: "Wardboy",
      Department: "OPD",
      "Leave From": "26/7/2021",
      "Leave Till": "27/7/2021",
      "Total Days": "2",
      "Leave Type": "EL",
    });

    this._columns = [
      {
        key: "column1",
        name: "Name",
        fieldName: "Name",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column2",
        name: "Designation",
        fieldName: "Designation",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column3",
        name: "Department",
        fieldName: "Department",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column4",
        name: "Leave From",
        fieldName: "Leave From",
        minWidth: 75,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column5",
        name: "Leave Till",
        fieldName: "Leave Till",
        minWidth: 70,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column6",
        name: "Total Days",
        fieldName: "Total Days",
        minWidth: 70,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column7",
        name: "Leave Type",
        fieldName: "Leave Type",
        minWidth: 70,
        maxWidth: 200,
        isResizable: true,
        isMultiline: true,
      },
    ];

    this.state = {
      items: this._allItems,
      selectionDetails: this._getSelectionDetails(),
      isModalOpen: false,
      selectedItem: "carrot",
    };
  }

  private hideModal = () => {
    this.setState({
      isModalOpen: false,
    });
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

                {/* Name, Department, Desgignation, Email */}
                <div className="row top-buffer">
                  <div className="col-sm-4">
                    <div className="form-group">
                      <TextField
                        label="Name"
                        readOnly
                        defaultValue="Test Employee 1"
                      />
                      <TextField
                        label="Designation"
                        readOnly
                        defaultValue="Nurse"
                      />
                    </div>
                  </div>
                  <div className="col-sm-4">
                    <div className="form-group">
                      <TextField
                        label="Department"
                        readOnly
                        defaultValue="OPD"
                      />

                      <TextField
                        label="Email"
                        readOnly
                        defaultValue="TestUser1@healthPoint.com"
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
                      defaultValue="1234567890"
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
                        value={new Date()}
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
                        value={new Date()}
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
                        value={new Date()}
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
                        value={"10"}
                      />
                    </div>
                  </div>
                  <div className="col-lg-6">
                    <div className="form-group">
                      <TextField
                        readOnly={true}
                        label="Applied Leave Type"
                        value="ML"
                      />
                    </div>
                  </div>
                </div>

                {/* Employee Leave Purpose */}
                <div className="row top-buffer">
                  <div className="col-lg-12">
                    <div className="form-group">
                      <TextField
                        readOnly
                        label={"Employee Leave Purpose"}
                        multiline
                        autoAdjustHeight
                        value={
                          "Need Urgent Leave because Lorem ipsum dolor sit amet,  \
                          consectetur adipiscing elit. Pellentesque eu euismod dui. \
                          Fusce sollicitudin mauris leo, et pharetra urna porttitor quis. \
                          Quisque augue massa, varius suscipit vehicula in, accumsan nec velit. \
                          Cras vulputate purus velit, quis sagittis mi volutpat vel. \
                          In sed convallis turpis. "
                        }
                        style={{ minWidth: 500 }}
                      />
                    </div>
                  </div>
                </div>

                {/* Total leaves left and type of leave in table form */}
                <div className="row top-buffer">
                  {/* <div className="col-lg-6">
                    <div className="form-group">
                      <br />
                      <label htmlFor="txtName">Total Leaves Left : </label>
                      {60}
                      {this.state.Completed_Activities}/
                            {this.state.Total_Activities}
                    </div>
                  </div> */}
                  <div className="col-lg-6">
                    <div className="form-group">
                      <table className="table table-borderless">
                        <thead>
                          <tr>
                            <th scope="col">Leave Type</th>
                            <th scope="col">Number</th>
                            {/* <th scope="col">Last</th>
                              <th scope="col">Handle</th> */}
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <th scope="row">CL</th>
                            <td>10</td>
                            {/* <td>Otto</td>
                              <td>@mdo</td> */}
                          </tr>
                          <tr>
                            <th scope="row">SL</th>
                            <td>5</td>
                            {/* <td>Thornton</td>
                              <td>@fat</td> */}
                          </tr>
                          <tr>
                            <th scope="row">PH</th>
                            <td>30</td>
                            {/* <td>@twitter</td> */}
                          </tr>
                          <tr>
                            <th scope="row">EL</th>
                            <td>10</td>
                          </tr>
                          <tr>
                            <th scope="row">Comm. Off</th>
                            <td>5</td>
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
                        // required={true}
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
                      />
                    </div>
                  </div>
                </div>

                {/* Remarks */}
                <div className="row top-buffer">
                  <div className="col-lg-12">
                    <div className="form-group">
                      <TextField label={"Remarks"} multiline />
                    </div>
                  </div>
                </div>

                {/* Submit button */}
                <div className="row top-buffer">
                  <div className="col-lg-12 text-center">
                    <br />
                    <PrimaryButton
                      text="Submit"
                      // onClick={_alertClicked}
                      allowDisabledFocus
                      // disabled={disabled}
                      // checked={checked}
                    />
                  </div>
                </div>
              </div>
            </div>
          </div>
        </Modal>
      </div>
    );
  }

  // depreciated method no longer needed
  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return "No items selected";
      case 1:
        return (
          "1 item selected: " +
          (this._selection.getSelection()[0] as IDetailsListBasicExampleItem)
            .Name
        );
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
        ? this._allItems.filter((i) => i.Name.toLowerCase().indexOf(text) > -1)
        : this._allItems,
    });
  };

  private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
    // alert(`Item invoked: ${item.Name}`);
    this.setState({
      isModalOpen: true,
    });
  };

  private _getPeoplePickerItems(items: any[]) {
    console.log("Items:", items);
  }
}
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
