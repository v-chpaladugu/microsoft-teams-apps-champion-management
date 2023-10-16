import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as microsoftTeams from '@microsoft/teams-js';
import * as React from 'react';
import * as stringsConstants from '../constants/strings';
import BootstrapTable from 'react-bootstrap-table-next';
import Col from 'react-bootstrap/esm/Col';
import moment from 'moment';
import Row from 'react-bootstrap/Row';
import { Component } from 'react';
import { Icon, initializeIcons } from 'office-ui-fabric-react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '../scss/Champions.scss';
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";

initializeIcons();

let currentUserName: string;

interface ChampionEventsProps {
    context: WebPartContext;
    callBack?: Function;
    filteredAllEvents: Array<any>;
    selectedMemberDetails?: any;
    parentComponent: string;
    selectedMemberID: string;
    loggedinUserEmail?: string;
}
interface ChampionEventsState {
    filteredUserActivities: Array<any>;
    userActivitiesPerPage: number;
    pageNumber: number;
    selectedUserActivities: Array<any>;
    selectedMemberDetails: Array<any>;
}

export default class ChampionEvents extends Component<
    ChampionEventsProps,
    ChampionEventsState
> {
    constructor(props: any) {
        super(props);
        this.state = {
            filteredUserActivities: [],
            userActivitiesPerPage: 5,
            pageNumber: 1,
            selectedUserActivities: [],
            selectedMemberDetails: []
        };

        currentUserName = this.props.context.pageContext.user.displayName;
        this._renderListAsync();
    }

    //Initializes the teams library and calling the methods to load the initial data  
    public _renderListAsync() {
        microsoftTeams.initialize();
    }

    //method to load the selected member activities
    public componentDidMount() {
        this.getChampionActivities(this.props.selectedMemberID);
    }

    //This method will be called whenever there is an update to the component
    public componentDidUpdate(prevProps: Readonly<ChampionEventsProps>, prevState: Readonly<ChampionEventsState>, snapshot?: any): void {

        if (prevProps.selectedMemberID !== this.props.selectedMemberID ||
            prevProps.filteredAllEvents !== this.props.filteredAllEvents ||
            prevProps.selectedMemberDetails !== this.props.selectedMemberDetails) {
            this.getChampionActivities(this.props.selectedMemberID);
            this.setState({ pageNumber: 1 })
        }
        if (prevState.selectedUserActivities.length !== this.state.selectedUserActivities.length || prevState.pageNumber !== this.state.pageNumber) {
            this.updatefilteredUserActivities();
        }
    }

    //Filtering the records based on page size for each page from total activities of the member
    private updatefilteredUserActivities = () => {

        const filteredData = this.state.selectedUserActivities.filter((activity, idx) => {
            return (idx >= (this.state.userActivitiesPerPage * this.state.pageNumber - this.state.userActivitiesPerPage) && idx < (this.state.pageNumber * this.state.userActivitiesPerPage));
        });
        this.setState({ filteredUserActivities: filteredData });
    }

    //Method to execute the deep link API in teams
    public openTask = (selectedTask: string) => {
        microsoftTeams.initialize();
        microsoftTeams.executeDeepLink(selectedTask);
    }

    //Default image to show in case of any error in loading user profile image
    public addDefaultSrc(ev: any) {
        ev.target.src = require("../assets/images/noprofile.png");
    }

    //get data for the activities table
    private async getChampionActivities(selectedChampion: string) {
        let memberActivitesArray: any = [];
        let memberDetails: any = [];
        this.setState({
            selectedUserActivities: [],
            selectedMemberDetails: []
        });
        if (selectedChampion !== stringsConstants.AllLabel && selectedChampion !== "") {

            //filtering the selected member's data from the array of all records
            let selectedMemberEvents = this.props.filteredAllEvents.filter((item) => item.MemberId === selectedChampion);

            //creating an array to store the required data for Activities table
            selectedMemberEvents.forEach((event) => {
                memberActivitesArray.push({
                    DateofEvent: event[stringsConstants.dateOfEventLabel] ? moment(new Date(event[stringsConstants.dateOfEventLabel])).format("MMMM Do, YYYY") : moment(new Date()).format("MMMM Do, YYYY"),
                    Type: event["EventName"] ? event["EventName"] : "",
                    Points: event["Count"] ? event["Count"] : ""
                });
            });

            //creating an array to store the required data to show member details
            if (this.props.parentComponent === stringsConstants.ChampionReportLabel) {
                memberDetails.push({
                    ID: this.props.selectedMemberDetails[0].ID,
                    Title: this.props.selectedMemberDetails[0].Title,
                    FirstName: this.props.selectedMemberDetails[0].FirstName,
                    LastName: this.props.selectedMemberDetails[0].LastName,
                });
            }
            else if (this.props.parentComponent === stringsConstants.ChampionsCardsLabel) {
                memberDetails.push({
                    Points: this.props.selectedMemberDetails.Points,
                    ID: this.props.selectedMemberDetails.ID,
                    Title: this.props.selectedMemberDetails.Title,
                    FirstName: this.props.selectedMemberDetails.FirstName,
                    LastName: this.props.selectedMemberDetails.LastName,
                    Rank: this.props.selectedMemberDetails.Rank
                });
            }

            this.setState({
                selectedUserActivities: memberActivitesArray,
                selectedMemberDetails: memberDetails
            });
        }
    }

    //render the sort caret on the header column for accessbility issues fix
    customSortCaret = (order: any, column: any) => {

        if (!order) {
            return (
                <span className="sort-order">
                    <span className="dropdown-caret">
                    </span>
                    <span className="dropup-caret">
                    </span>
                </span>);
        }
        else if (order === 'asc') {
            if (column.dataField === stringsConstants.dateOfEventLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', "");
            } else if (column.dataField === stringsConstants.typeLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', "");
            } else if (column.dataField === stringsConstants.pointsLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', stringsConstants.sortAscAriaSort);
            }

            return (
                <span className="sort-order">
                    <span className="dropup-caret">
                    </span>
                </span>);
        }
        else if (order === 'desc') {
            if (column.dataField === stringsConstants.dateOfEventLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', "");
            } else if (column.dataField === stringsConstants.typeLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', "");
            } else if (column.dataField === stringsConstants.pointsLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', stringsConstants.sortDescAriaSort);
            }

            return (
                <span className="sort-order">
                    <span className="dropdown-caret">
                    </span>
                </span>);
        }
        return null;
    }

    //Main render method
    public render() {
        // To determine whether the component is called from sidebar or not
        const isSidebar = this.props.parentComponent === stringsConstants.SidebarLabel;
        const isChampionReport = this.props.parentComponent === stringsConstants.ChampionReportLabel;

        const activitiesTableHeader = [
            {
                dataField: stringsConstants.dateOfEventLabel,
                text: LocaleStrings.DateofEventGridLabel,
                headerTitle: true,
                title: true,
                sort: true,
                sortCaret: this.customSortCaret,
                headerAttrs: { "id": stringsConstants.dateOfEventId, "role": "columnheader", "scope": "col" },
                attrs: { 'role': 'cell', 'headers': stringsConstants.dateOfEventId, "tabindex": "0", "rowspan": "1", "colspan": "1" }
            },
            {
                dataField: stringsConstants.typeLabel,
                text: LocaleStrings.EventTypeGridLabel,
                sort: true,
                sortCaret: this.customSortCaret,
                headerAttrs: { "id": stringsConstants.eventTypeId, "role": "columnheader", "scope": "col" },
                headerTitle: true,
                title: true,
                attrs: { 'role': 'cell', 'headers': stringsConstants.eventTypeId, "tabindex": "0", "rowspan": "1", "colspan": "1" }
            },
            {
                dataField: stringsConstants.pointsLabel,
                text: LocaleStrings.CMPSideBarPointsLabel,
                headerTitle: true,
                title: true,
                sort: true,
                sortCaret: this.customSortCaret,
                headerAttrs: { "id": stringsConstants.eventPointsId, "role": "columnheader", "scope": "col" },
                attrs: { 'role': 'cell', 'headers': stringsConstants.eventPointsId, "tabindex": "0", "rowspan": "1", "colspan": "1" }
            }
        ];

        return (
            <React.Fragment>
                <div className="gtc-cards">
                    <div className="showActivitiesPopupBody">
                        <Row xl={isSidebar ? 1 : 2} lg={isSidebar ? 1 : 2} md={1} sm={1} xs={1} className="report-profile-grid-wrapper">
                            {this.props.parentComponent !== stringsConstants.SidebarLabel &&
                                <Col xl={isChampionReport ? 3 : 4} lg={isChampionReport ? 3 : 4} md={isChampionReport ? 5 : 12} sm={12} xs={12} className={isChampionReport ? "report-profile-wrapper" : ""}>
                                    {this.state.selectedMemberDetails.length > 0 &&
                                        <>
                                            {isChampionReport && <div className='events-profile-heading'>{LocaleStrings.ChampionLabel}</div>}
                                            <div className="showActivitiesImage-IconArea">
                                                <Person
                                                    personQuery={this.state.selectedMemberDetails[0].Title}
                                                    view={3}
                                                    personCardInteraction={1}
                                                    verticalLayout={true}
                                                    className="activities-profile-image"
                                                />
                                                {this.props.parentComponent === stringsConstants.ChampionsCardsLabel &&
                                                    <div className="showActivities-rank-points-block">
                                                        <span className="showActivities-rank" title={`Rank ${this.state.selectedMemberDetails[0].Rank}`}>Rank <span className="showActivities-rank-value"># {this.state.selectedMemberDetails[0].Rank}</span></span>
                                                        <span className="showActivities-points" title={`#${this.state.selectedMemberDetails[0].Points}`}>
                                                            {this.state.selectedMemberDetails[0].Points}
                                                            <Icon iconName="FavoriteStarFill" className="showActivities-points-star" />
                                                        </span>
                                                    </div>
                                                }
                                                {this.props.loggedinUserEmail !== this.state.selectedMemberDetails[0].Title &&
                                                    <div className="showActivities-icon-area">
                                                        <div className="request-to-call-link"
                                                            title={LocaleStrings.RequestToCallLabel}
                                                            onClick={() => this.openTask("https://teams.microsoft.com/l/meeting/new?subject=" +
                                                                currentUserName + " / " + this.state.selectedMemberDetails[0].FirstName + " " + this.state.selectedMemberDetails[0].LastName + " " + LocaleStrings.MeetupSubject +
                                                                "&content=" + LocaleStrings.MeetupBody + "&attendees=" + this.state.selectedMemberDetails[0].Title)}
                                                        > {LocaleStrings.RequestToCallLabel}
                                                        </div>
                                                    </div>
                                                }
                                            </div>
                                        </>
                                    }
                                </Col>
                            }
                            <Col xl={isSidebar ? 12 : isChampionReport ? 7 : 8} lg={isSidebar ? 12 : isChampionReport ? 7 : 8} md={isChampionReport ? 7 : 12}
                                sm={12} xs={12} className={isChampionReport ? "report-grid-wrapper" : ""}>
                                {isChampionReport && <div className="events-grid-heading">{LocaleStrings.ActivitiesLabel}</div>}
                                <div className="showActivities-grid-area">
                                    {!isChampionReport && <div className="activities-grid-heading">{LocaleStrings.ActivitiesLabel}</div>}
                                    <BootstrapTable
                                        bootstrap4
                                        keyField={'dateOfEvents'}
                                        data={this.state.filteredUserActivities}
                                        columns={activitiesTableHeader}
                                        table-responsive={true}
                                        wrapperClasses={isSidebar || isChampionReport ? "events-table-wrapper-class" : ""}
                                        noDataIndication={() => (<div className='activities-noRecordsFound'>{LocaleStrings.NoActivitiesinGridLabel}</div>)}
                                    />
                                    {this.state.filteredUserActivities.length > 0 &&
                                        <div className="pagination-area" dir='ltr'>
                                            <span>
                                                {this.state.pageNumber} of {Math.ceil(this.state.selectedUserActivities.length / this.state.userActivitiesPerPage)}
                                                <Icon
                                                    iconName="ChevronLeft"
                                                    className="Chevron-Icon"
                                                    onClick={this.state.pageNumber > 1 ? () => { this.setState({ pageNumber: this.state.pageNumber - 1 }); } : null}
                                                />

                                                <Icon
                                                    iconName="ChevronRight"
                                                    className="Chevron-Icon"
                                                    onClick={this.state.pageNumber < Math.ceil(this.state.selectedUserActivities.length / this.state.userActivitiesPerPage) ? () => { this.setState({ pageNumber: this.state.pageNumber + 1 }); } : null}
                                                />
                                            </span>
                                        </div>
                                    }
                                </div>
                            </Col>
                        </Row>
                    </div>
                </div>

            </React.Fragment>
        );
    }
}
