// IncidentManagement.tsx

import * as React from 'react';
import styles from './IncidentManagement.module.scss';
import type { IIncidentManagementProps } from './IIncidentManagementProps';
import { Stack, Label, TextField, DatePicker, DefaultButton, StackItem, PrimaryButton, TimePicker, IComboBox, IStackTokens } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Utility from '../utils';
const moment = require('moment');
const logo = require('../assets/logo.png');

const utility = new Utility();

const IncidentManagement: React.FC<IIncidentManagementProps> = (props: IIncidentManagementProps) => {
    const [IncidentformState, setIncidentsetFormState] = React.useState({
        ApplicationName: "",
        ApplicationOwnerId: [],
        ApplicationSME: "",
        AssignmentGroup: "",
        BusinessImpact: "",
        Comments: "",
        Conferencebridgedetails: {Url: '', Description: ''},
        IncidentNumber: "",
        IssueDescription: "",
        NumberofusersImpacted: "",
        TeamsEngaged: "",
        CreatedTime: "",
        MIDeclaredTime: "",
        ResolutionTime: ""
    });
    const [CreateTimeDate, setCreateTimeDate] = React.useState(IncidentformState.CreatedTime);
    const [MIDeclaredTimeDate, setMIDeclaredTimeDate] = React.useState(new Date());
    const [ResolutionTimeDate, setResolutionTimeDate] = React.useState(new Date());
    const [listData, setListData] = React.useState([] as any)
    const urlParams = new URLSearchParams(window.location.search);
    const [formMode, setFormMode] = React.useState(urlParams.get('FormMode'))
    const [selectedId, setSelectedId] = React.useState(urlParams.get('ItemId'));
    const [user, setUser] = React.useState([] as any);

    React.useEffect(() => {
        init()
        setFormMode(urlParams.get('FormMode'));
        setSelectedId(urlParams.get('ItemId'))
    }, []);

    const init = async (): Promise<void> => {
        const data = formMode === 'View' || formMode === 'Edit' ? await utility.getSelectedItem(parseInt(selectedId as string), props.lists.title) : "";
        setListData(data);
    }

    console.log(listData)

    React.useEffect(() => {
        setIncidentsetFormState({
            ApplicationName: listData.ApplicationName,
            ApplicationOwnerId: listData.Id,
            ApplicationSME: listData.ApplicationSME,
            AssignmentGroup: listData.AssignmentGroup,
            BusinessImpact: listData.BusinessImpact,
            Comments: listData.Comments,
            Conferencebridgedetails: {Url: listData.Conferencebridgedetails?.Url, Description:listData.Conferencebridgedetails?.Url} ,
            IncidentNumber: listData.IncidentNumber,
            IssueDescription: listData.IssueDescription,
            NumberofusersImpacted: listData.NumberofusersImpacted,
            TeamsEngaged: listData.TeamsEngaged,
            CreatedTime: listData.CreatedTime,
            MIDeclaredTime: listData.MIDeclaredTime,
            ResolutionTime: listData.ResolutionTime
        });
        formMode === 'View' || formMode === 'Edit' ? utility.getUsers(listData.ApplicationOwnerId).then(k => {
            setUser(k);
        }) : ""
    }, [listData]);

    const onTextFieldChange = (newValue: string, stateName: string) => {
        setIncidentsetFormState({
            ...IncidentformState,
            [stateName]: newValue
        });
    }
    const onHyperLinkTextFieldChange = (newValue: string, stateName: string) => {
        setIncidentsetFormState({
            ...IncidentformState,
            [stateName]: {
                Url: newValue,
                Description: newValue
            }
        });
    }
    const onPeoplePickerChange = (newValue: any, stateName: string) => {
        setIncidentsetFormState({
            ...IncidentformState,
            [stateName]: newValue[0].id
        });
    }
    const onCreatedTimeDateChange = (date: any, stateName: string) => {
        setIncidentsetFormState({
            ...IncidentformState,
            [stateName]: date
        });
        setCreateTimeDate(date);

    };
    const onCreatedTimeChange = React.useCallback((_ev: React.FormEvent<IComboBox>, Time: Date) => {
        if (CreateTimeDate) {
            const concatenatedDateTime = moment(CreateTimeDate).format('MM-DD-YYYY') + " " + moment(Time).format('HH:mm');
            setIncidentsetFormState({
                ...IncidentformState,
                'CreatedTime': concatenatedDateTime
            });
        }
    }, [CreateTimeDate]);
    const onMIDeclaredTimeDateChange = (date: any, stateName: string) => {
        setMIDeclaredTimeDate(date);
        setIncidentsetFormState({
            ...IncidentformState,
            [stateName]: date
        });

    };
    const onMIDeclaredTimeChange = React.useCallback((_ev: React.FormEvent<IComboBox>, Time: Date) => {
        if (MIDeclaredTimeDate) {
            const concatenatedDateTime = moment(MIDeclaredTimeDate).format('MM-DD-YYYY') + " " + moment(Time).format('HH:mm');
            setIncidentsetFormState({
                ...IncidentformState,
                'MIDeclaredTime': concatenatedDateTime
            });
        }
    }, [MIDeclaredTimeDate]);
    const onResolutionTimeDateChange = (date: any, stateName: string) => {
        setResolutionTimeDate(date);
        setIncidentsetFormState({
            ...IncidentformState,
            [stateName]: date
        });

    };
    const onResolutionTimeChange = React.useCallback((_ev: React.FormEvent<IComboBox>, Time: Date) => {
        if (ResolutionTimeDate) {
            const concatenatedDateTime = moment(ResolutionTimeDate).format('MM-DD-YYYY') + " " + moment(Time).format('HH:mm');
            setIncidentsetFormState({
                ...IncidentformState,
                'ResolutionTime': concatenatedDateTime
            });
        }
    }, [ResolutionTimeDate]);
    console.log(IncidentformState)
    const CreatedTime = new Date(moment(IncidentformState.CreatedTime, 'MM-DD-YYYY HH:mm').toDate()) 

    return (
        <>
            <div className='info'>
                <div className={styles.Header}><img alt="FGLogo" src={logo} /><span className={styles.HeaderText}>{props.webpartTitle}</span></div>
                <br />
            </div>
            <div>
                <Stack horizontal>
                    <StackItem>
                        <Label>Application Name</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.ApplicationName} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onTextFieldChange(newValue, 'ApplicationName') }} />

                    </StackItem>
                    <StackItem>
                        <Label>Incident Number</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.IncidentNumber} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onTextFieldChange(newValue, 'IncidentNumber') }} />
                    </StackItem>
                </Stack>
                <Stack horizontal>
                    <StackItem>
                        <Label>Application Owner</Label>
                        <PeoplePicker
                        disabled={formMode === 'View' ? true : false}
                            // titleText="RM"
                            context={props?.context}
                            ensureUser
                            personSelectionLimit={100}
                            principalTypes={[PrincipalType.User]}
                            required={true}
                            onChange={(items) => {
                                onPeoplePickerChange(items, 'ApplicationOwnerId')
                            }}
                            defaultSelectedUsers={user}
                        />
                    </StackItem>
                    <StackItem>
                        <Label>Assignment Group</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.AssignmentGroup} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onTextFieldChange(newValue, 'AssignmentGroup') }} />
                    </StackItem>
                </Stack>
                <Stack horizontal>
                    <StackItem>
                        <Label>Business Impact</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.BusinessImpact} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onTextFieldChange(newValue, 'BusinessImpact') }} />
                    </StackItem>
                    <StackItem>
                        <Label>Application SME</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.ApplicationSME} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onTextFieldChange(newValue, 'ApplicationSME') }} />
                    </StackItem>
                </Stack>
                <Stack horizontal>
                    <StackItem>
                        <Label>Number of users Impacted</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.NumberofusersImpacted} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onTextFieldChange(newValue, 'NumberofusersImpacted') }} />
                    </StackItem>
                    <StackItem>
                        <Label>Created Time</Label>
                        <Stack horizontal className='dateAndTime' style={{ margin: '0px' }}>
                            <StackItem>
                                <DatePicker
                                disabled={formMode === 'View' ? true : false}
                                    firstWeekOfYear={1}
                                    showMonthPickerAsOverlay={true}
                                    placeholder="Select a date..."
                                    ariaLabel="Select a date"
                                    onSelectDate={(date: Date) => { onCreatedTimeDateChange(date, 'CreatedTime') }}
                                    value={new Date(CreateTimeDate)}
                                // onSelectDate={onDateChange}
                                />
                            </StackItem>
                            <StackItem>
                                <TimePicker disabled={formMode === 'View' ? true : false} placeholder="Select a time" onChange={onCreatedTimeChange}></TimePicker>
                            </StackItem>
                        </Stack>
                    </StackItem>
                </Stack>
                <Stack horizontal>
                    <StackItem>
                        <Label>Teams Engaged</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.TeamsEngaged} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onTextFieldChange(newValue, 'TeamsEngaged') }} />
                    </StackItem>
                    <StackItem>
                        <Label>MI Declared Time</Label>
                        <Stack horizontal className='dateAndTime' style={{ margin: '0px' }}>
                            <StackItem>
                                <DatePicker
                                disabled={formMode === 'View' ? true : false}
                                    firstWeekOfYear={1}
                                    showMonthPickerAsOverlay={true}
                                    placeholder="Select a date..."
                                    ariaLabel="Select a date"
                                    onSelectDate={(date: Date) => { onMIDeclaredTimeDateChange(date, 'MIDeclaredTime') }}
                                />
                            </StackItem>
                            <StackItem>
                                <TimePicker disabled={formMode === 'View' ? true : false} placeholder="Select a time" onChange={onMIDeclaredTimeChange}></TimePicker>
                            </StackItem>
                        </Stack>
                    </StackItem>
                </Stack>
                <Stack horizontal >
                    <StackItem >
                        <Label>Conference bridge details</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.Conferencebridgedetails.Url} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onHyperLinkTextFieldChange(newValue, 'Conferencebridgedetails') }} />
                    </StackItem>
                    <StackItem >
                        <Label>Resolution Time</Label>
                        <Stack horizontal className='dateAndTime' style={{ margin: '0px' }} >
                            <StackItem >
                                <DatePicker
                                disabled={formMode === 'View' ? true : false}
                                    firstWeekOfYear={1}
                                    showMonthPickerAsOverlay={true}
                                    placeholder="Select a date..."
                                    ariaLabel="Select a date"
                                    onSelectDate={(date: Date) => { onResolutionTimeDateChange(date, 'ResolutionTime') }}
                                />
                            </StackItem>
                            <StackItem className={styles.ResolutionTimePicker}>
                                <TimePicker disabled={formMode === 'View' ? true : false} placeholder="Select a time" onChange={onResolutionTimeChange}></TimePicker>
                            </StackItem>
                        </Stack>
                    </StackItem>
                </Stack>
                <Stack horizontal>
                    <StackItem>
                        <Label>Issue Description</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.IssueDescription} multiline rows={3} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onTextFieldChange(newValue, 'IssueDescription') }} />
                    </StackItem>
                </Stack>
                <Stack horizontal>
                    <StackItem>
                        <Label>Comments</Label>
                        <TextField disabled={formMode === 'View' ? true : false} value={IncidentformState.Comments} multiline rows={3} onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string) => { onTextFieldChange(newValue, 'Comments') }} />
                    </StackItem>
                </Stack>
                <Stack horizontal>
                    <StackItem>
                        {formMode === "Edit" ? "" : ""}
                        <PrimaryButton disabled={formMode === 'View' ? true : false} onClick={async () => {
                            const info = formMode === 'Edit' ? await utility.updateItem(selectedId, props.lists.title, IncidentformState) :
                                await utility.saveItem(IncidentformState, props.lists.title);
                            // await utility.saveItem(IncidentformState, props.lists.title); 
                        }}>Submit</PrimaryButton>
                        <DefaultButton className={styles.cancelButton}>Cancel</DefaultButton>
                    </StackItem>
                </Stack>
            </div>
        </>
    )
}
export default IncidentManagement;
