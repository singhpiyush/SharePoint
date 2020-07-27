import * as React from 'react';

import * as moment from 'moment';
//import { Stack, Label, Dropdown, IDropdownOption, IStackItemStyles, IconButton, DatePicker } from 'office-ui-fabric-react';

import Calendar from '@toast-ui/react-calendar';
import { ITemplateConfig } from 'tui-calendar';

import 'tui-calendar/dist/tui-calendar.css';
import './HighlightCalendar.css';
import styles from './HighlightCalendar.module.scss';

import CalendarHeader from '../CalendarHeader/CalendarHeader';

const templates: ITemplateConfig = {
    popupIsAllDay: () => {
        return 'All Day';
    },
    popupStateFree: () => {
        return 'Free';
    },
    popupStateBusy: () => {
        return 'Busy';
    },
    titlePlaceholder: () => {
        return 'Subject';
    },
    locationPlaceholder: () => {
        return 'Location';
    },
    startDatePlaceholder: () => {
        return 'Start date';
    },
    endDatePlaceholder: () => {
        return 'End date';
    },
    popupSave: () => {
        return 'Save';
    },
    popupUpdate: () => {
        return 'Update';
    },
    popupDetailDate: (isAllDay, start, end) => {
        /*
        //const isSameDate = moment(start).isSame(end);
        //const endFormat = (isSameDate ? '' : 'YYYY.MM.DD ') + 'hh:mm a';
        const endFormat = 'YYYY.MM.DD ' + 'hh:mm a';

        if (isAllDay) {
            //return moment(start).format('YYYY.MM.DD') + (isSameDate ? '' : ' - ' + moment(end).format('YYYY.MM.DD'));
            return moment(start).format('YYYY.MM.DD') + moment(end).format('YYYY.MM.DD');
        }

        //return (moment(start).format('YYYY.MM.DD hh:mm a') + ' - ' + moment(end).format(endFormat));
        return (moment(start).format('YYYY.MM.DD hh:mm a') + ' - ' + moment(end).format(endFormat));
        */
        return '';
    },
    popupDetailLocation: (schedule) => {
        //return 'Location : ' + schedule.location;
        return '';
    },
    popupDetailUser: (schedule) => {
        return 'User : ' + (schedule.attendees || []).join(', ');
    },
    popupDetailState: (schedule) => {
        return 'State : ' + schedule.state || 'Busy';
    },
    popupDetailRepeat: (schedule) => {
        return 'Repeat : ' + schedule.recurrenceRule;
    },
    popupDetailBody: (schedule) => {
        return 'Body : ' + schedule.body;
    },
    popupEdit: () => {
        return 'Edit';
    },
    popupDelete: () => {
        return 'Delete';
    }
};

const calendarOptions = {
    // sort of option properties.
    height: '450px',
    view: 'month',
    isReadonly: true,
    month: {
        isAlways6Week: false
    },
    template: templates,
    useDetailPopup: false,
    useCreationPopup: false,
    disableDblClick: false,
    disableClick: true,
    schedules: [
        {
            id: '1',
            calendarId: '0',
            title: 'TOAST UI Calendar Study',
            category: 'time',
            dueDateClass: styles.fullDayEffect,
            start: new Date(),
            end: new Date(),
            bgColor: 'red',
            isAllDay: true
        },
        {
            id: '2',
            calendarId: '0',
            title: 'Practice',
            category: 'milestone',
            dueDateClass: '',
            start: new Date((new Date()).setDate(new Date().getDate() + 5)),
            end: new Date((new Date()).setDate(new Date().getDate() + 7)),
            isReadOnly: true,
            bgColor: 'blue'
        },
        {
            id: '3',
            calendarId: '0',
            title: 'FE Workshop',
            category: 'allday',
            dueDateClass: '',
            start: new Date((new Date()).setDate(new Date().getDate() + 10)),
            end: new Date((new Date()).setDate(new Date().getDate() + 13)),
            isReadOnly: true,
            bgColor: 'yellow'
        },
        {
            id: '4',
            calendarId: '0',
            title: 'Report',
            category: 'time',
            dueDateClass: '',
            start: new Date((new Date()).setDate(new Date().getDate() + 20)),
            end: new Date((new Date()).setDate(new Date().getDate() + 25)),
            bgColor: 'red'
        }
    ]
};

// const dropdownOptions: IDropdownOption[] = [{
//     key: 'July 2020',
//     text: 'July 2020',
//     selected: true
// },
// {
//     key: 'August 2020',
//     text: 'August 2020',
// }];

// const stackItemStyles: IStackItemStyles = {
//     root: {
//         alignItems: 'center',
//         display: 'flex',
//         justifyContent: 'flex-end',
//         width: '25%'
//     },
// };

export default class HighlightCalendar extends React.PureComponent {

    private calendarRef: React.RefObject<Calendar>;

    constructor(props) {
        super(props);

        this.calendarRef = React.createRef();
    }

    public componentDidMount = (): void => {
        this.deleteSelectionBug();
    }

    public componentDidUpdate = (): void => {
        this.deleteSelectionBug();
    }

    public render = (): JSX.Element => {
        return (
            <>
                <CalendarHeader calendarRef={this.calendarRef} />
                <Calendar
                    style={{ border: '1px solid gray' }}
                    usageStatistics={false}
                    ref={this.calendarRef}
                    {...calendarOptions}
                    onBeforeCreateSchedule={(test) => console.log(`before ${test}`)}
                    //onClick={(test) => console.log(test)}
                    onDoubleClick={(test) => console.log(`db ${test}`)}
                />
            </>
        );
    }

    private deleteSelectionBug = (): void => {

        // Array.prototype.forEach.call(document.getElementsByClassName('tui-full-calendar-month-guide-block'), (el: HTMLElement) => {
        //     // Do stuff here
        //     el.remove();
        //     console.log('test')
        // });
    }
}