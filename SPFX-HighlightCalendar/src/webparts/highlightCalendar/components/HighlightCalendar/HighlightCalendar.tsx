import * as React from 'react';

import Calendar from '@toast-ui/react-calendar';
import 'tui-calendar/dist/tui-calendar.css';
import './HighlightCalendar.css';

const calendarOptions = {
    // sort of option properties.
    height: '450px',
    view: 'month',
    isReadonly: true,
    month: {
        isAlways6Week: false
    }
};

export default class HighlightCalendar extends React.PureComponent {

    private calendarRef: React.RefObject<Calendar>;

    constructor(props) {
        super(props);

        this.calendarRef = React.createRef();
    }

    public render = (): JSX.Element => {
        return (
            <Calendar
                style={{ border: '1px solid gray' }}
                usageStatistics={false}
                ref={this.calendarRef}
                {...calendarOptions}
            />
        );
    }
}