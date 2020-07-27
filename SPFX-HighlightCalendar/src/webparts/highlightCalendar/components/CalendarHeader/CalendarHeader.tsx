import * as React from 'react';

import { Stack, Label, Dropdown, IDropdownOption, IStackItemStyles, IconButton, DatePicker, PrimaryButton } from 'office-ui-fabric-react';

import styles from './CalendarHeader.module.scss';
import Calendar from '@toast-ui/react-calendar';
import TuiCalendar from 'tui-calendar';

const allMonths: string[] = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec'
];

interface ICalendarHeaderProps {
    calendarRef: React.RefObject<Calendar>;
}

interface ICalendarHeaderState {
    monthName: string;
}

export default class CalendarHeader extends React.Component<ICalendarHeaderProps, ICalendarHeaderState> {

    public constructor(props) {
        super(props);

        const _monthName: string = `${allMonths[(new Date()).getMonth()]} ${(new Date()).getFullYear()}`;

        this.state = {
            monthName: _monthName
        };
    }

    public render = (): JSX.Element => {
        const _monthName: string = this.state.monthName;

        return (
            <>
                <div className={styles.lineHeader}>
                </div>
                <Stack horizontal wrap horizontalAlign="center" className={styles.stackParent}>
                    <Stack.Item className={`${styles.stackLeft} ${styles.stackCommon}`}>
                        <IconButton className={styles.themeColor} iconProps={{ iconName: 'ChevronLeftSmall' }} title="Emoji" ariaLabel="Emoji" onClick={() => this.moveMonth(false)} />
                        <IconButton className={styles.themeColor} iconProps={{ iconName: 'ChevronRightSmall' }} title="Emoji" ariaLabel="Emoji" onClick={() => this.moveMonth(true)} />
                        <PrimaryButton text="Today" onClick={this.setToday} />
                    </Stack.Item>
                    <Stack.Item className={`${styles.stackCenter} ${styles.stackCommon}`}>
                        <Label className={styles.monthLabel}>{_monthName}</Label>
                    </Stack.Item>
                    <Stack.Item className={`${styles.stackRight} ${styles.stackCommon}`}>
                        <DatePicker
                            placeholder="Select a date..."
                            ariaLabel="Select a date"
                            highlightSelectedMonth={true}
                            onSelectDate={this.onSelectDate}
                        />
                    </Stack.Item>
                </Stack>
            </>
        );
    }

    private onSelectDate = (date: Date | null | undefined): void => {
        if (date) {
            const calendarInstance: TuiCalendar = this.getCalendarRefFromProp();

            calendarInstance.setDate(date);

            this.updateMonthLabel(calendarInstance.getDate().toDate());
        }
    }

    private setToday = (): void => {
        const calendarInstance: TuiCalendar = this.getCalendarRefFromProp();

        calendarInstance.today();

        this.updateMonthLabel(calendarInstance.getDate().toDate());
    }

    private moveMonth = (isForward: boolean) => {
        const calendarInstance: TuiCalendar = this.getCalendarRefFromProp();

        if (isForward) {
            calendarInstance.next();
        }
        else {
            calendarInstance.prev();
        }

        this.updateMonthLabel(calendarInstance.getDate().toDate());
    }

    private updateMonthLabel = (calendarInstanceDate: Date): void => {
        this.setState({
            monthName: `${allMonths[calendarInstanceDate.getMonth()]} ${calendarInstanceDate.getFullYear()}`
        });
    }

    private getCalendarRefFromProp = (): TuiCalendar => this.props.calendarRef.current.getInstance();
}