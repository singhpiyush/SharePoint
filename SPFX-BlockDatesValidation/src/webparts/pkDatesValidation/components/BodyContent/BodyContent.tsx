import React from 'react';

import { PrimaryButton, Panel, DatePicker, ICalendarProps, MessageBar, MessageBarType } from 'office-ui-fabric-react';

interface IBodyContent {
    showPanel: boolean;
    showMessagePanel: boolean;
    messageText: string;
}

export default class BodyContent extends React.PureComponent<{}, IBodyContent> {

    private calendarProps: ICalendarProps;

    constructor(props: Readonly<{}>) {
        super(props);

        this.state = {
            showPanel: false,
            showMessagePanel: false,
            messageText: ''
        };
    }

    public render(): JSX.Element {
        this.calendarProps = {
            showGoToToday: true,
            strings: null,
            restrictedDates: this._restrictedDates()
        };
        const _showPanel = this.state.showPanel,
            _showMessagePanel = this.state.showMessagePanel,
            _messageText = this.state.messageText;

        return (
            <div>
                This is an example to demonstrate how to block certain dates from being picked in the SPFx DatePicker control.<br />
                For this example, I'll be blocking 2 dates, yesterday and tomorrow. To see it in action, launch the Panel using the below button.
                <br /><br />
                <PrimaryButton>Launch the date picker panel</PrimaryButton>
                <Panel
                    isOpen={_showPanel}
                    onDismiss={this._hidePanel}
                    headerText='DatePicker Validation'
                >
                    {_showMessagePanel ? <MessageBar messageBarType={MessageBarType.success}>{_messageText}</MessageBar> : null}
                    <DatePicker calendarProps={this.calendarProps} label="Date Validtion" isRequired={true} onSelectDate={this._dateChange} />
                </Panel>
            </div>
        );
    }

    private _dateChange = (date: Date | null | undefined): void => {
        this.setState({
            showMessagePanel: true,
            messageText: `Date selected: ${date.toDateString()}`
        });
    }

    private _restrictedDates = (): Date[] => {
        const dateAry: Date[] = [],
            currentDate: Date = new Date(Date.now());

        const tomorrowDate: Date = new Date(new Date().setDate(currentDate.getDate() + 1)),
            previousDate: Date = new Date(new Date().setDate(currentDate.getDate() - 1));
        dateAry.push(previousDate);
        dateAry.push(tomorrowDate);

        return dateAry;
    }

    private _hidePanel = () => {
        this.setState({
            showPanel: false
        });
    }
}