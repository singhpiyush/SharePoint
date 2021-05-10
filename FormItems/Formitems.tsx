import { Callout, Checkbox, ChoiceGroup, DelayedRender, FontWeights, IChoiceGroupOption, IPersonaProps, ITooltipHostStyles, mergeStyleSets, SpinButton, Stack, TooltipHost } from 'office-ui-fabric-react';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { ComboBox, IComboBox, IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import * as React from 'react';
import Constants from '../../../../../utils/Constants';
import Utils from '../../../../../utils/Utils';
import { ControlProps, ControlType } from '../../../store/IStore';
import { UserRoleContext } from '../../Pod';
import { useId, useBoolean } from '@uifabric/react-hooks';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import styles from './FormItems.module.scss';

const __util = new Utils();

const helpLogo: any = require('../../../assets/img/Help.png');

export interface IFormItemsProps {
    formStyles?: React.CSSProperties;
    controlType: ControlType;
    controlProps: Partial<ControlProps>;
    controlKey: string;
    hidden?: boolean;
    disabled?: boolean;
    //_ref?: React.MutableRefObject<any>;
    callback: (key: string, value: string | number | Date | IDropdownOption | IComboBoxOption | IPersonaProps[] | IFilePickerResult) => void;
}

const FormItems: React.FunctionComponent<Partial<IFormItemsProps>> = (props) => {

    const [lookup, setLookup] = React.useState<IComboBoxOption[]>([]);

    const [isCalloutVisible, { toggle: toggleIsCalloutVisible }] = useBoolean(false);

    const uniqueID: string = useId();

    const _podServices = React.useContext(UserRoleContext);

    const formatDateValue = (): Date => {
        let _date: Date = null;
        if (props.controlProps.defaultValue) {
            _date = new Date(props.controlProps.defaultValue);
            if (isNaN(_date.getDate())) {
                _date = new Date(+props.controlProps.defaultValue);
            }
        }

        return _date;
    }

    const insertBtn = (): JSX.Element => <PrimaryButton text={props.controlProps.textLabel} style={props.formStyles} onClick={(event) => props.callback(props.controlKey, "")} disabled={props.disabled} />;

    const insertScndBtn = (): JSX.Element => <DefaultButton text={props.controlProps.textLabel} style={props.formStyles} disabled={props.disabled} />;

    const insertSpinBtn = (): JSX.Element => <SpinButton disabled={props.disabled} incrementButtonAriaLabel={'Increase value by 1'} decrementButtonAriaLabel={'Decrease value by 1'} step={1} label={props.controlProps.textLabel} />;

    const insertChoiceGroup = (): JSX.Element => <ChoiceGroup key={props.controlKey} label={props.controlProps.textLabel} defaultSelectedKey={props.controlProps.choiceGrp.selKey} disabled={props.disabled} options={props.controlProps.choiceGrp.options} style={props.formStyles} onChange={(ev?: React.FormEvent<HTMLInputElement | HTMLElement>, option?: IChoiceGroupOption) => props.callback(props.controlKey, option.key)} />;

    const insertLabel = (labelText: string, _className: string): JSX.Element => <Label className={_className} style={props.formStyles}>{labelText}</Label>;

    const insertText = (): JSX.Element => <TextField
        key={props.controlKey}
        multiline={props.controlProps.isMultiselect}
        defaultValue={props.controlProps.defaultValue}
        className={`${props.controlProps.className}`}
        //value={props.controlProps.value} 
        disabled={props.disabled}
        onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => props.callback(props.controlKey, newValue)} />;

    const insertDate = (): JSX.Element => <DatePicker
        key={props.controlKey}
        placeholder={props.controlProps.placeholder}
        //label={props.controlProps.textLabel}
        isRequired={props.controlProps.isRequired}
        //initialPickerDate={props.controlProps.defaultValue ? new Date(props.controlProps.defaultValue) : null}
        //value={props.controlProps.defaultValue && new Date(+props.controlProps.defaultValue)}
        value={formatDateValue()}
        className={props.controlProps.className}
        disabled={props.disabled}
        //componentRef={props.controlProps._ref}
        minDate={props.controlProps.date && props.controlProps.date.minDate}
        formatDate={(_date?: Date) => __util.getDateformatDDMMYYY(_date)}
        onSelectDate={(date: Date | null | undefined) => props.callback(props.controlKey, date.toISOString())} />;

    const insertDropDown = (): JSX.Element => <Dropdown
        key={props.controlKey}
        placeholder={props.controlProps.placeholder}
        label={props.controlProps.textLabel}
        ariaLabel={props.controlProps.textLabel}
        multiSelect={props.controlProps.isMultiselect}
        required={props.controlProps.isRequired}
        disabled={props.disabled}
        className={`${props.controlProps.className}`}
        defaultSelectedKey={props.controlProps.dropdown.selValue}
        defaultSelectedKeys={props.controlProps.dropdown.selValue as string[]}
        options={props.controlProps.dropdown.ddOptions}
        //componentRef={props.controlProps._ref}
        onChange={(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => props.callback(props.controlKey, option.key)} />;

    const insertLookup = (): JSX.Element => {

        let _localOptions: IComboBoxOption[] = null;

        if (lookup.length === 0) {
            _podServices.service.getLookupList(props.controlProps.lookup.listID, props.controlProps.lookup.selectedFields, props.controlProps.lookup.activeFiled)
                .then((onFulfilled: any[]) => {
                    //debugger;
                    console.log('Lookup Fetch');

                    let selectedVal: IComboBoxOption = null;

                    let _dropDownOpt: IComboBoxOption[] = [];
                    onFulfilled.forEach((option: any) => {
                        //debugger;
                        _dropDownOpt.push({
                            //key: option['ID'],
                            key: option[props.controlProps.lookup.keyField],
                            text: option[props.controlProps.lookup.displayField],
                            title: option[props.controlProps.lookup.titleField]
                        });

                        //if (option['ID'] === props.controlProps.lookup.selValue) {
                        if (option[props.controlProps.lookup.displayField] === props.controlProps.lookup.selValue) {
                            selectedVal = _dropDownOpt[_dropDownOpt.length - 1];
                        }
                    });

                    if (selectedVal) {
                        //set selection for existing items
                        props.callback(props.controlKey, selectedVal);
                    }

                    setLookup([..._dropDownOpt]);
                });
        }

        else {
            _localOptions = [...lookup];
            _localOptions.forEach((option: IComboBoxOption) => {
                option.selected = option.key == props.controlProps.lookup.selValue
            });
        }

        props.controlProps.combo = {
            ccOptions: _localOptions,
            defSelKey: props.controlProps.lookup.defSelValue as string,
            selKey: props.controlProps.lookup.selValue as string
        };

        /*return <ComboBox
            placeholder={props.controlProps.placeholder}
            allowFreeform
            //label={props.controlProps.textLabel}
            ariaLabel={props.controlProps.textLabel}
            multiSelect={props.controlProps.isMultiselect}
            required={props.controlProps.isRequired}
            disabled={props.disabled}
            className={`${props.controlProps.className}`}
            defaultSelectedKey={props.controlProps.lookup.selValue}
            //defaultSelectedKeys={props.controlProps.lookup.selValue as string[]}
            options={[...lookup]}
            //componentRef={props.controlProps._ref}
            onChange={(event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => props.callback(props.controlKey, option)} />;
            */

        return insertCombo();
    };

    const insertCombo = (): JSX.Element => <ComboBox
        key={props.controlKey}
        placeholder={props.controlProps.placeholder}
        //label={props.controlProps.textLabel}
        ariaLabel={props.controlProps.textLabel}
        multiSelect={props.controlProps.isMultiselect}
        required={props.controlProps.isRequired}
        disabled={props.disabled}
        className={`${props.controlProps.className}`}
        defaultSelectedKey={props.controlProps.combo.defSelKey}
        selectedKey={props.controlProps.combo.selKey}
        options={props.controlProps.combo.ccOptions || [...lookup]}
        //componentRef={props.controlProps._ref}
        onChange={(event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => props.callback(props.controlKey, option)} />;

    const insertCheckBox = (): JSX.Element => <Checkbox
        key={props.controlKey}
		disabled={props.disabled}
        defaultChecked={props.controlProps.defaultValue != null && props.controlProps.defaultValue != undefined && props.controlProps.defaultValue === props.controlProps.checkBox.trueValue}
        onChange={(ev: React.FormEvent<HTMLElement>, isChecked: boolean) => props.callback(props.controlKey, isChecked.toString())} />;

    const insertFileUpload = (): JSX.Element => <FilePicker
        key={props.controlKey}
        //onChange={(filePickerResult: IFilePickerResult) => props.callback(props.controlKey, filePickerResult)}
        onSave={(filePickerResult: IFilePickerResult) => props.callback(props.controlKey, filePickerResult)}
        buttonIcon={props.controlProps.fileUpload.buttonIcon}
        buttonIconProps={props.controlProps.fileUpload.buttonIconProps}
        buttonLabel={props.controlProps.fileUpload.buttonLabel}
        hideLinkUploadTab
        hideOneDriveTab
        hideOrganisationalAssetTab
        hideRecentTab
        hideSiteFilesTab
        hideStockImages
        hideWebSearchTab
        disabled={props.disabled}
        context={_podServices.service.getContext()} />;

    const insertPeoplePicker = (): JSX.Element => <PeoplePicker
        key={props.controlKey}
        context={_podServices.service.getContext()}
        personSelectionLimit={1}
        showtooltip={true}
        required={props.controlProps.isRequired}
        disabled={props.disabled}
        principalTypes={[PrincipalType.User]}
        resolveDelay={1000}
        defaultSelectedUsers={props.controlProps.people.defaultSelectedUsers}
        onChange={(items: IPersonaProps[]) => props.callback(props.controlKey, items)} />

    const _render = (): JSX.Element => {
        switch (props.controlType) {
            case ControlType.Text:
                return insertText();
            case ControlType.Date:
                //debugger;    
                return insertDate();
            case ControlType.DD:
                //debugger;
                return insertDropDown();
            case ControlType.Combo:
                return insertCombo();
            case ControlType.Label:
                return insertLabel(props.controlProps.textLabel, props.controlProps.className);
            case ControlType.Button:
                return insertBtn();
            case ControlType.ScndBtn:
                return insertScndBtn();
            case ControlType.ChoiceGrp:
                return insertChoiceGroup();
            case ControlType.SpinBtn:
                return insertSpinBtn();
            case ControlType.Lookup:
                return insertLookup();
            case ControlType.CheckBox:
                return insertCheckBox();
            case ControlType.FileUpload:
                return insertFileUpload();
            case ControlType.People:
                return insertPeoplePicker();
            default:
                return null;
        }
    };

    const stylesCallout = mergeStyleSets({
        callout: {
            maxWidth: 300,
        },
        subtext: [
            {
                margin: 0,
                height: '100%',
                padding: '24px 20px',
                fontWeight: FontWeights.semilight,
            },
        ],
    });

    //const tooltipId = useId('tooltip');

    return (
        <div className={`${styles.formItems}`} key={props.controlKey}>
            <div className={`${props.controlProps.errorMsg ? styles.errBox : ''}`}>
                {
                    //Exclude for label controls
                    props.controlType !== ControlType.Label && <Stack horizontal >
                        {insertLabel(props.controlProps.textLabel, props.controlProps.isRequired ? styles.required : '')}
                        {props.controlProps.tooltip &&
                            <img src={helpLogo} onClick={toggleIsCalloutVisible} className={uniqueID} onMouseOver={toggleIsCalloutVisible} />
                        }
                    </Stack>
                }
                {_render()}
                {props.controlProps.errorMsg && insertLabel(props.controlProps.errorMsg, styles.errLabel)}
                {isCalloutVisible && (
                    <Callout
                        className={stylesCallout.callout}
                        target={`.${uniqueID}`}
                        onDismiss={toggleIsCalloutVisible}
                        role="status"
                        aria-live="assertive"
                    >
                        <DelayedRender>
                            {/* <p className={stylesCallout.subtext}>
                                {props.controlProps.tooltip}
                            </p> */}
                            <div className={stylesCallout.subtext} dangerouslySetInnerHTML={{ __html: props.controlProps.tooltip }} />
                        </DelayedRender>
                    </Callout>
                )}
            </div>
        </div>
    );
};

// const tooltipStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
// const calloutProps = { gapSpace: 0 };

export default FormItems;