/**
 * @file PropertyFieldDateTimePickerHost.tsx
 * Renders the controls for PropertyFieldDateTimePicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldDateTimePickerPropsInternal, ITimeConvention } from './PropertyFieldDateTimePicker';
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import * as strings from 'sp-client-custom-fields/strings';

/**
 * @interface
 * PropertyFieldDateTimePickerHost properties interface
 *
 */
export interface IPropertyFieldDateTimePickerHostProps extends IPropertyFieldDateTimePickerPropsInternal {
}

/**
 * @class
 * Defines the labels of the DatePicker control (as months, days, etc.)
 *
 */
class DatePickerStrings implements IDatePickerStrings {
    /**
     * An array of strings for the full names of months.
     * The array is 0-based, so months[0] should be the full name of January.
     */
    public months: string[] = [
      strings.DatePickerMonthLongJanuary, strings.DatePickerMonthLongFebruary,
      strings.DatePickerMonthLongMarch, strings.DatePickerMonthLongApril,
      strings.DatePickerMonthLongMay, strings.DatePickerMonthLongJune, strings.DatePickerMonthLongJuly,
      strings.DatePickerMonthLongAugust, strings.DatePickerMonthLongSeptember, strings.DatePickerMonthLongOctober,
      strings.DatePickerMonthLongNovember, strings.DatePickerMonthLongDecember
    ];
    /**
     * An array of strings for the short names of months.
     * The array is 0-based, so shortMonths[0] should be the short name of January.
     */
    public shortMonths: string[] = [
      strings.DatePickerMonthShortJanuary, strings.DatePickerMonthShortFebruary,
      strings.DatePickerMonthShortMarch, strings.DatePickerMonthShortApril,
      strings.DatePickerMonthShortMay, strings.DatePickerMonthShortJune, strings.DatePickerMonthShortJuly,
      strings.DatePickerMonthShortAugust, strings.DatePickerMonthShortSeptember, strings.DatePickerMonthShortOctober,
      strings.DatePickerMonthShortNovember, strings.DatePickerMonthShortDecember
    ];
    /**
     * An array of strings for the full names of days of the week.
     * The array is 0-based, so days[0] should be the full name of Sunday.
     */
    public days: string[] = [
      strings.DatePickerDayLongSunday, strings.DatePickerDayLongMonday, strings.DatePickerDayLongTuesday,
      strings.DatePickerDayLongWednesday, strings.DatePickerDayLongThursday, strings.DatePickerDayLongFriday,
      strings.DatePickerDayLongSaturday
    ];
    /**
     * An array of strings for the initials of the days of the week.
     * The array is 0-based, so days[0] should be the initial of Sunday.
     */
    public shortDays: string[] = [
      strings.DatePickerDayShortSunday, strings.DatePickerDayShortMonday, strings.DatePickerDayShortTuesday,
      strings.DatePickerDayShortWednesday, strings.DatePickerDayShortThursday, strings.DatePickerDayShortFriday,
      strings.DatePickerDayShortSaturday
    ];
    /**
     * String to render for button to direct the user to today's date.
     */
    public goToToday: string = "";
    /**
     * Error message to render for TextField if isRequired validation fails.
     */
    public isRequiredErrorMessage: string = "";
    /**
     * Error message to render for TextField if input date string parsing fails.
     */
    public invalidInputErrorMessage: string = "";
}

export interface IPropertyFieldDateTimePickerHostPropsState {
  day?: Date;
  hours?: number;
  minutes?: number;
  seconds?: number;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldDateTimePicker component
 */
export default class PropertyFieldDateTimePickerHost extends React.Component<IPropertyFieldDateTimePickerHostProps, IPropertyFieldDateTimePickerHostPropsState> {

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldDateTimePickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onSelectDate = this.onSelectDate.bind(this);
    this.dropdownHoursChanged = this.dropdownHoursChanged.bind(this);
    this.dropdownMinutesChanged = this.dropdownMinutesChanged.bind(this);
    this.dropdownSecondsChanged = this.dropdownSecondsChanged.bind(this);

    this.state = {
      day: (this.props.initialDate != null && this.props.initialDate != '') ? new Date(this.props.initialDate) : null,
      hours: (this.props.initialDate != null && this.props.initialDate != '') ? new Date(this.props.initialDate).getHours() : 0,
      minutes: (this.props.initialDate != null && this.props.initialDate != '') ? new Date(this.props.initialDate).getMinutes() : 0,
      seconds: (this.props.initialDate != null && this.props.initialDate != '') ? new Date(this.props.initialDate).getSeconds() : 0,
      errorMessage: ''
    };

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  /**
   * @function
   * Function called when the DatePicker Office UI Fabric component selected date changed
   */
  private onSelectDate(date: Date): void {
    if (date == null)
      return;
    this.state.day = date;
    this.setState(this.state);
    this.saveDate();
  }

  private dropdownHoursChanged(element?: IDropdownOption): void {
    this.state.hours = Number(element.key);
    this.setState(this.state);
    this.saveDate();
  }

  private dropdownMinutesChanged(element?: any): void {
    this.state.minutes = Number(element.key);
    this.setState(this.state);
    this.saveDate();
  }

  private dropdownSecondsChanged(element?: any): void {
    this.state.seconds = Number(element.key);
    this.setState(this.state);
    this.saveDate();
  }

  private saveDate(): void {
    if (this.state.day == null)
      return;
    var finalDate = new Date(this.state.day.toISOString());
    finalDate.setHours(this.state.hours);
    finalDate.setMinutes(this.state.minutes);
    finalDate.setSeconds(this.state.seconds);

    if (finalDate != null) {
      var finalDateAsString: string = '';
      if (this.props.formatDate) {
        finalDateAsString = this.props.formatDate(finalDate);
      }
      else {
        finalDateAsString = finalDate.toString();
      }
      this.delayedValidate(finalDateAsString);
    }
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialDate, value);
      return;
    }

    if (this.latestValidateValue === value)
      return;
    this.latestValidateValue = value;

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialDate, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialDate, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.initialDate, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string, newValue: string) {
    if (this.props.onPropertyChange && newValue != null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
      if (!this.props.disableReactivePropertyChanges && this.props.render != null)
        this.props.render();
    }
  }

  /**
   * @function
   * Called when the component will unmount
   */
  public componentWillUnmount() {
    this.async.dispose();
  }

  /**
   * @function
   * Renders the control
   */
  public render(): JSX.Element {
    //Defines the DatePicker control labels
    var dateStrings: DatePickerStrings = new DatePickerStrings();
    //Constructs a Date type object from the initalDate string property
    var hours: IDropdownOption[] = [];
    for (var i = 0; i < 24; i++) {
      var digit: string;
      if (this.props.timeConvention == ITimeConvention.Hours24) {
        //24 hours time convention
        if (i < 10)
          digit = '0' + i;
        else
          digit = i.toString();
      }
      else {
        //12 hours time convention
        if (i == 0)
          digit = '12 am';
        else if (i < 12) {
          digit = i + ' am';
        }
        else {
          if (i == 12)
            digit = '12 pm';
          else {
            digit = (i % 12) + ' pm';
          }
        }
      }
      var selected: boolean = false;
      if (i == this.state.hours)
        selected = true;
      hours.push({ key: i, text: digit, isSelected: selected});
    }
    var minutes: IDropdownOption[] = [];
    for (var j = 0; j < 60; j++) {
      var digitMin: string;
      if (j < 10)
        digitMin = '0' + j;
      else
        digitMin = j.toString();
      var selected2: boolean = false;
      if (j == this.state.minutes)
        selected2 = true;
      minutes.push({ key: j, text: digitMin, isSelected: selected2});
    }
    var seconds: IDropdownOption[] = [];
    for (var k = 0; k < 60; k++) {
      var digitSec: string;
      if (k < 10)
        digitSec = '0' + k;
      else
        digitSec = k.toString();
      var selected3: boolean = false;
      if (k == this.state.seconds)
        selected3 = true;
      seconds.push({ key: k, text: digitSec, isSelected: selected3});
    }
    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <table cellPadding="0" cellSpacing="0" width="100%" style={{marginTop: '10px'}}>
          <tbody>
            <tr>
              <td style={{verticalAlign: 'top'}}><Label style={{marginRight: '4px'}}>{strings.DateTimePickerDate}</Label></td>
              <td style={{verticalAlign: 'top'}}>
                <DatePicker value={this.state.day} strings={dateStrings}
                  isMonthPickerVisible={false} onSelectDate={this.onSelectDate} allowTextInput={false}
                />
              </td>
            </tr>
            <tr>
              <td style={{verticalAlign: 'top'}}><Label style={{marginRight: '4px'}}>{strings.DateTimePickerTime}</Label></td>
              <td style={{verticalAlign: 'top'}}>
                <table cellPadding="0" cellSpacing="0">
                  <tbody>
                    <tr>
                      <td width="79">
                        <Dropdown
                          label=""
                          options={hours} onChanged={this.dropdownHoursChanged}
                          />
                      </td>
                      <td width="4" style={{paddingLeft: '2px', paddingRight: '2px'}}><Label>:</Label></td>
                      <td width="71">
                        <Dropdown
                          label=""
                          options={minutes} onChanged={this.dropdownMinutesChanged} />
                      </td>
                      <td width="4" style={{paddingLeft: '2px', paddingRight: '2px'}}><Label>:</Label></td>
                      <td width="71">
                        <Dropdown
                          label=""
                          options={seconds} onChanged={this.dropdownSecondsChanged} />
                      </td>
                    </tr>
                  </tbody>
                </table>
              </td>
            </tr>
          </tbody>
        </table>
        { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div style={{paddingBottom: '8px'}}><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
            : ''}
      </div>
    );
  }
}