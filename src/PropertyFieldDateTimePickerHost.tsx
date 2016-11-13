/**
 * @file PropertyFieldDateTimePickerHost.tsx
 * Renders the controls for PropertyFieldDateTimePicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldDateTimePickerPropsInternal } from './PropertyFieldDateTimePicker';
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
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
}

/**
 * @class
 * Renders the controls for PropertyFieldDateTimePicker component
 */
export default class PropertyFieldDateTimePickerHost extends React.Component<IPropertyFieldDateTimePickerHostProps, IPropertyFieldDateTimePickerHostPropsState> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldDateTimePickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onSelectDate = this.onSelectDate.bind(this);
    this.dropdownHoursChanged = this.dropdownHoursChanged.bind(this);
    this.dropdownMinutesChanged = this.dropdownMinutesChanged.bind(this);

    this.state = {
      day: (this.props.initialDate != null && this.props.initialDate != '') ? new Date(this.props.initialDate) : null,
      hours: (this.props.initialDate != null && this.props.initialDate != '') ? new Date(this.props.initialDate).getHours() : 0,
      minutes: (this.props.initialDate != null && this.props.initialDate != '') ? new Date(this.props.initialDate).getMinutes() : 0,
    };
    this.setState(this.state);
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

  private saveDate(): void {
    if (this.state.day == null)
      return;
    var finalDate = new Date(this.state.day.toISOString());
    finalDate.setHours(this.state.hours);
    finalDate.setMinutes(this.state.minutes);

    if (this.props.onPropertyChange && finalDate != null) {
      //Checks if a formatDate function has been defined
      if (this.props.formatDate) {
        this.props.properties[this.props.targetProperty] = this.props.formatDate(finalDate);
        this.props.onPropertyChange(this.props.targetProperty, this.props.initialDate, this.props.formatDate(finalDate));
      }
      else {
        this.props.properties[this.props.targetProperty] = finalDate.toString();
        this.props.onPropertyChange(this.props.targetProperty, this.props.initialDate, finalDate.toString());
      }
    }
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {
    //Defines the DatePicker control labels
    var dateStrings: DatePickerStrings = new DatePickerStrings();
    //Constructs a Date type object from the initalDate string property
    var hours: IDropdownOption[] = [];
    for (var i = 0; i < 24; i++) {
      var digit: string;
      if (i < 10)
        digit = '0' + i;
      else
        digit = i.toString();
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
    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <div style={{display: 'inline-flex'}}>
          <div style={{width:'180px', paddingTop: '10px', marginRight:'2px'}}>
              <DatePicker value={this.state.day} strings={dateStrings}
                isMonthPickerVisible={false} onSelectDate={this.onSelectDate} allowTextInput={false}
                />
          </div>
          <div style={{display: 'inline-flex', marginBottom: '8px'}}>
            <div style={{width:'47px'}}>
              <Dropdown
                label=""
                options={hours} onChanged={this.dropdownHoursChanged}
                />
            </div>
            <div style={{paddingTop: '16px', paddingLeft: '2px', paddingRight: '2px'}}>:</div>
            <div style={{width:'47px'}}>
                <Dropdown
                label=""
                options={minutes} onChanged={this.dropdownMinutesChanged} />
            </div>
          </div>
        </div>
      </div>
    );
  }
}