/**
 * @file PropertyFieldDatePickerHost.tsx
 * Renders the controls for PropertyFieldDatePicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldDatePickerPropsInternal } from './PropertyFieldDatePicker';
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import * as strings from 'propertyFieldsStrings';

/**
 * @interface
 * PropertyFieldDatePickerHost properties interface
 *
 */
export interface IPropertyFieldDatePickerHostProps extends IPropertyFieldDatePickerPropsInternal {
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

/**
 * @class
 * Renders the controls for PropertyFieldDatePicker component
 */
export default class PropertyFieldDatePickerHost extends React.Component<IPropertyFieldDatePickerHostProps, {}> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldDatePickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onSelectDate = this.onSelectDate.bind(this);
  }

  /**
   * @function
   * Function called when the DatePicker Office UI Fabric component selected date changed
   */
  private onSelectDate(date: Date): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChange && date != null) {
      //Checks if a formatDate function has been defined
      if (this.props.formatDate)
        this.props.onPropertyChange(this.props.targetProperty, this.props.formatDate(date));
      else
        this.props.onPropertyChange(this.props.targetProperty, date.toDateString());
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
    var date: Date;
    if (this.props.initialDate != null && this.props.initialDate != '')
      date = new Date(this.props.initialDate);
    //Renders content
    return (
      <div>
        <DatePicker { ...this.props }  value={date} strings={dateStrings}
          isMonthPickerVisible={false} onSelectDate={this.onSelectDate} allowTextInput={false}
          formatDate={this.props.formatDate}
           />
      </div>
    );
  }
}