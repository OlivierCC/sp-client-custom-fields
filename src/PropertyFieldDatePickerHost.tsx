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
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import * as strings from 'sp-client-custom-fields/strings';

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

export interface IPropertyFieldDatePickerHostState {
  date?: string;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldDatePicker component
 */
export default class PropertyFieldDatePickerHost extends React.Component<IPropertyFieldDatePickerHostProps, IPropertyFieldDatePickerHostState> {

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldDatePickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onSelectDate = this.onSelectDate.bind(this);

    this.state = {
        date: this.props.initialDate,
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
    var dateAsString: string = '';
    if (this.props.formatDate) {
      dateAsString = this.props.formatDate(date);
    }
    else {
      dateAsString = date.toDateString();
    }
    this.state.date = dateAsString;
    this.setState(this.state);
    this.delayedValidate(dateAsString);
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
    var date: Date;
    if (this.state.date != null && this.state.date != '')
      date = new Date(this.state.date);
    //Renders content
    return (
      <div>
        <DatePicker label={this.props.label}  value={date} strings={dateStrings}
          isMonthPickerVisible={false} onSelectDate={this.onSelectDate} allowTextInput={false}
          formatDate={this.props.formatDate}
           />
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