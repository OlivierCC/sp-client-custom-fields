/**
 * @file PropertyFieldPhoneNumberHost.tsx
 * Renders the controls for PropertyFieldPhoneNumber component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldPhoneNumberPropsInternal, IPhoneNumberFormat } from './PropertyFieldPhoneNumber';
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * @interface
 * PropertyFieldPhoneNumberHost properties interface
 *
 */
export interface IPropertyFieldPhoneNumberHostProps extends IPropertyFieldPhoneNumberPropsInternal {
}

/**
 * @interface
 * Defines the masked input properties
 *
 */
interface IMaskedInputProps {
  type?: string;
  id?: string;
  placeholder?: string;
  className?: string;
  pattern?: string;
  maxLength?: string;
  title?: string;
  label?: string;
  dataCharset?: string;
  required?: boolean;
  value?: string;
  initialValue?: string;
  onChange?(newValue?: string): void;
}

/**
 * @interface
 * Defines the state of masked input control
 *
 */
interface IMaskedInputState {
  firstLoading?: boolean;
  value?: string;
}

/**
 * @interface
 * MaskedInput control.
 * This control is a fork of the input masking component available on GitHub
 * https://github.com/estelle/input-masking
 * by Estelle Weyl. & Alex Schmitz (c)
 *
 */
class MaskedInput extends React.Component<IMaskedInputProps, IMaskedInputState> {

  constructor(props: IPropertyFieldPhoneNumberHostProps) {
    super(props);
    //Binds events
    this.handleChange = this.handleChange.bind(this);
    this.handleFocus = this.handleFocus.bind(this);
    this.handleBlur = this.handleBlur.bind(this);

    //Inits default value
    this.state = {
      firstLoading: true,
      value: this.props.initialValue != null ? this.props.initialValue : ''
    };

  }

  public componentDidMount(): void {
    var e = this.refs['inputShell'];
    var event = { target: e };
    this.handleChange(event);
  }

  private handleChange(e): void {
    var previousValue = this.state.value;
    e.target.value = this.handleCurrentValue(e);
    this.state.value = e.target.value;

    if (this.state.firstLoading === true && previousValue == '')
      this.state.value = '';
    this.state.firstLoading = false;
    this.setState(this.state);

    //var maskElement = document.getElementById(this.props.id + 'Mask');
    //if (maskElement != null)
    //  maskElement.innerHTML = this.setValueOfMask(e);

    if (this.props.onChange != null)
      this.props.onChange(e.target.value);
  }

  private handleCurrentValue(e): string {
    var isCharsetPresent = e.target.getAttribute('data-charset'),
        maskedNumber = 'XMDY',
        maskedLetter = '_',
        placeholder = isCharsetPresent || e.target.getAttribute('data-placeholder'),
        value = e.target.value, l = placeholder.length, newValue = '',
        i, j, isInt, isLetter, strippedValue, matchesNumber, matchesLetter;

    // strip special characters
    strippedValue = isCharsetPresent ? value.replace(/\W/g, "") : value.replace(/\D/g, "");

    for (i = 0, j = 0; i < l; i++) {
        isInt = !isNaN(parseInt(strippedValue[j]));
        isLetter = strippedValue[j] ? strippedValue[j].match(/[A-Z]/i) : false;
        matchesNumber = (maskedNumber.indexOf(placeholder[i]) >= 0);
        matchesLetter = (maskedLetter.indexOf(placeholder[i]) >= 0);
        if ((matchesNumber && isInt) || (isCharsetPresent && matchesLetter && isLetter)) {
                newValue += strippedValue[j++];
          } else if ((!isCharsetPresent && !isInt && matchesNumber) || (isCharsetPresent && ((matchesLetter && !isLetter) || (matchesNumber && !isInt)))) {
                //this.options.onError( e ); // write your own error handling function
                return newValue;
        } else {
            newValue += placeholder[i];
        }
        // break if no characters left and the pattern is non-special character
        if (strippedValue[j] == undefined) {
          break;
        }
    }

    if (this.props['data-valid-example']) {
      return this.validateProgress(e, newValue);
    }

    return newValue;
  };

/*
  private setValueOfMask(e): string {
    var value = e.target.value,
        placeholder = e.target.getAttribute('data-placeholder');

    return "<i style=\"font-style: normal;color: transparent;opacity: 0;visibility: hidden;\">" + value + "</i>" + placeholder.substr(value.length);
  };
*/
  private validateProgress(e, value): string {
      var validExample = this.props['data-valid-example'],
          pattern = new RegExp(this.props.pattern),
          placeholder = e.target.getAttribute('data-placeholder'),
          l = value.length, testValue = '', i;

      //convert to months
      if ((l == 1) && (placeholder.toUpperCase().substr(0,2) == 'MM')) {
        if(value > 1 && value < 10) {
          value = '0' + value;
        }
        return value;
      }

    for ( i = l; i >= 0; i--) {
        testValue = value + validExample.substr(value.length);
        if (pattern.test(testValue)) {
          return value;
        } else {
          value = value.substr(0, value.length-1);
        }
    }

      return value;
  };

  private handleBlur(e): void {
      var currValue = e.target.value, pattern;

      // if value is empty, remove label parent class
      if(currValue.length == 0) {

        if(e.target.required) {
          this.updateLabelClass(e, "required", true);
          this.handleError(e, 'required');
        }

      } else {
        pattern = new RegExp('^' + this.props.pattern + '$');

        if(pattern.test(currValue)) {
          this.updateLabelClass(e, "good", true);
        } else {
          this.updateLabelClass(e, "error", true);
          this.handleError(e, 'invalidValue');
        }

      }
  };

  private handleFocus(e): void {
        this.updateLabelClass(e, 'focus', false);
  };

  private updateLabelClass(e, className, replaceExistingClass): void {
       var parentLI = e.target.parentNode.parentNode,
           pastClasses = ['error', 'required', 'focus', 'good'],
           i;

       if (replaceExistingClass) {
           for(i = 0; i < pastClasses.length; i++) {
                parentLI.classList.remove(pastClasses[i]);
           }
       }

       parentLI.classList.add(className);
  };

  private handleError(e, errorMsg): boolean {
        // the event and errorMsg name are passed. Label is already handled. What else do we do with error?
        //var possibleErrorMsgs = ['invalidValue', 'required'];
        return true;
  };

  public render(): JSX.Element {
        var props = {
                 type: (this.props && this.props.type) || '' ,
                 id: this.props.id,
                 placeholder: this.props.placeholder,
                 className: "masked " + (this.props.className || ''),
                 pattern: this.props.pattern,
                 maxLength: this.props.pattern.length,
                 title: this.props.title,
                 label: this.props.label,
                 dataCharset: this.props['data-charset'],
                 required: this.props.required,
                 initialValue: this.props.initialValue
             };

        var shellStyle = {
          position: 'relative',
          lineHeight: '1',
        };
        var shellStyleSpan = {
            position: 'absolute',
            left: '12px',
            top: '3px',
            color: '#ccc',
            pointerEvents: 'none',
            fontSize: '16px',
            fontFamily: 'monospace',
            paddingRight: '10px',
            backgroundColor: 'transparent',
            textTransform: 'uppercase'
        };
        var shellStyleSpanI = {
              fontStyle: 'normal',
              color: 'transparent',
              opacity: '0',
              visibility: 'hidden'
        };
        var inputShell = {
            fontSize: '16px',
            fontFamily: 'monospace',
            paddingRight: '10px',
            backgroundColor: 'transparent',
            textTransform: 'uppercase'
        };

      var placeHolderContent = props.placeholder.substr(this.state.value.length);

      return (
            <span style={shellStyle}>
                <span style={shellStyleSpan}
                  aria-hidden="true"
                  ref="spanMask"
                  id={props.id + 'Mask'}><i style={shellStyleSpanI}>{this.state.value}</i>{placeHolderContent}</span>
                <input style={inputShell}
                id={props.id}
                ref="inputShell"
                onChange={this.handleChange}
                onFocus={this.handleFocus}
                onBlur={this.handleBlur}
                name={props.id}
                type={props.type}
                className={props.className}
                data-placeholder={props.placeholder}
                data-pattern={props.pattern}
                aria-required={props.required}
                data-charset={props.dataCharset}
                required={props.required}
                value={this.state.value}
                title={props.title}/>
              </span>
          );
      };
}

interface IPhoneNumberFormatPattern {
  type: IPhoneNumberFormat;
  pattern: string;
  placeHolder: string;
  maxLenght: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldPhoneNumber component
 */
export default class PropertyFieldPhoneNumberHost extends React.Component<IPropertyFieldPhoneNumberHostProps, {}> {

  private patterns: IPhoneNumberFormatPattern[] = [
    { type: IPhoneNumberFormat.UnitedStates, pattern: "\(\d{3}\) \d{3}\-\d{4}", placeHolder: "(XXX) XXX-XXXX", maxLenght: '14'},
    { type: IPhoneNumberFormat.Canada, pattern: "\d{3}\-\d{3}\-\d{4}", placeHolder: "XXX-XXX-XXXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.Quebec, pattern: "\d{3} \d{3}\-\d{4}", placeHolder: "XXX XXX-XXXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.Mexico, pattern: "\d{10}", placeHolder: "XXXXXXXXXX", maxLenght: '10'},
    { type: IPhoneNumberFormat.France, pattern: "\d{2} \d{2} \d{2} \d{2} \d{2}", placeHolder: "XX XX XX XX XX", maxLenght: '14'},
    { type: IPhoneNumberFormat.Denmark, pattern: "\d{2} \d{2} \d{2} \d{2}", placeHolder: "XX XX XX XX", maxLenght: '11'},
    { type: IPhoneNumberFormat.Iceland, pattern: "\d{3}\-\d{4}", placeHolder: "XXX-XXXX", maxLenght: '8'},
    { type: IPhoneNumberFormat.NorwayLandLine, pattern: "\d{2} \d{2} \d{2} \d{2}", placeHolder: "XX XX XX XX", maxLenght: '11'},
    { type: IPhoneNumberFormat.NorwayMobile, pattern: "\d{3} \d{2} \d{3}", placeHolder: "XXX XX XXX", maxLenght: '11'},
    { type: IPhoneNumberFormat.Portugal, pattern: "\d{2} \d{2} \d{2} \d{3}", placeHolder: "XX XX XX XXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.PolandLandLine, pattern: "\d{2}\-\d{3}\-\d{2}\-\d{2}", placeHolder: "XX-XXX-XX-XX", maxLenght: '12'},
    { type: IPhoneNumberFormat.PolandMobile, pattern: "\d{3}\-\d{3}\-\d{3}", placeHolder: "XXX-XXX-XXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.Spain, pattern: "\d{2} \d{3} \d{2} \d{2}", placeHolder: "XX XXX XX XX", maxLenght: '12'},
    { type: IPhoneNumberFormat.Switzerland, pattern: "\d{3} \d{3} \d{2} \d{2}", placeHolder: "XXX XXX XX XX", maxLenght: '13'},
    { type: IPhoneNumberFormat.Turkey, pattern: "\d{4} \d{3} \d{2} \d{2}", placeHolder: "XXXX XXX XX XX", maxLenght: '14'},
    { type: IPhoneNumberFormat.Russian, pattern: "\d{4} \d{2}\-\d{2}\-\d{2}", placeHolder: "XXXX XX-XX-XX", maxLenght: '13'},
    { type: IPhoneNumberFormat.Germany, pattern: "\d{5} \d{6}", placeHolder: "XXXXX XXXXXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.BelgiumLandLine, pattern: "\d{3} \d{2} \d{2} \d{2}", placeHolder: "XXX XX XX XX", maxLenght: '12'},
    { type: IPhoneNumberFormat.BelgiumMobile, pattern: "\d{4} \d{2} \d{2} \d{2}", placeHolder: "XXXX XX XX XX", maxLenght: '13'},
    { type: IPhoneNumberFormat.UK, pattern: "\d{5} \d{6}", placeHolder: "XXXXX XXXXXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.Pakistan, pattern: "\(\d{3}\) \d{7}", placeHolder: "(XXX) XXXXXXX", maxLenght: '13'},
    { type: IPhoneNumberFormat.IndiaLandLine, pattern: "\d{3}\-\d{7}", placeHolder: "XXX-XXXXXXX", maxLenght: '11'},
    { type: IPhoneNumberFormat.IndiaMobile, pattern: "\d{5}\-\d{5}", placeHolder: "XXXXX-XXXXX", maxLenght: '11'},
    { type: IPhoneNumberFormat.ChinaLandLine, pattern: "\(\d{4}\) \d{4} \d{4}", placeHolder: "(XXXX) XXXX XXXX", maxLenght: '16'},
    { type: IPhoneNumberFormat.ChinaMobile, pattern: "\d{3} \d{4} \d{4}", placeHolder: "XXX XXXX XXXX", maxLenght: '13'},
    { type: IPhoneNumberFormat.HongKong, pattern: "\d{4} \d{4}", placeHolder: "XXXX XXXX", maxLenght: '9'},
    { type: IPhoneNumberFormat.Japan, pattern: "\(\d{3}\) \d{3}\-\d{4}", placeHolder: "(XXX) XXX-XXXX", maxLenght: '14'},
    { type: IPhoneNumberFormat.Malaysia, pattern: "\d{3}\-\d{3} \d{4}", placeHolder: "XXX-XXX XXXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.Philippines, pattern: "\(\d{3}\) \d{3} \d{4}", placeHolder: "(XXX) XXX XXXX", maxLenght: '14'},
    { type: IPhoneNumberFormat.Singapore, pattern: "\d{4} \d{4}", placeHolder: "XXXX XXXX", maxLenght: '9'},
    { type: IPhoneNumberFormat.TaiwanLandLine, pattern: "\(\d{2}\) \d{4} \d{4}", placeHolder: "(XX) XXXX XXXX", maxLenght: '14'},
    { type: IPhoneNumberFormat.TaiwanMobile, pattern: "\(\d{3}\) \d{6}", placeHolder: "(XXX) XXXXXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.SouthKoreaMobile, pattern: "\d{3}\-\d{3}\-\d{4}", placeHolder: "XXX-XXX-XXXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.Australia, pattern: "\d{2} \d{4} \d{4}", placeHolder: "XX XXXX XXXX", maxLenght: '12'},
    { type: IPhoneNumberFormat.NewZealand, pattern: "\(\d{2}\) \d{3}\-\d{4}", placeHolder: "(XX) XXX-XXXX", maxLenght: '13'},
    { type: IPhoneNumberFormat.CostaRica, pattern: "\d{4}\-\d{4}", placeHolder: "XXXX-XXXX", maxLenght: '9'},
    { type: IPhoneNumberFormat.ElSalvador, pattern: "\d{4}\-\d{4}", placeHolder: "XXXX-XXXX", maxLenght: '9'},
    { type: IPhoneNumberFormat.Guatemala, pattern: "\d{4}\-\d{4}", placeHolder: "XXXX-XXXX", maxLenght: '9'},
    { type: IPhoneNumberFormat.HondurasLandLine, pattern: "\d{3}\-\d{4}", placeHolder: "XXX-XXXX", maxLenght: '8'},
    { type: IPhoneNumberFormat.HondurasMobile, pattern: "\d{4}\-\d{4}", placeHolder: "XXXX-XXXX", maxLenght: '9'},
    { type: IPhoneNumberFormat.BrazilLandLine, pattern: "\d{4}\-\d{4}", placeHolder: "XXXX-XXXX", maxLenght: '9'},
    { type: IPhoneNumberFormat.BrazilMobile, pattern: "\d{5}\-\d{4}", placeHolder: "XXXXX-XXXX", maxLenght: '10'},
    { type: IPhoneNumberFormat.Peru, pattern: "\(\d{3}\) \d{2}\-\d{4}", placeHolder: "(XXX) XX-XXXX", maxLenght: '13'}
  ];

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldPhoneNumberHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onValueChanged = this.onValueChanged.bind(this);
  }

  /**
   * @function
   * Function called when the the text changed
   */
  private onValueChanged(element: string): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChange && element != null) {
      this.props.onPropertyChange(this.props.targetProperty, element);
    }
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    var selectedFormat: IPhoneNumberFormat = IPhoneNumberFormat.UnitedStates;
    if (this.props.phoneNumberFormat != null)
      selectedFormat = this.props.phoneNumberFormat;
    var selectedPattern: IPhoneNumberFormatPattern;
    for (var i = 0; i < this.patterns.length; i++) {
      if (this.patterns[i].type === selectedFormat)
        selectedPattern = this.patterns[i];
    }

    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>
        <MaskedInput
            id="tel"
            type="tel"
            placeholder={selectedPattern.placeHolder}
            pattern={selectedPattern.pattern}
            className="ms-TextField-field"
            maxLength={selectedPattern.maxLenght}
            onChange={this.onValueChanged}
            initialValue={this.props.initialValue}
        />
      </div>
    );
  }
}